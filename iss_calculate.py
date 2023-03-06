from inspect import Parameter
import pulp
import cplex 
import xlrd
import pandas as pd
from pulp import CPLEX_CMD

def main():
    #入力用エクセルファイルを設定
    file_name = 'input_data1.xlsx'
    #計算結果のファイル名を設定
    output_file = 'output.txt'
    last_month = 'last_month_result.xlsx'
    shop = shop_data(file_name)
    staff = staff_data(shop)
    last_shift = read_last_month(last_month,shop,staff)
    x,r,h,result = calculate(shop,staff,last_shift,0)
    if result == 'Infeasible':
        x,r,h,result = calculate(shop,staff,last_shift,result)
    opt_output(x,r,h,staff,shop,output_file)
    return

#居酒屋の営業に必要なデータなど
class shop_data:
    def __init__(self,file_name):
        try:
            self.df = pd.read_excel(file_name, sheet_name=None, index_col = 0)
        except FileNotFoundError:
            print(file_name + ' not found')
            exit()
        self.cd = self.df['constraint_data']
        self.u = int(self.cd.at['値','勤務コマ数の上限'])
        self.v = int(self.cd.at['値','勤務コマ数の下限'])
        self.s = int(self.cd.at['値','連勤の上限'])
        self.day_max, self.df_days, self.df_wd, self.df_we, self.df_s, self.df_hol, self.df_phol, \
        self.H_mon, self.H_tue, self.H_wed, self.H_thu, self.H_fri, self.H_sat, self.H_sun, self.H_hol, self.H_phol, \
        self.T, self.T_2, self.T_3 = self.make_T(self.df)

    #日付と時間帯のタプルを生成
    def make_T(self,df):
        day_max = df['constraint_data']['日数']['値']
        df_days = df['calendar'][:day_max]
        df_wd = df_days[df_days['曜日'].isin(['月曜日','火曜日','水曜日','木曜日'])]
        df_we = df_days[df_days['曜日'].isin(['金曜日','土曜日'])]
        df_s = df_days[df_days['曜日'].isin(['日曜日'])]
        df_hol = df_days[df_days['営業形態'].isin(['祝日'])]
        df_phol = df_days[df_days['営業形態'].isin(['祝前日'])]
        H_mon = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','月曜日']*10),int(df['Shop_data'].at['勤務終了時間','月曜日']*10),5)]
        H_tue = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','火曜日']*10),int(df['Shop_data'].at['勤務終了時間','火曜日']*10),5)]
        H_wed = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','水曜日']*10),int(df['Shop_data'].at['勤務終了時間','水曜日']*10),5)]
        H_thu = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','木曜日']*10),int(df['Shop_data'].at['勤務終了時間','木曜日']*10),5)]
        H_fri = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','金曜日']*10),int(df['Shop_data'].at['勤務終了時間','金曜日']*10),5)]
        H_sat = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','土曜日']*10),int(df['Shop_data'].at['勤務終了時間','土曜日']*10),5)]
        H_sun = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','日曜日']*10),int(df['Shop_data'].at['勤務終了時間','日曜日']*10),5)]
        H_hol = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','祝日']*10),int(df['Shop_data'].at['勤務終了時間','祝日']*10),5)]
        H_phol = [i / 10 for i in range(int(df['Shop_data'].at['勤務開始時間','祝前日']*10),int(df['Shop_data'].at['勤務終了時間','祝前日']*10),5)]
        T = []
        T_2 = []
        T_3 = []
        for j in df_days['日付'].values:
            if df_days.at[j-1,'営業形態'] == '祝前日':
                H = H_phol
            elif df_days.at[j-1,'営業形態'] == '祝日':
                H = H_hol
            elif df_days.at[j-1,'曜日'] == '日曜日':
                H = H_sun
            elif df_days.at[j-1,'曜日'] == '土曜日':
                H = H_sat
            elif df_days.at[j-1,'曜日'] == '金曜日':
                H = H_fri
            elif df_days.at[j-1,'曜日'] == '木曜日':
                H = H_thu
            elif df_days.at[j-1,'曜日'] == '水曜日':
                H = H_wed
            elif df_days.at[j-1,'曜日'] == '火曜日':
                H = H_tue
            elif df_days.at[j-1,'曜日'] == '月曜日':
                H = H_mon
            for k in H:
                if k == H[0]:
                    T_3.append((j,k-0.5))
                if k != H[-1]:
                    T_2.append((j,k))
                T.append((j,k))
                T_3.append((j,k))
            T_3.append((j,k+0.5))
        return day_max, df_days, df_wd, df_we, df_s, df_hol, df_phol, H_mon, H_tue, H_wed, H_thu, H_fri, H_sat, H_sun, H_hol, H_phol, T, T_2, T_3
    
    #ある日のある時間帯であるポジションに必要な人数を返す
    def take_necessary_member(self,day,p_idx,slot,position):
        ans = 0
        if self.df_days['営業形態'][day-1] != '通常営業':
            att = '営業形態'
        else:
            att = '曜日'
        for i in self.df[str(int(self.df['necessary_membar'][self.df_days[att][day-1]][slot]))+'人のパターン'].loc[p_idx].values:
            if i == position:
                ans += 1
        return ans

    #ある日の時間帯のリストを返す
    def select_H(self,day):
        if self.df_days.at[day-1,'営業形態'] == '祝前日':
            H = self.H_phol
        elif self.df_days.at[day-1,'営業形態'] == '祝日':
            H = self.H_hol
        elif self.df_days.at[day-1,'曜日'] == '日曜日':
            H = self.H_sun
        elif self.df_days.at[day-1,'曜日'] == '土曜日':
            H = self.H_sat
        elif self.df_days.at[day-1,'曜日'] == '金曜日':
            H = self.H_fri
        elif self.df_days.at[day-1,'曜日'] == '木曜日':
            H = self.H_thu
        elif self.df_days.at[day-1,'曜日'] == '水曜日':
            H = self.H_wed
        elif self.df_days.at[day-1,'曜日'] == '火曜日':
            H = self.H_tue
        elif self.df_days.at[day-1,'曜日'] == '月曜日':
            H = self.H_mon
        return H

    #日付の集合を切り取る
    def cut_T(self,days):
        ans = []
        for j,k in self.T:
            if j in days:
                ans.append((j,k))
        return ans

#スタッフに関するデータなど
class staff_data:
    def __init__(self,shop):
        self.Staffs = shop.df['Staff_names']
        self.Staff_names = dict(zip(['Staff'+str(i) for i in range(1,len(shop.df['Staff_names'].index.values)+1)],shop.df['Staff_names'].index.values))
        self.Staffs = self.Staffs.replace('〇',1)
        self.Staffs = self.Staffs.replace('×',0)
        self.res = self.Staffs[self.Staffs['責任者'] == 1].index.values
        self.leader = list(self.Staffs[self.Staffs['属性'] == '店長'].index.values).pop()
        self.Staffs.insert(self.Staffs.columns.get_loc('時給')+1,'コマ給',self.Staffs['時給']/2)
        self.Staffs['コマ給'] = self.Staffs['コマ給'].astype('int')
        self.Staffs.insert(self.Staffs.columns.get_loc('コマ給')+1,'深夜コマ給',self.Staffs['コマ給']*1.25)
        self.Staffs['深夜コマ給'] = self.Staffs['深夜コマ給'].astype('int')
        for i in shop.df['Si_Position_data'].index.values:
            si_list = []
            s = pd.Series([1]*len(self.Staffs.index.values),index=self.Staffs.index.values)
            col_name = '['
            for j in range(len(shop.df['Si_Position_data'].loc[i])):
                if shop.df['Si_Position_data'].loc[i].notnull()[j]:
                    si_list.append(shop.df['Si_Position_data'].loc[i].values[j])
            for k in si_list:
                s = s * self.Staffs[k].values
                col_name = col_name + "'" + k + "', "
            col_name = col_name[:-2] + ']'
            self.Staffs.insert(len(self.Staffs.columns.values),col_name,s)
        self.Pos = self.Staffs.iloc[:,self.Staffs.columns.get_loc('ポジション')+1:].columns.values
        self.Pos_names = dict(zip(['Position'+str(i) for i in range(1,len(self.Pos)+1)],self.Pos))
        self.dw = self.read_staff_dw(shop.df)
        self.para = self.make_parameter(shop.df)
    
    #各スタッフの休み希望を読み込む
    def read_staff_dw(self,df):
        dw = df.copy()
        for i in self.Staffs.index.values:
            dw[i] = dw[i].replace('〇',0)
            dw[i] = dw[i].replace('×',1)
        return dw

    #あるスタッフの休み希望を出している日のリストを返す
    def off_day(self,shop,staff):
        off = []
        for j in shop.df_days.index.values:
            if self.dw[staff][j+1].max() == 1:
                off.append(j)
        return off

    #スタッフの属性による重みを返す
    def make_parameter(self,df):
        para = self.Staffs['属性'].copy()
        para = para.replace('店長',1)
        para = para.replace('社員',1)
        for i in df['Att_data'].columns.values:
            para = para.replace(i,1+((df['Att_data'].at['優先度',i]-1)*2)/10)
        return para

#前の月のデータを読み込む
def read_last_month(file_name,shop,staff):
    try:
        df = pd.read_excel(file_name, sheet_name=None, index_col = 0)
    except FileNotFoundError:
        print(file_name + ' not found')
        exit()
    last = df['月間シフト'].iloc[:,-shop.s:]
    new_columns = []
    for i in range(0,len(last.columns.values)):
        new_columns.append(-i)
    new_columns.reverse()
    last.columns = new_columns
    last_y = last.copy()
    last_y[last_y.notnull()] = 1
    last_y = last_y.fillna(0)
    for i in staff.Staffs.index.values:
        if i not in last_y.index.values:
            last_y.loc[i] = 0
    return last_y

#計算結果のファイルのファイルを生成
def opt_output(x,r,h,staff,shop,file_name):
    N = list(staff.Staff_names.keys())
    D = list(shop.df_days['日付'].values)
    T = shop.T
    P = list(staff.Pos_names.keys())
    f = open(file_name,'w',encoding='utf-8')
    f.write('---スタッフ---\n')
    N_2 = [staff.Staff_names[i]+',' for i in N] + ['不足(責任者)\n']
    f.writelines(N_2)
    f.write('\n---日数---\n')
    str_D = [str(d)+'\n' if d == D[-1] else str(d)+',' for d in D]
    f.writelines(str_D)
    max_H, min_H = max_min_H(T)
    f.write('\n---オープン---\n')
    f.write(str(min_H)+'\n')
    f.write('\n---クローズ---\n')
    f.write(str(max_H)+'\n')
    f.write('\n---各スタッフの給与---\n')
    c_list = [s+':'+str(c)+'\n' for s,c in staff.Staffs['時給'].to_dict().items()] + ['不足(責任者):0\n']
    f.writelines(c_list)
    f.write('\n---各スタッフの希望総給与---\n')
    d_list = [s+':'+str(int(d))+'\n' for s,d in staff.Staffs['希望総給与'].to_dict().items()] + ['不足(責任者):0\n']
    f.writelines(d_list)
    f.write('\n---ポジション---\n')
    P_2 = [staff.Pos_names[l]+'\n'  if l == P[-1] else staff.Pos_names[l]+';' for l in P]
    f.writelines(P_2)
    f.write('\n---解---\n')
    for i in N:
        for j,k in T:
            for l in P:
                if round(x[i][j,k][l].value()) == 1:
                    f.write(staff.Staff_names[i]+';'+str(j)+';'+str(k)+';'+staff.Pos_names[l]+'\n')
    for j,k in T:
        for l in P:
            if round(r[j,k][l].value()) == 1:
                f.write('不足(責任者);'+str(j)+';'+str(k)+';'+staff.Pos_names[l]+'\n')
            if round(h[j,k][l].value()) != 0 and h[j,k][l].value() != None:
                for m in range(round(h[j,k][l].value())):
                    f.write('不足;'+str(j)+';'+str(k)+';'+staff.Pos_names[l]+'\n')
    f.close()
    return

#日と時間帯のタプルから最も早い時間帯と最も遅い時間帯を返す
def max_min_H(T):
    max_H = 0
    min_H = float('inf')
    for j,k in T:
        if max_H < k:
            max_H = k
        if min_H > k:
            min_H = k
    return max_H, min_H

#変数zの引数を作成(日、時間帯、パターンのタプル)
def make_z_para(T,shop):
    ans = []
    for j,k in T:
        if shop.df_days['営業形態'][j-1] != '通常営業':
            att = '営業形態'
        else:
            att = '曜日'
        for i in range(len(shop.df[str(int(shop.df['necessary_membar'][shop.df_days[att][j-1]][k]))+'人のパターン'].index)):
            ans.append((j,k,i))
    return ans

#辞書の値からキー値を返す
def get_keys(d, val):
    return [k for k, v in d.items() if v == val][0]

#計算を行う
def calculate(shop,staff,last_shift,sta):
    N = list(staff.Staff_names.keys())
    T = shop.T
    T_2 = shop.T_2
    T_3 = shop.T_3
    P = list(staff.Pos_names.keys())
    D = list(shop.df_days['日付'].values)
    para = staff.para
    max_H, min_H = max_min_H(T)
    M = int(len(P)*(max_H-min_H)*2+1)
    h_cost = staff.Staffs['深夜コマ給'].max() * shop.df['parameter'].at['重み','不足コマ数']
    
    prob = pulp.LpProblem(sense = pulp.LpMinimize)

    #変数
    all_var = 0

    #スタッフiが日jの時間帯kでポジションlに勤務するとき1,そうでなければ0
    x = pulp.LpVariable.dicts('x',(N,T_3,P),cat = pulp.LpBinary)
    len_x = len(N)*len(T_3)*len(P)
    print('変数xの数={}'.format(len_x))
    all_var += len_x

    #スタッフiが日jで勤務するとき1,そうでなければ0
    y = pulp.LpVariable.dicts('y',(N,list(range(-shop.s+1,1))+D),cat = pulp.LpBinary)
    len_y = len(N)*len(list(range(-shop.s+1,1))+D)
    print('変数yの数={}'.format(len_y))
    all_var += len_y

    #j日目の時間帯kでポジションのパターンpが選ばれるとき1,そうでなければ0
    pat = make_z_para(T,shop)
    z = pulp.LpVariable.dicts('z',(pat),cat = pulp.LpBinary)
    len_z = len(pat) 
    print('変数zの数={}'.format(len_z))
    all_var += len_z

    #スタッフiのj日目の時間帯kがその日の1番最初の勤務であるとき1,そうでなければ0
    S = pulp.LpVariable.dicts('S',(N,T),cat = pulp.LpBinary)
    len_S = len(N)*len(T)
    print('変数Sの数={}'.format(len_S))
    all_var += len_S

    #スタッフiのj日目の時間帯kがその日の1番最後の勤務であるとき1,そうでなければ0
    S2 = pulp.LpVariable.dicts('S2',(N,T),cat = pulp.LpBinary)
    len_S2 = len(N)*len(T)
    print('変数S2の数={}'.format(len_S2))
    all_var += len_S2

    #スタッフのポジション移動に関するペナルティ
    abso = pulp.LpVariable.dicts('abso',(N,T_2,P),lowBound = 0, cat = pulp.LpInteger)
    len_abso = len(N)*len(T_2)*len(P)
    print('変数absoの数={}'.format(len_abso))
    all_var += len_abso

    #各スタッフの給与の目安からのずれに関するペナルティ
    p1 = pulp.LpVariable.dicts('p1',N,lowBound = 0, cat = pulp.LpInteger)
    len_p1 = len(N)
    print('変数p1の数={}'.format(len_p1))
    all_var += len_p1

    #責任者の不足
    r = pulp.LpVariable.dicts('r',(T,P),cat = pulp.LpBinary)
    len_r = len(T)*len(P)
    print('変数rの数={}'.format(len_r))
    all_var += len_r

    #責任者以外の不足
    h = pulp.LpVariable.dicts('h',(T,P),lowBound = 0, cat = pulp.LpInteger)
    len_h = len(T)*len(P)
    print('変数hの数={}'.format(len_h))
    all_var += len_h

    print('全ての変数={}'.format(all_var))

    #目的関数
    sum_day_cost = pulp.lpSum(shop.df['parameter'].at['重み','総給料'] * staff.Staffs['コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k < 22 else 0 for i in N for j,k in T for l in P)
    sum_night_cost = pulp.lpSum(shop.df['parameter'].at['重み','総給料'] * staff.Staffs['深夜コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k >= 22 else 0 for i in N for j,k in T for l in P)
    job_change = pulp.lpSum(shop.df['parameter'].at['重み','ポジション移動'] * abso[i][j,k][l] for i in N for j,k in T_2 for l in P)
    penalty1 = pulp.lpSum(shop.df['parameter'].at['重み','希望総給料とのずれ'] * para[staff.Staff_names[i]] * p1[i] for i in N) #希望総給与のずれに関するペナルティ
    re_help = pulp.lpSum(h_cost * r[j,k][l] for j,k in T for l in P)
    h_help = pulp.lpSum(h_cost * h[j,k][l] for j,k in T for l in P)
    obj = sum_day_cost + sum_night_cost + job_change + penalty1 + re_help + h_help
    prob.setObjective(obj)

    count_con = 0

    #前の月の最後のシフトを設定
    count_con_p = 0
    for i in N:
        for j in list(range(-shop.s+1,1)):
            prob += y[i][j] == int(last_shift[j][staff.Staff_names[i]])
            count_con_p += 1
    print('前の月のシフトを設定={}'.format(count_con_p))
    count_con += count_con_p

    #ダミーの時間帯を設定
    count_con_d = 0
    for i in N:
        for j in D:
            for l in P:
                prob += x[i][j,shop.select_H(j)[0]-0.5][l] == 0
                prob += x[i][j,shop.select_H(j)[-1]+0.5][l] == 0
                count_con_d += 2
    print('制約dの数{}'.format(count_con_d))
    count_con += count_con_d

    #constraint_st4(休み希望を設定)
    count_con_st4 = 0
    for i in N:
        for j,k in T:
            if staff.dw[staff.Staff_names[i]][j][k] == 1:
                for l in P:
                    prob += x[i][j,k][l] == 0
                    count_con_st4 += 1
    print('制約st4の数{}'.format(count_con_st4))
    count_con += count_con_st4

    #constraint_sh3(準備は必ず最初の30分のみ、他のポジションは最初の30分勤務できない)
    count_con_sh3 = 0
    for i in N:
        for j in D:
            H = shop.select_H(j)
            for k in H:
                if k == H[0]:
                    for l in P:
                        if staff.Pos_names[l] != '準備':
                            prob += x[i][j,k][l] == 0
                            prob += h[j,k][l] == 0
                            prob += r[j,k][l] == 0
                            count_con_sh3 += 1
                else:
                    prob += x[i][j,k][get_keys(staff.Pos_names, '準備')] == 0
                    prob += h[j,k][get_keys(staff.Pos_names, '準備')] == 0
                    prob += r[j,k][get_keys(staff.Pos_names, '準備')] == 0
                    count_con_sh3 += 1
    print('制約sh3の数={}'.format(count_con_sh3))
    count_con += count_con_sh3

    #constraint_st3(不可能ポジションを設定)
    count_con_st3 = 0
    Pn = staff.Staffs.iloc[:,staff.Staffs.columns.get_loc('ポジション')+1:]
    for i in N:
        for l in P:
            if Pn[staff.Pos_names[l]][staff.Staff_names[i]] == 0:
                for j,k in T:
                    prob += x[i][j,k][l] == 0
                    count_con_st3 += 1
    print('制約st3の数={}'.format(count_con_st3))
    count_con += count_con_st3

    #constraint_sh6(ポジション移動に関するペナルティ)
    count_con_sh6 = 0
    for i in N:
        for j,k in T_2:
            for l in P:
                prob += x[i][j,k][l] - x[i][j,k+0.5][l] + abso[i][j,k][l] >= 0
                prob += x[i][j,k][l] - x[i][j,k+0.5][l] - abso[i][j,k][l] <= 0
                count_con_sh6 += 2
    print('絶対値の設定={}'.format(count_con_sh6))
    count_con += count_con_sh6
    
    #constraint_sh2(各日の各時間帯で責任者が必要)
    count_con_sh2 = 0
    for j,k in T:
        prob += pulp.lpSum(x[get_keys(staff.Staff_names, i)][j,k][l] for i in staff.res for l in P) + pulp.lpSum(r[j,k][l] for l in P) >= 1
        count_con_sh2 += 1
    print('制約sh2の数={}'.format(count_con_sh2))
    count_con += count_con_sh2

    #constraint_sh1(金土日、祝日、祝前日は店長はなるべく勤務しなければならい)
    if sta != 'Infeasible':
        count_con_sh1 = 0
        for j in list(set(shop.df_we.index.values)|set(shop.df_s.index.values)|set(shop.df_hol.index.values)|set(shop.df_phol.index.values)):
            prob += y[get_keys(staff.Staff_names, staff.leader)][j] == 1
            count_con_sh1 += 1
        print('制約sh1の数={}'.format(count_con_sh1))
        count_con += count_con_sh1
    else:
        count_con_sh1 = 0
        for j in list((set(shop.df_we.index.values)|set(shop.df_s.index.values)|set(shop.df_hol.index.values)|set(shop.df_phol.index.values))- set(staff.off_day(shop,staff.leader))):
            prob += y[get_keys(staff.Staff_names, staff.leader)][j] == 1
            count_con_sh1 += 1
        print('制約sh1の数={}'.format(count_con_sh1))
        count_con += count_con_sh1

    #constraint_st2(各スタッフの連続勤務日数の上限を設定(yの値も設定))
    count_con_st2 = 0
    for i in N:
        for j in D:
            prob += pulp.lpSum(x[i][j,k][l] for k in shop.select_H(j) for l in P) - M * y[i][j] <= 0
            prob += pulp.lpSum(x[i][j,k][l] for k in shop.select_H(j) for l in P) - y[i][j] >= 0
            count_con_st2 += 2

    for i in N:
        for j2 in D:
            prob += pulp.lpSum(y[i][j] for j in range(j2-shop.s,j2+1)) <= shop.s
            count_con_st2 += 1
    print('制約st2の数={}'.format(count_con_st2))
    count_con += count_con_st2

    #constraint_st1(各スタッフの各日の総勤務時間数の上下限を設定)
    count_con_st1 = 0
    for i in N:
        for j in D:
            prob += pulp.lpSum(x[i][j,k][l] for k in shop.select_H(j) for l in P) - shop.u * y[i][j] <= 0
            prob += pulp.lpSum(x[i][j,k][l] for k in shop.select_H(j) for l in P) - shop.v * y[i][j] >= 0
            count_con_st1 += 2
    print('制約st1の数={}'.format(count_con_st1))
    count_con += count_con_st1

    #constraint_sh4(各時間帯ごとに必要な人数とポジションを設定)
    count_con_sh4 = 0

    for j,k in T:
        if shop.df_days['営業形態'][j-1] != '通常営業':
            att = '営業形態'
        else:
            att = '曜日'
        for p in range(len(shop.df[str(int(shop.df['necessary_membar'][shop.df_days[att][j-1]][k]))+'人のパターン'].index)):
            for l in P:
                prob += pulp.lpSum(x[i][j,k][l] for i in N) + h[j,k][l] -shop.take_necessary_member(j,p,k,staff.Pos_names[l]) * z[j,k,p] >= 0
                count_con_sh4 += 1
        prob += pulp.lpSum(z[j,k,q] for q in range(len(shop.df[str(int(shop.df['necessary_membar'][shop.df_days[att][j-1]][k]))+'人のパターン'].index))) == 1
        count_con_sh4 += 1
    print('制約sh4の数={}'.format(count_con_sh4))
    count_con += count_con_sh4
    
    #constraint_st6(各スタッフの希望総給与に関するペナルティ)
    count_con_st6 = 0
    for i in N:
        prob += pulp.lpSum(staff.Staffs['コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k < 22 else 0 for j,k in T for l in P) + pulp.lpSum(staff.Staffs['深夜コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k >= 22 else 0  for j,k in T for l in P) + p1[i] >= staff.Staffs['希望総給与'][staff.Staff_names[i]]
        prob += pulp.lpSum(staff.Staffs['コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k < 22 else 0 for j,k in T for l in P) + pulp.lpSum(staff.Staffs['深夜コマ給'][staff.Staff_names[i]] * x[i][j,k][l] if k >= 22 else 0  for j,k in T for l in P) - p1[i] <= staff.Staffs['希望総給与'][staff.Staff_names[i]]
        count_con_st6 += 2
    print('制約6の数={}'.format(count_con_st6))
    count_con += count_con_st6

    #constraint_x(各スタッフは各日の各時間帯でたかだか1つのポジションしか勤務できない)
    count_con_x = 0
    for i in N:
        for j,k in T:
            prob += pulp.lpSum(x[i][j,k][l] for l in P) <= 1
            count_con_x += 1
    print('制約xの数={}'.format(count_con_x))
    count_con += count_con_x

    #constraint_sh5(勤務に入る日は必ず連続とならなければならない)
    count_con_sh5 = 0
    for i in N:
        for j in D:
            prob += pulp.lpSum(S[i][j,k] for k in shop.select_H(j)) <= y[i][j]
            prob += pulp.lpSum(S2[i][j,k] for k in shop.select_H(j)) <= y[i][j]
            count_con_sh5 += 2
    
    for i in N:
        for j,k in T:
            prob += pulp.lpSum(x[i][j,k][l] for l in P) - pulp.lpSum(x[i][j,k-0.5][l] for l in P) - S[i][j,k] <= 0
            prob += pulp.lpSum(x[i][j,k][l] for l in P) - pulp.lpSum(x[i][j,k+0.5][l] for l in P) - S2[i][j,k] <= 0
            count_con_sh5 += 2

    for i in N:
        for j,k in T:
            prob += pulp.lpSum(S[i][j,k2] for k2 in shop.select_H(j)[0:shop.select_H(j).index(k)+1]) + pulp.lpSum(S2[i][j,k2] for k2 in shop.select_H(j)[shop.select_H(j).index(k):]) - pulp.lpSum(x[i][j,k][l] for l in P) - y[i][j] == 0
            count_con_sh5 += 1

    print('制約sh5の数={}'.format(count_con_sh5))
    count_con += count_con_sh5

    #constraint_st7(1週間の総勤務時間数を設定)
    count_con_st7 = 0
    for i in N:
        for j in shop.df_s.index.values:
            if j == shop.df_s.index.values[-1]:
                week = list(range(j,shop.day_max+1))
            else:
                week = list(range(j,j+7))
            prob += pulp.lpSum(x[i][j,k][l] for j,k in shop.cut_T(week) for l in P) <= 47.5 * 2
            count_con_st7 += 1
    print('制約st7の数={}'.format(count_con_st7))
    count_con += count_con_st7

    print('全ての制約の数={}'.format(count_con))

    #timelimitで計算時間の上限(秒数)を設定,最適性ギャップがgapRelの値以下になったら計算を終了
    status = prob.solve(pulp.CPLEX_CMD(keepFiles = 1, timelimit = 3600 * 1, gapRel = 0.3, options = ['set emphasis mip 4']))

    return x,r,h,pulp.LpStatus[status]

if __name__ == '__main__':
    main()