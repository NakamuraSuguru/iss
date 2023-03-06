from select import select
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import re
import openpyxl
import xlwt
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import xlsxwriter
import datetime
import itertools
import copy
import numpy as np
import calendar
import jpholiday

#データフレーム内に値valがあるか探す
def existence(dataframe,val):
    for i in list(dataframe.columns):
        for j in dataframe[i]:
            if j == val:
                return True
    return False

#同時に行えるポジションの文字列を作成
def make_text(select):
    d = {}
    for i in select:
        if type(i) == list:
            i = 'と'.join(i)
        if i in list(d.keys()):
            d[i] += 1
        else:
            d[i] = 1
    count=0
    for j,k in d.items():
        if count==0:
            t = str(j) + ':' + str(k) + '人'
            count+=1
        else:
            t = t + ',' + str(j) + ':' + str(k) + '人'
    return t

#ポジションのパターンを生成
def make_pattern3(position,position_num,si_position,num):
    pro = list(itertools.combinations_with_replacement(position+si_position,num))
    pro2 = copy.copy(pro)
    for i in pro:
        l = []
        for j in i:
            if type(j) == list:
                for k in j:
                    l.append(k)
            else:
                l.append(j)
        s = set(l)
        if set(position) > s:
            pro2.remove(i)
        else:
            for m in position:
                if l.count(m) > position_num[position.index(m)]:
                    pro2.remove(i)
                    break
    return pro2

#エクセルファイルを開く
def open_excel():
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    return output,writer

#エクセルファイルを追加する
def add_excel(writer,df,name):
    df.to_excel(writer, sheet_name=name)
    return 

#エクセルファイルを閉じる
def close_excel(output,writer):
    writer.save()
    data = output.getvalue()
    return data

#--:--を--or--.5に変換
def change(time):
    cut1 = int(time[:time.index(':')])
    cut2 = time[time.index(':')+1:]
    if cut2 == '30':
        ans = cut1 + 0.5
    else:
        ans = cut1
    return ans

#changeの逆
def change2(time):
    cut1 = int(time)
    if time-int(time) == 0.5:
        ans = str(cut1)+':30'
    else:
        ans = str(cut1)+':00'
    return ans

#ポジションのデータを削除
def delete_posi(select_posi):
    for i in select_posi:
        del st.session_state['position_num'][st.session_state['position'].index(i['ポジション名'])]
        st.session_state['position'].remove(i['ポジション名'])

#同時に行えるポジションのデータを削除
def delete_si_posi(select_si_posi):
    for i in select_si_posi:
        st.session_state['si_position'].remove([j for j in list(i.values())[2:] if j != ''])

#nanがあるかを探す
def find_nan(my_str):
    if my_str==my_str:
        return 1
    return np.nan

def main():
    st.title('店舗データ入力')
    if 'opening_time' not in st.session_state:
        st.session_state['opening_time'] = ['0'] * 9
    if 'pre_opening_time' not in st.session_state:
        st.session_state['pre_opening_time'] = ['0'] * 9
    if 'closed_time' not in st.session_state:
        st.session_state['closed_time'] = ['0'] * 9
    if 'after_closed_time' not in st.session_state:
        st.session_state['after_closed_time'] = ['0'] * 9
    if 'dataframe' not in st.session_state:
        st.session_state['dataframe'] = {}
    if 'position' not in st.session_state:
        st.session_state['position'] = []
    if 'position_num' not in st.session_state:
        st.session_state['position_num'] = []
    if 'si_position' not in st.session_state:
        st.session_state['si_position'] = []
    if 'att' not in st.session_state:
        st.session_state['att'] = ['フリーター(専業)','フリーター(兼業)','学生(専業)','学生(兼業)']
    if 'para_att' not in st.session_state:
        st.session_state['para_att'] = [1,1,1,1]
    if 'pattern' not in st.session_state:
        st.session_state['pattern'] = {}
    if 'count' not in st.session_state:
        st.session_state['count'] = 0
    if 'staffs' not in st.session_state:
        st.session_state['staffs'] = []
    if 'check_pre_val' not in st.session_state:
        st.session_state['check_pre_val'] = False

    st.subheader('店舗データアップロード')
    upload_file = st.file_uploader("店舗の基本データをアップロードしてください。", type='xlsx')
    if upload_file is not None:
        if st.session_state['count']==0:
            df_s = pd.read_excel(upload_file, header=0, sheet_name=None, index_col=0)
            st.session_state['opening_time'] = [change2(i) for i in df_s['Shop_data'].loc['開店時間'].to_list()]
            st.session_state['pre_opening_time'] = [change2(i) for i in df_s['Shop_data'].loc['勤務開始時間'].to_list()]
            st.session_state['closed_time'] = [change2(i) for i in df_s['Shop_data'].loc['閉店時間'].to_list()]
            st.session_state['after_closed_time'] = [change2(i) for i in df_s['Shop_data'].loc['勤務終了時間'].to_list()]
            st.session_state['position'] = list(df_s['Position_data'].columns)
            if '準備' in st.session_state['position']:
                st.session_state['check_pre_val'] = True
            else:
                st.session_state['check_pre_val'] = False
            st.session_state['position_num'] = df_s['Position_data'].loc['最大の人数'].to_list()
            st.session_state['si_position'] = []
            for i in df_s['Si_Position_data'].index:
                st.session_state['si_position'].append([j for j in df_s['Si_Position_data'].loc[i].to_list() if not np.isnan(find_nan(j))])
            st.session_state['att'] = df_s['Att_data'].columns.to_list()
            st.session_state['para_att'] = df_s['Att_data'].loc['優先度'].to_list()
            st.session_state['con'] = df_s['constraint_data'].loc['値'].to_list()
            st.session_state['parameter'] = df_s['parameter'].loc['重み'].to_list()
            st.session_state['dataframe']['Staff_names'] = df_s['Staff_names']
            st.session_state['staffs'] = list(df_s['Staff_names'].index)
            st.session_state['calendar'] = df_s['calendar']
            st.session_state['dataframe']['necessary_membar'] = df_s['necessary_membar']
            st.session_state['dataframe']['constraint_data'] = df_s['constraint_data']
            st.session_state['dataframe']['parameter'] = df_s['parameter']
            for p in df_s.keys():
                if '人のパターン' in p:
                    num = int(re.search(r'\d+', p).group())
                    st.session_state['pattern'][num] = df_s[p]
            st.session_state['count']=1
    else:
        st.session_state['count']=0

    week = ['月曜日','火曜日','水曜日','木曜日','金曜日','土曜日','日曜日','祝日','祝前日']

    #-----------------------------------

    st.subheader('勤務表作成期間の入力')

    col = st.columns(2)

    with col[0]:
        year_input = st.selectbox('勤務表を作成したい年を選んで下さい。',[2000+i for i in range(100)],format_func=lambda x: str(x) + '年')
    with col[1]:
        month_input = st.selectbox('勤務表を作成したい月を選んで下さい。',[i for i in range(1,13)],format_func=lambda x: str(x) + '月')

    if 'year_input' not in st.session_state:
        st.session_state['year_input'] = year_input
    if 'month_input' not in st.session_state:
        st.session_state['month_input'] = month_input

    if st.button('決定',key=1):
        st.session_state['year_input'] = year_input
        st.session_state['month_input'] = month_input
    
    c = calendar.Calendar(firstweekday=0)

    #数字のみのカレンダー
    cal = [i for i in list(c.itermonthdays(st.session_state['year_input'],st.session_state['month_input'])) if i != 0]

    #曜日込みのカレンダー
    cal_2 = []

    for i in range(len(c.monthdays2calendar(st.session_state['year_input'],st.session_state['month_input']))):
        cal_2 = cal_2 + c.monthdays2calendar(st.session_state['year_input'],st.session_state['month_input'])[i]

    week2 = ['月曜日','火曜日','水曜日','木曜日','金曜日','土曜日','日曜日']

    cal_2 = [(j,k) for j,k in cal_2 if j != 0]

    #曜日のみのカレンダー
    cal_3 = [week2[k] for j,k in cal_2]

    holidays = list(map(lambda d: d[0].day, jpholiday.month_holidays(st.session_state['year_input'],st.session_state['month_input'])))
    list_hol = []
    for i in cal:
        if i + 1 in holidays:
            list_hol.append('祝前日')
        elif i in holidays:
            list_hol.append('祝日')
        else:
            list_hol.append('通常営業')
    
    df_cal = pd.DataFrame(cal,columns=['日付'],index=None)
    sr = pd.Series(cal_3,index=None,name='曜日')
    sr2 = pd.Series(list_hol, index=None, name='営業形態')
    df_cal = pd.concat([df_cal, sr], axis=1, sort=False)
    df_cal = pd.concat([df_cal, sr2], axis=1, sort=False)

    st.write('特殊な営業形態がある場合は下の表から入力してください。')

    #表の生成
    gb = GridOptionsBuilder.from_dataframe(df_cal, editable=True)

    gb.configure_columns(['日付','曜日'],
        editable=False
    )

    gb.configure_column('営業形態',
        cellEditor='agRichSelectCellEditor',
        cellEditorParams={'values':['通常営業','祝前日','祝日']},
        cellEditorPopup=True
    )

    gb.configure_grid_options(enableRangeSelection=True)
    response = AgGrid(
        df_cal,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        updateMode=GridUpdateMode.VALUE_CHANGED
    )

    df_cal = response['data']
    st.session_state['dataframe']['calendar'] = df_cal

    #-------------------------------------------------

    st.subheader('営業時間の入力')

    st.write('曜日を選択してその曜日の開店時間、閉店時間、準備の時間、締めの時間を選択し決定ボタンを押して下さい。曜日は複数選択可能です。(深夜24時以降は、25時のようにして下さい。)')
    
    col9 = st.columns(9)
    
    check_week = dict(zip(week,[False]*9))
    for i in week:
        check_week[i] = col9[week.index(i)].checkbox(i)

    time_set = [change2(i/10) for i in range(0,300,5)]
    col2 = st.columns(2)

    with col2[0]:
        opening_time = st.selectbox('開店時間を入力して下さい。',time_set)
        closed_time = st.selectbox('閉店時間を入力して下さい。',time_set)
    with col2[1]:
        pre_time = st.selectbox('開店準備の時間を入力して下さい。',[i/2 for i in range(0,4)],format_func=lambda x: '営業開始'+str(int(x*60))+'分前')
        after_time = st.selectbox('締めの時間を入力して下さい。',[i/2 for i in range(0,4)],format_func=lambda x: '営業終了'+str(int(x*60))+'分後')

    if st.button('決定',key=11):
        for i,j in check_week.items():
            if j:
                st.session_state['opening_time'][week.index(i)] = opening_time
                pre_opening_time = change2(change(opening_time)-pre_time)
                st.session_state['pre_opening_time'][week.index(i)] = pre_opening_time
                st.session_state['closed_time'][week.index(i)] = closed_time
                after_closed_time = change2(change(closed_time)+after_time)
                st.session_state['after_closed_time'][week.index(i)] = after_closed_time

    basetime = datetime.time(00, 00, 00)

    st.write('下の表をクリックすることでも時間を変更できます。')
    df_shop = pd.DataFrame([st.session_state['opening_time'],st.session_state['pre_opening_time'],st.session_state['closed_time'],st.session_state['after_closed_time']],
                            columns=week)
    
    df_shop.insert(0,' ',['開店時間','勤務開始時間','閉店時間','勤務終了時間'])

    #表の生成
    gb = GridOptionsBuilder.from_dataframe(df_shop, editable=True)

    gb.configure_columns(' ',
        editable=False
    )

    gb.configure_columns(week,
        cellEditor='agRichSelectCellEditor',
        cellEditorParams={'values':time_set},
        cellEditorPopup=True
    )

    gb.configure_grid_options(enableRangeSelection=True)
    response = AgGrid(
        df_shop,
        gridOptions=gb.build(),
        allow_unsafe_jscode=True,
        enable_enterprise_modules=True,
        height=150,
        updateMode=GridUpdateMode.VALUE_CHANGED
    )

    df_shop = response['data'].set_index(' ')
    if not existence(df_shop,'0'):
        df_shop = pd.DataFrame([[change(df_shop.loc['開店時間'][i]) for i in week],[change(df_shop.loc['勤務開始時間'][i]) for i in week],[change(df_shop.loc['閉店時間'][i]) for i in week],[change(df_shop.loc['勤務終了時間'][i]) for i in week]],
                                index=['開店時間','勤務開始時間','閉店時間','勤務終了時間'],
                                columns=week)
        st.session_state['dataframe']['Shop_data'] = df_shop
    
        st.session_state['opening_time'] = [change2(i) for i in list(df_shop.loc['開店時間'])]
        st.session_state['pre_opening_time'] = [change2(i) for i in list(df_shop.loc['勤務開始時間'])]
        st.session_state['closed_time'] = [change2(i) for i in list(df_shop.loc['閉店時間'])]
        st.session_state['after_closed_time'] = [change2(i) for i in list(df_shop.loc['勤務終了時間'])]
    
    #---------------------

    st.write('営業時間を入力したら下のボタンからシフト希望提出用シートをダウンロードしてスタッフに配って下さい。店長、社員(固定給のスタッフ)は希望総給与を0として下さい。')

    time_index = ['休み']+['\''+str(change2(i/10)) for i in range(int(min(df_shop.loc['勤務開始時間'].to_list())*10),int(max(df_shop.loc['勤務終了時間'].to_list())*10),5)]
    time_index2 = ['休み']+['\''+str(change2(i/10+0.5)) for i in range(int(min(df_shop.loc['勤務開始時間'].to_list())*10),int(max(df_shop.loc['勤務終了時間'].to_list())*10),5)]

    output = BytesIO()

    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    ws = wb.add_worksheet(name="シフト希望")

    line = wb.add_format({'border': 1})

    ws.write('A1','氏名',line)
    ws.write('B1',' ',line)
    ws.write('A3','希望総給料',line)
    ws.write('B3',' ',line)
    ws.write('C3','円',line)
    ws.write('A5','希望シフト',line)
    ws.write_row('A6',['日付','曜日'],line)
    ws.merge_range('C6:E6', '勤務可能時間',line)
    ws.write_column('A7',cal,line)
    ws.write_column('B7',cal_3,line)
    ws.write_column('D7',['~']*len(cal),line)
    ws.write_column('G7',week2,line)
    ws.write_column('I7',['～']*7,line)
    ws.merge_range('G6:J6', '曜日毎',line)
    ws.write_column('H7',[' ']*7,line)
    ws.write_column('J7',[' ']*7,line)
    ran1 = 'C7:C'+str(6+len(cal))
    ran2 = 'E7:E'+str(6+len(cal))
    ws.data_validation(ran1, {'validate': 'list',
                              'source': time_index})
    ws.data_validation(ran2, {'validate': 'list',
                              'source': time_index2})
    ws.data_validation('H7:H13', {'validate': 'list',
                              'source': time_index})
    ws.data_validation('J7:J13', {'validate': 'list',
                              'source': time_index2})
    for i in range(len(cal)):
        ran_c = 'C'+str(7+i)
        ran_e = 'E'+str(7+i)
        if df_cal['営業形態'][i] == '休み':
            ws.write(ran_c,'休み',line)
            ws.write(ran_e,'休み',line)
        else:
            fanc_c = '=VLOOKUP(B'+str(7+i)+',G:H,2,0)'
            fanc_e = '=VLOOKUP(B'+str(7+i)+',G:J,4,0)'
            ws.write(ran_c,fanc_c,line)
            ws.write(ran_e,fanc_e,line)
        
    wb.close()

    st.download_button(
        label="シフト希望提出用シートをダウンロード",
        data=output.getvalue(),
        file_name="submit_shift.xlsx",
        mime="application/vnd.ms-excel"
    )
    #---------------------

    st.subheader('ポジションデータ入力')

    col2_2 = st.columns(2)
    with col2_2[0]:
        position_input = st.text_input('営業に必要なポジションを入力して下さい。')
    with col2_2[1]:
        position_up = st.number_input(str(position_input)+'に勤務可能な最大の人数を入力して下さい。',min_value=1)

    if st.button('追加',key=21):
        st.session_state['position'].append(position_input)
        st.session_state['position_num'].append(position_up)

    check_pre = st.checkbox('準備をポジション関係なく行う場合はチェックを入れて下さい。',value=st.session_state['check_pre_val'])
    if check_pre == True:
        if '準備' not in st.session_state['position']:
            st.session_state['position'].append('準備')
            st.session_state['position_num'].append(1)
    else:
        if '準備' in st.session_state['position']:
            st.session_state['position_num'].remove(st.session_state['position_num'][st.session_state['position'].index('準備')])
            st.session_state['position'].remove('準備')
    
    if len(st.session_state['position']) != 0:
        df_position = pd.DataFrame({'ポジション名':st.session_state['position'],
                                    '最大の人数':st.session_state['position_num']})
        
        #表の生成
        gb2 = GridOptionsBuilder.from_dataframe(df_position, editable =True)

        gb2.configure_selection(selection_mode="multiple", use_checkbox=True)

        position = AgGrid(df_position,
                            gridOptions=gb2.build(),
                            fit_columns_on_grid_load=True,
                            height=150,
                            update_mode=GridUpdateMode.VALUE_CHANGED|GridUpdateMode.SELECTION_CHANGED)
    
        st.session_state['position'] = list(position['data']['ポジション名'])

        st.session_state['select_posi'] = position['selected_rows']

        st.button('チェックしたポジションを削除',on_click=delete_posi,args = (st.session_state['select_posi'],),key=22)

        df_position = position['data'].set_index('ポジション名').T

        st.session_state['dataframe']['Position_data'] = df_position
    #---------------------------------

    si_input = st.multiselect('同時に行うことが可能なポジションを選択して下さい。',[i for i in st.session_state['position'] if i != '準備'])
    if st.button('追加',key=31):
        st.session_state['si_position'].append(si_input)

    if len(st.session_state['si_position']) != 0:
        si_select = [' ']+['ポジション'+str(i) for i in range(1,len([i for i in st.session_state['position'] if i != '準備'])+1)]
        ini = []
        for i in st.session_state['si_position']:
            ini.append(['同時に行えるポジション'+str(st.session_state['si_position'].index(i)+1)]+i+['']*(len([i for i in st.session_state['position'] if i != '準備'])-len(i)))
        
        df_siposition = pd.DataFrame(ini,columns=si_select)
        
        #表の生成
        gb3 = GridOptionsBuilder.from_dataframe(df_siposition, editable=True)

        gb3.configure_columns(' ',
            editable=False
        )

        gb3.configure_columns(si_select,
            cellEditor='agRichSelectCellEditor',
            cellEditorParams={'values':['']+st.session_state['position']},
            cellEditorPopup=True
        )

        gb3.configure_selection(selection_mode="multiple", use_checkbox=True)

        gb3.configure_grid_options(enableRangeSelection=True)

        si_position = AgGrid(
            df_siposition,
            gridOptions=gb3.build(),
            fit_columns_on_grid_load=True,
            enable_enterprise_modules=True,
            height=150,
            updateMode=GridUpdateMode.VALUE_CHANGED|GridUpdateMode.SELECTION_CHANGED
        )

        df_siposition = si_position['data'].set_index(' ')

        st.session_state['select_si_posi'] = si_position['selected_rows']

        st.button('チェックした項目を削除',on_click=delete_si_posi,args = (st.session_state['select_si_posi'],),key=32)
        st.session_state['dataframe']['Si_Position_data'] = df_siposition

    if len(st.session_state['position']) != 0:
        st.write('スタッフの人数を選びその人数で営業する際にスタッフのポジションで許可できるパターンを選択して下さい。(営業する可能性のある最大人数まで入力して下さい。)')
        if len(st.session_state['staffs']) != 0:
            mem = len(st.session_state['staffs'])
        else:
            mem = 10
        num = int(st.number_input('営業するスタッフの人数を選んで下さい。',1,20,1))
        pattern_posi = st.multiselect('許可できるパターンを選択して下さい。(複数選択、変更するときは選び直して決定ボタンを押して下さい。)',[['準備']*num]+make_pattern3([i for i in st.session_state['position'] if i != '準備'],st.session_state['position_num'],st.session_state['si_position'],num),format_func=lambda x: make_text(x))

        if st.button('決定',key=33):
            df_pt = pd.DataFrame(pattern_posi)
            st.session_state['pattern'][num] = df_pt

        if num in st.session_state['pattern'].keys():
            pat_idx = dict(zip(st.session_state['pattern'][num].index.values,['パターン'+str(i) for i in range(1,len(st.session_state['pattern'][num].index)+1)]))
            pat_col = dict(zip(st.session_state['pattern'][num].columns.values,['ポジション'+str(i) for i in range(1,len(st.session_state['pattern'][num].columns)+1)]))
            pat_df = st.session_state['pattern'][num].rename(columns=pat_col, index=pat_idx)
            st.write('↓'+str(num)+'人のとき同時に勤務できるポジションのパターン')
            st.dataframe(pat_df)

    #---------------------------------

    st.subheader('スタッフの種類の入力')

    col2_3 = st.columns(2)

    with col2_3[0]:
        att_input = st.text_input('追加したいスタッフの種類を入力して下さい。')
    with col2_3[1]:
        att_pri = st.number_input(str(att_input)+'の優先度を入力して下さい。',min_value=1,max_value=5)

    if st.button('追加',key=41):
        st.session_state['att'].append(att_input)
        st.session_state['para_att'].append(att_pri)

    df_att = pd.DataFrame({'属性':st.session_state['att'],
                            '優先度':st.session_state['para_att']})

    #表の生成
    gb4 = GridOptionsBuilder.from_dataframe(df_att, editable=True)

    gb4.configure_column('優先度',
            cellEditor='agRichSelectCellEditor',
            cellEditorParams={'values':[1,2,3,4,5]},
            cellEditorPopup=True
        )

    gb4.configure_selection(selection_mode="multiple", use_checkbox=True)

    st.write('スタッフの種類')

    att = AgGrid(
            df_att,
            gridOptions=gb4.build(),
            fit_columns_on_grid_load=True,
            enable_enterprise_modules=True,
            height=150,
            updateMode=GridUpdateMode.VALUE_CHANGED
        )

    st.session_state['att'] = list(att['data']['属性'])
    st.session_state['para_att'] = list(att['data']['優先度'])

    if st.button('チェックした項目を削除',key=42):
        for i in att['selected_rows']:
            st.session_state['att'].remove(i['属性'])
            st.session_state['para_att'].remove(i['優先度'])

    df_att = att['data'].set_index('属性')
    
    st.session_state['dataframe']['Att_data'] = df_att.T

    #---------------------------

    check1 = st.checkbox('店舗のデータを全て入力したらチェックを入れて下さい。')
    check2 = st.checkbox('スタッフのデータを全て入力したらチェックを入れて下さい。')
    check3 = st.checkbox('必要なポジションのデータを入力したらチェックを入れて下さい。')

    if check1 and check2 and check3:
        output,writer = open_excel()
        for i in st.session_state['dataframe'].keys():
            add_excel(writer,st.session_state['dataframe'][i],i)
        for i in st.session_state['pattern'].keys():
            add_excel(writer,st.session_state['pattern'][i],str(i)+'人のパターン')
        for i in st.session_state['staffs']:
            add_excel(writer,st.session_state['d_shift'][i],i)
        staff_xlsx = close_excel(output,writer) 
        st.download_button(label='ダウンロード',
                            data=staff_xlsx ,
                            file_name= 'input_data.xlsx')
    return

if __name__ == '__main__':
    main()
