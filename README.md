# iss(izakaya-shift-supporter)
 
issは居酒屋スタッフの勤務表を自動的に作成することができるソフトウェアです。
 
# Features
 
issは30分毎にスタッフに勤務を割り当てることができるため、固定のシフトが存在しないような居酒屋の勤務表の作成を行うことができます。

また、居酒屋によっては複数のポジション(ホール、キッチンなど)を1人のスタッフが行うという場面も想定されますが、そのような場面に対応した勤務表の作成も行うことができます。

スタッフの不足する時間帯が発生したときにはすぐにわかるようになっていて、多店舗へのヘルプの要請などを行いやすいようになっています。
 
# Requirement
 
* cplex 12.10
* python 3.7
* pulp 2.3.1
* streamlit 1.12.0

\* cplexは有料です。
 
# Usage

* 入力フォーム

「店舗データ.py」が入力フォームを表示するファイルとなっています。streamlitのマルチページ機能を使用しているため、ディレクトリ「pages」は名前を変更せずに「店舗データ.py」が入っているディレクトリと同じディレクトリに入れて下さい。

ターミナルで、

```bash
streamlit run 店舗データ.py
```

を実行することでWebブラウザが立ち上がりアプリケーションが起動します。

まず、店舗データのページの指示に従いデータを入力して下さい。

店舗データの入力が終わったら他のページのデータを入力して下さい。

全てのデータの入力が終わったらシフト希望提出のページから各スタッフのシフト希望提出用シートをアップロードして下さい。シフト希望提出用シートは店舗データのページからダウンロードすることができます。

スタッフ全員分のシフト希望提出用シートをアップロードしたら、店舗データのページ最下部のチェック項目を確認しながらチェックを入れてください。チェックを全て入れることでダウンロードボタンが出てくるのでボタンをクリックし、入力用のエクセルファイルをダウンロードして下さい。

同じ店舗で再び勤務表を作成するときには、店舗データのページ上部から入力用のエクセルファイルをアップロードすることで店舗データの入力を行うことができます。

* 「iss_calculate.py」

「iss_calculate.py」では、入力フォームから得られた入力用のエクセルファイルを読み込むことで計算を行いその結果を出力します。

入力用エクセルファイルは、10行目に

```bash
file_name = 'input_data1.xlsx'
```

と設定して下さい。。「input_data1.xlsx」、「input_data2.xlsx」、「input_data3.xlsx」の3つファイルは入力用のサンプルデータで既にデータの入力が行われています。

また、入力には前の月の勤務表も必要になります。
前の月の勤務表は、13行目に

```bash
last_month = 'last_month_result.xlsx'
```

と入力して下さい。初めてissを使う時はサンプルデータ「last_month_result.xlsx」を入力して下さい。ある月の勤務表を作成し、次の月の勤務表を作成するときには作成済みの勤務表を入力して下さい。

計算結果はテキストファイルで得られます。
結果のファイル名は、12行目に

```bash
output_file = 'output.txt'
```

と設定して下さい。

計算時間の上限と計算を停止する相対的な最適性ギャップを設定することができます。計算時間の上限は、timelimitで秒数を指定することができます。 計算を停止する相対的な最適性ギャップはgapRelで指定することができます。例えば、相対的な最適性ギャップが10%を下回ったときに計算を停止したいときには「gapRel = 0.1」と設定します。

```bash
status = prob.solve(pulp.CPLEX_CMD(keepFiles = 1, timelimit = 3600 * 1, gapRel = 0.3, options = ['set emphasis mip 4']))
```

全ての設定が終わったら、ターミナルから

```bash
python iss_calculate.py
```

を実行することで、「iss_calculate.py」と同じディレクトリに計算結果のファイルを作成します。

* 「iss_make_schedule.py」

「iss_make_schedule.py」では、実際に勤務表の作成を行います。

入力ファイルとして、「iss_calculate.py」で得られた計算結果のファイルを、13行目に

```bash
input_file = 'output.txt'
```

と入力して下さい。

出力される勤務表のファイル名は、

15行目に

```bash
output_file = 'iss_result.xlsx'
```

と入力して下さい。

入力が終わったら、ターミナルから

```bash
python iss_make_schedule.py
```

を実行することで、「iss_make_schedule.py」と同じディレクトリに勤務表のファイルを作成します。
 
# Note

計算に有料ソルバーcplexを利用しています。使用前に購入をお願いします。

streamlitのマルチページ機能を使うためにはstreamlit v1.10.0以上が必要です。

規定のブラウザとしてGoogle Chromeを設定することを推奨します。
 
# Author
 
* 中村克
* 電気通信大学大学院
* s.nakamura1110@gmail.com
