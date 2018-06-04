Dim a As Integer'ChartObjectのこと
Dim b As Integer'(Full)Seriescollectionのこと
Dim c As Integer'Pointsのこと
a = InputBox("何番目に作った表のラベルを付けたいですか？整数で入れてください．")
b=InputBox("選んだ表の何番目のデータの種類にラベルをつけたいですか？整数で入れてください．")
c=InputBox("選んだデータ種類のの何番目のデータにラベルをつけたいですか？整数で入れてください．")

'1.積み立て棒グラフ等で一部のデータ種類のみ，関係ないセルの値を表示させる(a,bのみ選択．cは自動でやらす.)
'2.積み立て棒グラフ等で各データ種類の特定番号のデータについて，データラベルの値を表示させる(a,cのみ選択. bは自動選択)
'3. 積み立て棒グラフ等で，まったく関係ないRangeのデータをデータラベルにする(選択されたデータRangeとSeriesCollection, Pointsの数が等しいことをまず確認)

