Option Explicit
Dim cha1 As String'ChartObjectのこと．msgboxのための箱
Dim col1 As String'(Full)Seriescollectionのこと．msgboxのための箱
Dim poi1 As String'Pointsのこと．msgboxのための箱
Dim cha2 As Integer'ChartObjectのこと
Dim col2 As Integer'(Full)Seriescollectionのこと
Dim poi2 As Integer'Pointsのこと
cha1 = InputBox("何番目に作ったグラフのラベルを付けたいですか？整数で入れてください．" & Chr(13) & "グラフが1つの時は気にせずEnterを押してください")
col1=InputBox("選んだグラフの何番目のデータの種類にラベルをつけたいですか？整数で入れてください．" & Chr(13) & "グラフが1つの時は気にせずEnterを押してください")
poi1=InputBox("選んだデータ種類の何番目のデータにラベルをつけたいですか？整数で入れてください．" & Chr(13) & "なにも入れていないときは，最後のデータの値のみ示します")

If cha1 = "" Then
    cha2 = 1
Else
  cha2 = CInt(cha1)'String→Integerの変換
End If

If col1 = "" Then
    col1 = 1
Else
col2 = CInt(col1)  
End If

If poi1 = "" Then
  poi1=ActiveSheet.ChartObjects(cha1).Chart.SeriesCollection(col1).Points.Count
Else
poi2 = CInt(poi1)  
End If


'以下のチャートで
'1.積み立て棒グラフ等で一部のデータ種類のみ，関係ないセルの値を表示させる(cha,colのみ選択．poiは自動でやらす.)
'2.積み立て棒グラフ等で各データ種類の特定番号のデータについて，データラベルの値を表示させる(cha,poiのみ選択. colは自動選択)
'3. 積み立て棒グラフ等で，まったく関係ないRangeのデータをデータラベルにする(chaのみ選択．選択されたデータRangeとSeriesCollection, Pointsの数が等しいことをまず確認)

r = Selection(1).Row '行のスタート
c = Selection(1).Column '列の指定
rc = Section.Rows.Count '行の数
cc=Section.Column.Count '列の数

if rc=.SeriesCollection.Count And cc=.SeriesCollection(col1).Points.Count Then
  'それぞれのデータラベルに代入していく
Else
  msgbox "設定したいラベルのデータ数とデータ数が一致しません．" & Chr(13) & "疲れていませんか？？休憩しませんか？？＾＾"
  
