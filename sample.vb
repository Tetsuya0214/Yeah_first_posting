'http://officetanaka.net/excel/vba/tips/tips111.htm

'選択してボタン押すと，グラフに別設定したラベルがつく設定を作る

'現在，何行目か等は完全に手打ちしているので自動化


Sub New1()

r = Selection(1).Row '行のスタート
c = Selection(1).Column '列の指定
't = Section.Rows.Count '行の数

Cells(r + 5, c).Value = 100

End Sub


If *****=t Then
	****** '列の指定のとこ=c