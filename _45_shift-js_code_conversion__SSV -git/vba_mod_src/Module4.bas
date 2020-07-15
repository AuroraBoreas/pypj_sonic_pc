Attribute VB_Name = "Module4"
Sub henkan()

Dim Cx0, Cx1, Cy0, Cy1, Cu0, Cu1, Cv0, Cv1, uv, JND As Double

On Error GoTo エラー処理へ

Sheets("⊿u'v'⇒JND変換").Range("c8:d8").Value = ""

Cx0 = Sheets("⊿u'v'⇒JND変換").Range("e8").Value '基準Cx
Cy0 = Sheets("⊿u'v'⇒JND変換").Range("f8").Value '基準Cy
Cx1 = Sheets("⊿u'v'⇒JND変換").Range("g8").Value '対象Cx
Cy1 = Sheets("⊿u'v'⇒JND変換").Range("h8").Value '対象Cy


Cu0 = 4 * Cx0 / (-2 * Cx0 + 12 * Cy0 + 3) '基準Cu'
Cv0 = 9 * Cy0 / (-2 * Cx0 + 12 * Cy0 + 3) '基準Cv'
Cu1 = 4 * Cx1 / (-2 * Cx1 + 12 * Cy1 + 3) '対象Cu'
Cv1 = 9 * Cy1 / (-2 * Cx1 + 12 * Cy1 + 3) '対象Cv'


''''⊿u'v'計算''''

uv = ((Cu1 - Cu0) * (Cu1 - Cu0) + (Cv1 - Cv0) * (Cv1 - Cv0)) ^ (1 / 2) '⊿u'v'値
Sheets("⊿u'v'⇒JND変換").Range("c8").Value = uv '⊿u'v'値結果表示


''''JND計算''''

Sheets("JND計算").Range("i16").Value = Cx0 'JND計算シート代入(基準Cx)
Sheets("JND計算").Range("i17").Value = Cx1 'JND計算シート代入(対象Cx)
Sheets("JND計算").Range("j16").Value = Cy0 'JND計算シート代入(基準Cy)
Sheets("JND計算").Range("j17").Value = Cy1 'JND計算シート代入(対象Cy)

JND = Sheets("JND計算").Range("k17").Value 'JND値
Sheets("⊿u'v'⇒JND変換").Range("d8").Value = JND 'JND値結果表示

Exit Sub


''''エラー処理''''

エラー処理へ:

Sheets("⊿u'v'⇒JND変換").Range("c8:h8").Value = ""



End Sub
