Attribute VB_Name = "Module3"

Sub JND21()
Attribute JND21.VB_Description = "マクロ記録日 : 2006/8/29  ユーザー名 : masaoka"
Attribute JND21.VB_ProcData.VB_Invoke_Func = " \n14"

Dim AAA, BBB, CCC, DDD, AA, BB As Long
Dim CC, DD As Double


''''''''''''''''''''''''''21MaxJND'''''''''''''''''''''''''''''''''''''''''''''''
    
   Sheets("JND計算").Range("I17:J512").Value = 0 '前回対象色度CxCy消去
   Sheets("JND計算").Range("o14:Q454").Value = 0 '前回結果消去
   AA = Sheets("JND計算").Range("B3").Value 'point数
   BB = AA + 1
   
'''''対象色度CxCy入力'''''

For Counter0 = 2 To BB

   Sheets("JND計算").Cells(Counter0 + 15, 9).Value = Sheets("xy入力").Cells(Counter0, 2).Value 'x 対象色度Cx入力
   Sheets("JND計算").Cells(Counter0 + 15, 10).Value = Sheets("xy入力").Cells(Counter0, 3).Value 'y 対象色度Cy入力
   Sheets("JND計算").Cells(Counter0 + 15, 12).Value = Counter0 - 1 '対象point入力

Next Counter0


'''''基準色度CxCy入力と全データ表記'''''

  For Counter1 = 2 To BB
  
    CC = 14 + AA * (Counter1 - 2)
    DD = CC + AA - 1
    
   Sheets("JND計算").Range("I16").Value = Sheets("xy入力").Cells(Counter1, 2).Value 'x 基準色度Cx入力
   Sheets("JND計算").Range("J16").Value = Sheets("xy入力").Cells(Counter1, 3).Value 'y 基準色度Cy入力
   Sheets("JND計算").Range("L16").Value = Counter1 - 1 '基準point入力

 
   Sheets("JND計算").Range(Sheets("JND計算").Cells(CC, 15), Sheets("JND計算").Cells(DD, 15)).Value = Sheets("JND計算").Range(Sheets("JND計算").Cells(17, 11), Sheets("JND計算").Cells(AA + 17, 11)).Value 'JND表記
   Sheets("JND計算").Range(Sheets("JND計算").Cells(CC, 16), Sheets("JND計算").Cells(DD, 16)).Value = Sheets("JND計算").Range("L16").Value '基準point表記
   Sheets("JND計算").Range(Sheets("JND計算").Cells(CC, 17), Sheets("JND計算").Cells(DD, 17)).Value = Sheets("JND計算").Range(Sheets("JND計算").Cells(17, 12), Sheets("JND計算").Cells(AA + 17, 12)).Value '対象point表記
  
  Next Counter1
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
