VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Dim AAA As Integer

For Counter2 = 1 To 2
 For Counter0 = 2 To Sheet1.Range("I1").Value
  For Counter1 = 2 To Sheet1.Range("I1").Value
  
    AAA = 21 * (Counter0 - 2) + Counter1
   Sheet1.Range("I3").Value = Sheet1.Cells(Counter0, 2).Value 'u Šî€
   Sheet1.Range("J3").Value = Sheet1.Cells(Counter0, 3).Value 'v Šî€
   Sheet1.Range("I4").Value = Sheet1.Cells(Counter1, 2).Value 'u ‘ÎÛ
   Sheet1.Range("J4").Value = Sheet1.Cells(Counter1, 3).Value 'v ‘ÎÛ
 
   Sheet1.Cells(AAA, 13).Value = Sheet1.Cells(Counter0, 1).Value 'Šî€‘¤
   Sheet1.Cells(AAA, 14).Value = Sheet1.Cells(Counter1, 1).Value '‘ÎÛ‘¤
   Sheet1.Cells(AAA, 15).Value = Sheet1.Range("K3").Value '‡™u'v'
  'Sheets("InputOutput").Cells("D4").Value = Sheet1.Cells(Counter, 3).Value 'y Šî€
  'Sheets("InputOutput").Cells("C13").Value = Sheet1.Cells(Counter + 1, 2).Value 'x ‘ÎÛ
  'Sheets("InputOutput").Cells("D13").Value = Sheet1.Cells(Counter + 1, 3).Value 'y ‘ÎÛ

  Next Counter1
 Next Counter0
 
   '19-9
   Sheet1.Range("I3").Value = Sheet1.Range("B20").Value 'u Šî€
   Sheet1.Range("J3").Value = Sheet1.Range("C20").Value 'v Šî€
   Sheet1.Range("I4").Value = Sheet1.Range("B10").Value 'u ‘ÎÛ
   Sheet1.Range("J4").Value = Sheet1.Range("C10").Value 'v ‘ÎÛ
 
   Sheet1.Range("F7").Value = Sheet1.Range("K3").Value '‡™u'v'
   
 Next Counter2
 
 
End Sub
