Attribute VB_Name = "Module3"

Sub JND21()
Attribute JND21.VB_Description = "�}�N���L�^�� : 2006/8/29  ���[�U�[�� : masaoka"
Attribute JND21.VB_ProcData.VB_Invoke_Func = " \n14"

Dim AAA, BBB, CCC, DDD, AA, BB As Long
Dim CC, DD As Double


''''''''''''''''''''''''''21MaxJND'''''''''''''''''''''''''''''''''''''''''''''''
    
   Sheets("JND�v�Z").Range("I17:J512").Value = 0 '�O��ΏېF�xCxCy����
   Sheets("JND�v�Z").Range("o14:Q454").Value = 0 '�O�񌋉ʏ���
   AA = Sheets("JND�v�Z").Range("B3").Value 'point��
   BB = AA + 1
   
'''''�ΏېF�xCxCy����'''''

For Counter0 = 2 To BB

   Sheets("JND�v�Z").Cells(Counter0 + 15, 9).Value = Sheets("xy����").Cells(Counter0, 2).Value 'x �ΏېF�xCx����
   Sheets("JND�v�Z").Cells(Counter0 + 15, 10).Value = Sheets("xy����").Cells(Counter0, 3).Value 'y �ΏېF�xCy����
   Sheets("JND�v�Z").Cells(Counter0 + 15, 12).Value = Counter0 - 1 '�Ώ�point����

Next Counter0


'''''��F�xCxCy���͂ƑS�f�[�^�\�L'''''

  For Counter1 = 2 To BB
  
    CC = 14 + AA * (Counter1 - 2)
    DD = CC + AA - 1
    
   Sheets("JND�v�Z").Range("I16").Value = Sheets("xy����").Cells(Counter1, 2).Value 'x ��F�xCx����
   Sheets("JND�v�Z").Range("J16").Value = Sheets("xy����").Cells(Counter1, 3).Value 'y ��F�xCy����
   Sheets("JND�v�Z").Range("L16").Value = Counter1 - 1 '�point����

 
   Sheets("JND�v�Z").Range(Sheets("JND�v�Z").Cells(CC, 15), Sheets("JND�v�Z").Cells(DD, 15)).Value = Sheets("JND�v�Z").Range(Sheets("JND�v�Z").Cells(17, 11), Sheets("JND�v�Z").Cells(AA + 17, 11)).Value 'JND�\�L
   Sheets("JND�v�Z").Range(Sheets("JND�v�Z").Cells(CC, 16), Sheets("JND�v�Z").Cells(DD, 16)).Value = Sheets("JND�v�Z").Range("L16").Value '�point�\�L
   Sheets("JND�v�Z").Range(Sheets("JND�v�Z").Cells(CC, 17), Sheets("JND�v�Z").Cells(DD, 17)).Value = Sheets("JND�v�Z").Range(Sheets("JND�v�Z").Cells(17, 12), Sheets("JND�v�Z").Cells(AA + 17, 12)).Value '�Ώ�point�\�L
  
  Next Counter1
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
