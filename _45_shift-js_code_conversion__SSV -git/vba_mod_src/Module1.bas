Attribute VB_Name = "Module1"
Sub JND496()

Dim AAA, BBB, CCC, DDD, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double

    
'''''''''''''''''''''''''''''''''496MaxJND'''''''''''''''''''''''''''''''''''''''''
      
'''''�ΏېF�xCxCy����'''''
      
       Sheets("JND�v�Z").Range("H9").Value = 0 '�O��MaxJND����
       Sheets("JND�v�Z").Range("I17:J512").Value = 0 '�O��ΏېF�xCxCy����
       
       AAA = Sheets("JND�v�Z").Range("B8").Value '�s��
       BBB = Sheets("JND�v�Z").Range("B9").Value '��

     For AA = 7 To BBB + 6
       For CCC = 2 To AAA + 1
       
       CC = AAA * (AA - 6) + CCC - 1 + (16 - AAA)
       
        Sheets("JND�v�Z").Cells(CC, 9).Value = Sheets("xy����").Cells(CCC, AA).Value 'x �ΏېF�xCx
        Sheets("JND�v�Z").Cells(CC, 10).Value = Sheets("xy����").Cells(CCC + 16, AA).Value 'y �ΏېF�xCy
        Sheets("JND�v�Z").Cells(CC, 12).Value = CCC - 1 'y �Ώ�(�s)����
        Sheets("JND�v�Z").Cells(CC, 13).Value = AA - 6 'y �Ώ�(��)����
        
       Next CCC
     Next AA

      
'''''�CxCy���͂�MaxJND�l�Epoint����'''''

      EE = 0
      
      For CounterB = 7 To BBB + 6
       For CounterC = 2 To AAA + 1
   
       
        DD = Sheets("JND�v�Z").Range("H9").Value '����MaxJND
  
                
        If DD < EE Then '����MaxJND VS �ǉ�MaxJND �召��r

        Sheets("JND�v�Z").Range("H9").Value = Sheets("JND�v�Z").Range("K14").Value '496MaxJND����
        Sheets("JND�v�Z").Range("j9").Value = Sheets("JND�v�Z").Range("l16").Value 'MaxJND �(�s)����
        Sheets("JND�v�Z").Range("k9").Value = Sheets("JND�v�Z").Range("m16").Value 'MaxJND �(��)����
        Sheets("JND�v�Z").Range("l9").Value = Sheets("JND�v�Z").Range("l14").Value 'MaxJND �Ώ�(�s)����
        Sheets("JND�v�Z").Range("m9").Value = Sheets("JND�v�Z").Range("m14").Value 'MaxJND �Ώ�(��)����
        
        End If
            

        Sheets("JND�v�Z").Cells(16, 9).Value = Sheets("xy����").Cells(CounterC, CounterB).Value 'x ��F�xCx����
        Sheets("JND�v�Z").Cells(16, 10).Value = Sheets("xy����").Cells(CounterC + 16, CounterB).Value 'y ��F�xCy����
        Sheets("JND�v�Z").Cells(16, 12).Value = CounterC - 1 'x �point����
        Sheets("JND�v�Z").Cells(16, 13).Value = CounterB - 6 'y �point����
        
        
        EE = Sheets("JND�v�Z").Range("K14").Value '�ǉ�MaxJND
 
                  
       Next CounterC
      Next CounterB
      
      
'''''Max point�ڈ�'''''

        Sheets("xy����").Range("E2:AK33").Interior.ColorIndex = 0 '�O��ڈ����
  
        FF = Sheets("JND�v�Z").Range("j9").Value + 1 'MaxJND �(�s)
        GG = Sheets("JND�v�Z").Range("k9").Value + 6 'MaxJND �(��)
        HH = FF + 16
        II = Sheets("JND�v�Z").Range("l9").Value + 1 'MaxJND �Ώ�(�s)
        JJ = Sheets("JND�v�Z").Range("m9").Value + 6 'MaxJND �Ώ�(��)
        KK = II + 16
        
        
        Sheets("xy����").Cells(FF, GG).Interior.ColorIndex = 3 'x �point�ڈ�
        Sheets("xy����").Cells(HH, GG).Interior.ColorIndex = 3 'y �point�ڈ�
        
        Sheets("xy����").Cells(II, JJ).Interior.ColorIndex = 7 'x �Ώ�point�ڈ�
        Sheets("xy����").Cells(KK, JJ).Interior.ColorIndex = 7 'y �Ώ�point�ڈ�
        
       
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

