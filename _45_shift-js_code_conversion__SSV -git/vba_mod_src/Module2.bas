Attribute VB_Name = "Module2"
Sub uv21()

Dim AAA, BBB, CCC, DDD, AA, BB As Long
Dim CC, DD, EE, EEE As Double

   
  ''''''''''''''''''''''''''21Max��u'v''''''''''''''''''''''''''''''''''''''''''''''''
   
   Sheets("��u'v'�v�Z").Range("h17:i512").Value = 0 '�O��ΏېF�xCxCy����
   Sheets("��u'v'�v�Z").Range("o14:r454").Value = 0 '�O�񌋉ʏ���
   AA = Sheets("��u'v'�v�Z").Range("B3").Value 'point��
   BB = AA + 1
     
'''''�ΏېF�xu'v'����'''''

For Counter0 = 2 To BB

   Sheets("��u'v'�v�Z").Cells(Counter0 + 15, 8).Value = Sheets("u'v'�ϊ�").Cells(Counter0, 2).Value 'u �ΏېF�xu'����
   Sheets("��u'v'�v�Z").Cells(Counter0 + 15, 9).Value = Sheets("u'v'�ϊ�").Cells(Counter0, 3).Value 'v �Ώ�v'�F�x����
   Sheets("��u'v'�v�Z").Cells(Counter0 + 15, 12).Value = Counter0 - 1 '�Ώ�point����

Next Counter0


'''''��F�xu'v'���͂ƑS�f�[�^�\�L'''''

  For Counter1 = 2 To BB
  
    CC = 14 + AA * (Counter1 - 2)
    DD = CC + AA - 1
    

   Sheets("��u'v'�v�Z").Range("h16").Value = Sheets("u'v'�ϊ�").Cells(Counter1, 2).Value 'u ��F�xu'
   Sheets("��u'v'�v�Z").Range("i16").Value = Sheets("u'v'�ϊ�").Cells(Counter1, 3).Value 'v ��F�xv'
   Sheets("��u'v'�v�Z").Range("L16").Value = Counter1 - 1 '�point����

 
   Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(CC, 15), Sheets("��u'v'�v�Z").Cells(DD, 15)).Value = Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(17, 10), Sheets("��u'v'�v�Z").Cells(AA + 17, 10)).Value ' u'>v' ��u'v'�\�L
   Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(CC, 16), Sheets("��u'v'�v�Z").Cells(DD, 16)).Value = Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(17, 11), Sheets("��u'v'�v�Z").Cells(AA + 17, 11)).Value ' u'<=v' ��u'v'�\�L
   Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(CC, 17), Sheets("��u'v'�v�Z").Cells(DD, 17)).Value = Sheets("��u'v'�v�Z").Range("L16").Value '�point�\�L
   Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(CC, 18), Sheets("��u'v'�v�Z").Cells(DD, 18)).Value = Sheets("��u'v'�v�Z").Range(Sheets("��u'v'�v�Z").Cells(17, 12), Sheets("��u'v'�v�Z").Cells(AA + 17, 12)).Value '�Ώ�point�\�L

  Next Counter1
 
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
End Sub
Sub uv496()

Dim AAA, BBB, CCC, DDD, EEE, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double
Dim FFF, GGG, HHH, III, JJJ, KKK As Double

    
'''''''''''''''''''''''''''''''''496Max��u'v'''''''''''''''''''''''''''''''''''''''''


      Sheets("��u'v'�v�Z").Range("H9:H10").Value = 0 '�O��Max��u'v'����
      Sheets("��u'v'�v�Z").Range("h17:i512").Value = 0 '�O��F�xu'v'����
  
  
       AAA = Sheets("��u'v'�v�Z").Range("B8").Value '�s��
       BBB = Sheets("��u'v'�v�Z").Range("B9").Value '��
       
       
'''''�ΏېF�xu'v'����'''''

      For AA = 7 To BBB + 6
       For CCC = 2 To AAA + 1
        CC = AAA * (AA - 6) + CCC - 1 + (16 - AAA)
        Sheets("��u'v'�v�Z").Cells(CC, 8).Value = Sheets("u'v'�ϊ�").Cells(CCC, AA).Value 'u �ΏېF�xu'����
        Sheets("��u'v'�v�Z").Cells(CC, 9).Value = Sheets("u'v'�ϊ�").Cells(CCC + 16, AA).Value 'v �ΏېF�xv'����
        Sheets("��u'v'�v�Z").Cells(CC, 12).Value = CCC - 1 'u �Ώ�(�s)����
        Sheets("��u'v'�v�Z").Cells(CC, 13).Value = AA - 6 'v �Ώ�(��)����
       Next CCC
      Next AA
      
      
'''''�u'v'���͂�Max��u'v'�l�Epoint����'''''
        
        
       EE = 0
       EEE = 0
       
  For CounterB = 7 To BBB + 6
  
   For CounterC = 2 To AAA + 1
   
       CC = AAA * (CounterB - 6) + CounterC - 1 + (16 - AAA)
      
        DDD = Sheets("��u'v'�v�Z").Range("H9").Value 'u'>v' ����Max��u'v'
        DD = Sheets("��u'v'�v�Z").Range("H10").Value 'u'<=v' ����Max��u'v'
        

   ''' u'<=v' Max��u'v'�l�Epoint����'''
        
        If DD < EE Then

        Sheets("��u'v'�v�Z").Range("H10").Value = Sheets("��u'v'�v�Z").Range("K14").Value '496Max��u'v'����
        Sheets("��u'v'�v�Z").Range("k10").Value = Sheets("��u'v'�v�Z").Range("l16").Value 'Max��u'v' �(�s)����
        Sheets("��u'v'�v�Z").Range("l10").Value = Sheets("��u'v'�v�Z").Range("m16").Value 'Max��u'v' �(��)����
        Sheets("��u'v'�v�Z").Range("m10").Value = Sheets("��u'v'�v�Z").Range("l14").Value 'Max��u'v' �Ώ�(�s)����
        Sheets("��u'v'�v�Z").Range("n10").Value = Sheets("��u'v'�v�Z").Range("m14").Value 'Max��u'v' �Ώ�(��)����
        
        End If
                
                
        
   ''' u'>v' Max��u'v'�l�Epoint����'''
        
        If DDD < EEE Then

        Sheets("��u'v'�v�Z").Range("H9").Value = Sheets("��u'v'�v�Z").Range("g14").Value '496Max��u'v'����
        Sheets("��u'v'�v�Z").Range("k9").Value = Sheets("��u'v'�v�Z").Range("l16").Value 'Max��u'v' �(�s)����
        Sheets("��u'v'�v�Z").Range("l9").Value = Sheets("��u'v'�v�Z").Range("m16").Value 'Max��u'v' �(��)����
        Sheets("��u'v'�v�Z").Range("m9").Value = Sheets("��u'v'�v�Z").Range("h14").Value 'Max��u'v' �Ώ�(�s)����
        Sheets("��u'v'�v�Z").Range("n9").Value = Sheets("��u'v'�v�Z").Range("i14").Value 'Max��u'v' �Ώ�(��)����
        
        End If
        
        
   '''�u'v'����'''
        
        Sheets("��u'v'�v�Z").Cells(16, 8).Value = Sheets("u'v'�ϊ�").Cells(CounterC, CounterB).Value 'x ��F�xu'����
        Sheets("��u'v'�v�Z").Cells(16, 9).Value = Sheets("u'v'�ϊ�").Cells(CounterC + 16, CounterB).Value 'y ��F�xv'����
        Sheets("��u'v'�v�Z").Cells(16, 12).Value = CounterC - 1 'x �(�s)����
        Sheets("��u'v'�v�Z").Cells(16, 13).Value = CounterB - 6 'y �(��)����
        
        EEE = Sheets("��u'v'�v�Z").Range("g14").Value ' u'>v' �ǉ�Max��u'v'
        EE = Sheets("��u'v'�v�Z").Range("k14").Value ' u'<=v' �ǉ�Max��u'v'

    Next CounterC
   Next CounterB
  
  
'''''Max point�ڈ�'''''
  
        Sheets("u'v'�ϊ�").Range("E2:Ak33").Interior.ColorIndex = 0 '�O��ڈ����
        
  ''' u'>v' '''
        FF = Sheets("��u'v'�v�Z").Range("k9").Value + 1 'Max��u'v' �(�s)
        GG = Sheets("��u'v'�v�Z").Range("l9").Value + 6 'Max��u'v' �(��)
        HH = FF + 16
        II = Sheets("��u'v'�v�Z").Range("m9").Value + 1 'Max��u'v' �Ώ�(�s)
        JJ = Sheets("��u'v'�v�Z").Range("n9").Value + 6 'Max��u'v' �Ώ�(��)
        KK = II + 16
        
        Sheets("u'v'�ϊ�").Cells(FF, GG).Interior.ColorIndex = 3 ' u' �point�ڈ�
        Sheets("u'v'�ϊ�").Cells(HH, GG).Interior.ColorIndex = 3 ' v' �point�ڈ�
        
        Sheets("u'v'�ϊ�").Cells(II, JJ).Interior.ColorIndex = 7 ' u' �Ώ�point�ڈ�
        Sheets("u'v'�ϊ�").Cells(KK, JJ).Interior.ColorIndex = 7 ' v' �Ώ�point�ڈ�
        
        
        
 ''' u'<=v' '''
        FFF = Sheets("��u'v'�v�Z").Range("k10").Value + 1 'Max��u'v' �(�s)
        GGG = Sheets("��u'v'�v�Z").Range("l10").Value + 6 'Max��u'v' �(��)
        HHH = FFF + 16
        III = Sheets("��u'v'�v�Z").Range("m10").Value + 1 'Max��u'v' �Ώ�(�s)
        JJJ = Sheets("��u'v'�v�Z").Range("n10").Value + 6 'Max��u'v' �Ώ�(��)
        KKK = III + 16
        
        Sheets("u'v'�ϊ�").Cells(FFF, GGG).Interior.ColorIndex = 5 ' u' �point�ڈ�
        Sheets("u'v'�ϊ�").Cells(HHH, GGG).Interior.ColorIndex = 5 ' v' �point�ڈ�
        
        Sheets("u'v'�ϊ�").Cells(III, JJJ).Interior.ColorIndex = 8 ' u' �Ώ�point�ڈ�
        Sheets("u'v'�ϊ�").Cells(KKK, JJJ).Interior.ColorIndex = 8 ' v' �Ώ�point�ڈ�
        
        
'''''Max��u'v'��JND�ϊ�'''''

  ''' u'>v' '''

        Sheets("��u'v'�v�Z").Range("o9").Value = Sheets("xy����").Cells(FF, GG).Value ' u' �pointCx����
        Sheets("��u'v'�v�Z").Range("p9").Value = Sheets("xy����").Cells(HH, GG).Value ' v' �pointCy����
        Sheets("��u'v'�v�Z").Range("q9").Value = Sheets("xy����").Cells(II, JJ).Value ' u' �Ώ�pointCx����
        Sheets("��u'v'�v�Z").Range("r9").Value = Sheets("xy����").Cells(KK, JJ).Value ' v' �Ώ�pointCy����

        Sheets("JND�v�Z").Range("I16").Value = Sheets("��u'v'�v�Z").Range("o9").Value 'JND�v�Z�V�[�g�ɊCx���
        Sheets("JND�v�Z").Range("j16").Value = Sheets("��u'v'�v�Z").Range("p9").Value 'JND�v�Z�V�[�g�ɊCy���
        Sheets("JND�v�Z").Range("I17").Value = Sheets("��u'v'�v�Z").Range("q9").Value 'JND�v�Z�V�[�g�ɑΏ�Cx���
        Sheets("JND�v�Z").Range("j17").Value = Sheets("��u'v'�v�Z").Range("r9").Value 'JND�v�Z�V�[�g�ɑΏ�Cy���
        
        Sheets("��u'v'�v�Z").Range("j9").Value = Sheets("JND�v�Z").Range("k17").Value 'JND�v�Z���ʓ���

 ''' u'<=v' '''
        Sheets("��u'v'�v�Z").Range("o10").Value = Sheets("xy����").Cells(FFF, GGG).Value ' u' �pointCx����
        Sheets("��u'v'�v�Z").Range("p10").Value = Sheets("xy����").Cells(HHH, GGG).Value ' v' �pointCy����
        Sheets("��u'v'�v�Z").Range("q10").Value = Sheets("xy����").Cells(III, JJJ).Value ' u' �Ώ�pointCx����
        Sheets("��u'v'�v�Z").Range("r10").Value = Sheets("xy����").Cells(KKK, JJJ).Value ' v' �Ώ�pointCy����
        
        Sheets("JND�v�Z").Range("I16").Value = Sheets("��u'v'�v�Z").Range("o10").Value 'JND�v�Z�V�[�g�ɊCx���
        Sheets("JND�v�Z").Range("j16").Value = Sheets("��u'v'�v�Z").Range("p10").Value 'JND�v�Z�V�[�g�ɊCy���
        Sheets("JND�v�Z").Range("I17").Value = Sheets("��u'v'�v�Z").Range("q10").Value 'JND�v�Z�V�[�g�ɑΏ�Cx���
        Sheets("JND�v�Z").Range("j17").Value = Sheets("��u'v'�v�Z").Range("r10").Value 'JND�v�Z�V�[�g�ɑΏ�Cy���
        
        Sheets("��u'v'�v�Z").Range("j10").Value = Sheets("JND�v�Z").Range("k17").Value 'JND�v�Z���ʓ���
        
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
