VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MUTOMOTO 
   Caption         =   "UserForm1"
   ClientHeight    =   7410
   ClientLeft      =   50
   ClientTop       =   590
   ClientWidth     =   10260
   OleObjectBlob   =   "MUTOMOTO.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MUTOMOTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()


On Error GoTo �G���[������


 Set myappl = CreateObject("CAS20W.Application") 'CA2000�\�t�g�Ƃ̃����N


'''''''''''����p�v���O����''''''''''''''

    Dim Ret As Long
    Ret = myappl.Measure() 'CA2000 ����v���O����
    
    Do
    'Ret = MsgBox("Cancel?", MsgBoxStyle.OKCancel, "")
    'If Ret = MsgBoxResult.OK Then
    'If myappl.MeasureCancel() = 0 Then
    'Exit Do
    'End If
    'Else
    
    If myappl.PollingMeasure() = 0 Then
    Exit Do
    End If
    'End If
    'Loop Until Ret = MsgBoxResult.OK
    Loop Until Ret = -1
    


''''''''''496Spot�f�[�^�擾�v���O����''''''''''''

    Dim SpotCond As Object
    Dim x, y As Long

    Ret = myappl.SelectData()

    Set SpotCond = myappl.GetSpotCondition() 'CA2000 Spot�ݒ�������擾
    
    Dim SpotCount As Long
    SpotCount = SpotCond.GetSpotCount() 'CA2000 Spot�����擾
    
    Dim SpotData(1000000) As Single 'CA2000 Spot�����̔z��m��
    
        
 '''''''CA2000 496Cx�f�[�^�擾'''''''
    
    Element = 4 'Spot Cx�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    x = 1
    y = 1
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(y + 1, x + 6).Value = ""  '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(y + 1, x + 6).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(y + 1, x + 6).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(y + 1, x + 6).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
        x = x + 1
        
        If x = 32 Then
        y = y + 1
        x = 1
        
        End If

    Next i
    
    
    
    For i = 496 To 504
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 494, 41).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 494, 41).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 494, 41).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 494, 41).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
    
    
'''''''CA2000 25point&SJpoint Cx�f�[�^�擾 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 452, 41).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 452, 41).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 452, 41).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 452, 41).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    
    
    
 '''''''CA2000 496Cy�f�[�^�擾'''''''
    
        Element = 5 'Spot Cy�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    x = 1
    y = 1
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(y + 17, x + 6).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(y + 17, x + 6).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(y + 17, x + 6).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(y + 17, x + 6).Value = SpotData(i) '�擾����Cy�f�[�^��Excel�ɓ���
            
        End If
        
        x = x + 1
        
        If x = 32 Then
        y = y + 1
        x = 1
        
        End If

    Next i
     
     
    For i = 496 To 504
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 494, 42).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 494, 42).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 494, 42).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 494, 42).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
    
 '''''''CA2000 25point&SJpoint Cy�f�[�^�擾 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 452, 42).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 452, 42).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 452, 42).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 452, 42).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    

'''''''CA2000 496�P�x�f�[�^�擾'''''''

       Element = 3 'Spot �P�x�f�[�^�̎擾
       Ret = myappl.GetSpotData(Element, SpotData)
       x = 1
       y = 1
    
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(y + 33, x + 6).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(y + 33, x + 6).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(y + 33, x + 6).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(y + 33, x + 6).Value = SpotData(i) '�擾�����P�x�f�[�^��Excel�ɓ���
            
        End If
        
        x = x + 1
        
        If x = 32 Then
        y = y + 1
        x = 1
        
        End If

    Next i
    
    
    For i = 496 To 504
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 494, 40).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 494, 40).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 494, 40).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 494, 40).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
      
'''''''CA2000 25point&SJpoint�P�x�f�[�^�擾 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy����").Cells(i - 452, 40).Value = "" '�G���[�l���󔒂�Excel����
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy����").Cells(i - 452, 40).Value = "" '�G���[�l���󔒂�Excel����

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy����").Cells(i - 452, 40).Value = "" '�G���[�l���󔒂�Excel����
           
        Else
            Sheets("xy����").Cells(i - 452, 40).Value = SpotData(i) '�擾����Cx�f�[�^��Excel�ɓ���
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim AAA, BBB, CCC, DDD, EEE, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double
Dim FFF, GGG, HHH, III, JJJ, KKK As Double

    
'''''''''''''''''''''''''''''''''496MaxJND�v�Z'''''''''''''''''''''''''''''''''''''''''
      
'''''�ΏېF�xCxCy����'''''
      
       Sheets("JND�v�Z").Range("H9").Value = 0 '�O��MaxJND����
       Sheets("JND�v�Z").Range("I17:J512").Value = 0 '�O��ΏېF�xCxCy����
       
       AAA = 16 '�s��
       BBB = 31 '��

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

Dim CounterB, CounterC As Long

      EE = 0
      
      For CounterB = 7 To BBB + 6
       For CounterC = 2 To AAA + 1
   
       
        DD = Sheets("JND�v�Z").Range("H9").Value '����MaxJND
  
                
        If DD < EE Then '����MaxJND VS �ǉ�MaxJND �召��r

        Sheets("JND�v�Z").Range("H9").Value = Sheets("JND�v�Z").Range("K14").Value '496MaxJND����
        Sheets("JND�v�Z").Range("l9").Value = Sheets("JND�v�Z").Range("l16").Value 'MaxJND �(�s)����
        Sheets("JND�v�Z").Range("m9").Value = Sheets("JND�v�Z").Range("m16").Value 'MaxJND �(��)����
        Sheets("JND�v�Z").Range("n9").Value = Sheets("JND�v�Z").Range("l14").Value 'MaxJND �Ώ�(�s)����
        Sheets("JND�v�Z").Range("o9").Value = Sheets("JND�v�Z").Range("m14").Value 'MaxJND �Ώ�(��)����
        
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
  
        FF = Sheets("JND�v�Z").Range("l9").Value + 1 'MaxJND �(�s)
        GG = Sheets("JND�v�Z").Range("m9").Value + 6 'MaxJND �(��)
        HH = FF + 16
        II = Sheets("JND�v�Z").Range("n9").Value + 1 'MaxJND �Ώ�(�s)
        JJ = Sheets("JND�v�Z").Range("o9").Value + 6 'MaxJND �Ώ�(��)
        KK = II + 16
        
        
        Sheets("xy����").Cells(FF, GG).Interior.ColorIndex = 3 'x �point�ڈ�
        Sheets("xy����").Cells(HH, GG).Interior.ColorIndex = 3 'y �point�ڈ�
        
        Sheets("xy����").Cells(II, JJ).Interior.ColorIndex = 3 'x �Ώ�point�ڈ�
        Sheets("xy����").Cells(KK, JJ).Interior.ColorIndex = 3 'y �Ώ�point�ڈ�
        
        
        Sheets("JND�v�Z").Range("p9").Value = Sheets("u'v'�ϊ�").Cells(FF, GG).Value 'x ��F�xCu'����
        Sheets("JND�v�Z").Range("q9").Value = Sheets("u'v'�ϊ�").Cells(HH, GG).Value 'y ��F�xCv'����
        
        Sheets("JND�v�Z").Range("r9").Value = Sheets("u'v'�ϊ�").Cells(II, JJ).Value 'x �ΏېF�xCu'����
        Sheets("JND�v�Z").Range("s9").Value = Sheets("u'v'�ϊ�").Cells(KK, JJ).Value 'y �ΏېF�xCv'����
        
        Sheets("��u'v'�v�Z").Range("h16").Value = Sheets("JND�v�Z").Range("p9").Value '��u'v'�v�Z�V�[�g�ɊCu'���
        Sheets("��u'v'�v�Z").Range("i16").Value = Sheets("JND�v�Z").Range("q9").Value '��u'v'�v�Z�V�[�g�ɊCv'���
        Sheets("��u'v'�v�Z").Range("h17").Value = Sheets("JND�v�Z").Range("r9").Value '��u'v'�v�Z�V�[�g�ɑΏ�Cu'���
        Sheets("��u'v'�v�Z").Range("i17").Value = Sheets("JND�v�Z").Range("s9").Value '��u'v'�v�Z�V�[�g�ɑΏ�Cv'���
        
        Sheets("JND�v�Z").Range("j9:k9").Value = Sheets("��u'v'�v�Z").Range("j17:k17").Value '��u'v'�v�Z���ʓ���
        
        
''''''CA2000�f�[�^�����p����''''''''

        Sheets("xy����").Range("aa56").Value = Sheets("xy����").Cells(FF, GG).Value 'x ��F�xCx����
        Sheets("xy����").Range("ab56").Value = Sheets("xy����").Cells(HH, GG).Value 'y ��F�xCy����
        Sheets("xy����").Range("ac56").Value = Sheets("xy����").Cells(II, JJ).Value 'x �ΏېF�xCx����
        Sheets("xy����").Range("ad56").Value = Sheets("xy����").Cells(KK, JJ).Value 'y �ΏېF�xCy����
      

    
'''''''''''''''''''''''''''''''''496Max��u'v�v�Z'''''''''''''''''''''''''''''''''''''''''


      Sheets("��u'v'�v�Z").Range("H9:H10").Value = 0 '�O��Max��u'v'����
      Sheets("��u'v'�v�Z").Range("h17:i512").Value = 0 '�O��F�xu'v'����
  
  
       AAA = 16 '�s��
       BBB = 31 '��
       
       
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
        
        Sheets("xy����").Cells(FF, GG).Interior.ColorIndex = 4 ' u' �point�ڈ�
        Sheets("xy����").Cells(HH, GG).Interior.ColorIndex = 4 ' v' �point�ڈ�
        
        Sheets("xy����").Cells(II, JJ).Interior.ColorIndex = 4 ' u' �Ώ�point�ڈ�
        Sheets("xy����").Cells(KK, JJ).Interior.ColorIndex = 4 ' v' �Ώ�point�ڈ�
        
        
        
 ''' u'<=v' '''
        FFF = Sheets("��u'v'�v�Z").Range("k10").Value + 1 'Max��u'v' �(�s)
        GGG = Sheets("��u'v'�v�Z").Range("l10").Value + 6 'Max��u'v' �(��)
        HHH = FFF + 16
        III = Sheets("��u'v'�v�Z").Range("m10").Value + 1 'Max��u'v' �Ώ�(�s)
        JJJ = Sheets("��u'v'�v�Z").Range("n10").Value + 6 'Max��u'v' �Ώ�(��)
        KKK = III + 16
        
        Sheets("xy����").Cells(FFF, GGG).Interior.ColorIndex = 5 ' u' �point�ڈ�
        Sheets("xy����").Cells(HHH, GGG).Interior.ColorIndex = 5 ' v' �point�ڈ�
        
        Sheets("xy����").Cells(III, JJJ).Interior.ColorIndex = 5 ' u' �Ώ�point�ڈ�
        Sheets("xy����").Cells(KKK, JJJ).Interior.ColorIndex = 5 ' v' �Ώ�point�ڈ�
        
        
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
        
        
''''''''''CA2000�f�[�^�����pCxCy����'''''''''''''''''
        
        Sheets("xy����").Range("l56").Value = Sheets("xy����").Cells(FF, GG).Value ' u' �pointCx����
        Sheets("xy����").Range("m56").Value = Sheets("xy����").Cells(HH, GG).Value ' v' �pointCy����
        Sheets("xy����").Range("n56").Value = Sheets("xy����").Cells(II, JJ).Value ' u' �Ώ�pointCx����
        Sheets("xy����").Range("o56").Value = Sheets("xy����").Cells(KK, JJ).Value ' v' �Ώ�pointCy����
        
        Sheets("xy����").Range("s56").Value = Sheets("xy����").Cells(FFF, GGG).Value ' u' �pointCx����
        Sheets("xy����").Range("t56").Value = Sheets("xy����").Cells(HHH, GGG).Value ' v' �pointCy����
        Sheets("xy����").Range("u56").Value = Sheets("xy����").Cells(III, JJJ).Value ' u' �Ώ�pointCx����
        Sheets("xy����").Range("v56").Value = Sheets("xy����").Cells(KKK, JJJ).Value ' v' �Ώ�pointCy����
  
'''''''''''���茋�ʕ\���v���O����'''''''''''''''''

Dim MX As Double

Sheets("xy����").Cells(55, 8).Value = TextBox1 'File��(TextBox1)��Excel�ɓ���
MX = Format(Now, "yyyymmddhhnnss") '����

Sheets("xy����").Cells(55, 7).Value = MX '���������Excel�ɓ���

'��u'v'(u'����)�f�[�^�̐���
Sheets("xy����").Range(Sheets("xy����").Cells(55, 9), Sheets("xy����").Cells(55, 15)).Value = Sheets("��u'v'�v�Z").Range("h9:n9").Value

'��u'v'(v'����)�f�[�^�̐���
Sheets("xy����").Range(Sheets("xy����").Cells(55, 16), Sheets("xy����").Cells(55, 22)).Value = Sheets("��u'v'�v�Z").Range("h10:n10").Value

'MaxJND�f�[�^�̐���
Sheets("xy����").Range(Sheets("xy����").Cells(55, 23), Sheets("xy����").Cells(55, 30)).Value = Sheets("JND�v�Z").Range("h9:o9").Value


'''MaxJND UserForm�ɕ\�L'''

'���S�l

Label37 = Round(Sheets("xy����").Range("an6").Value, 0)
Label38 = Round(Sheets("xy����").Range("ao6").Value, 4)
Label39 = Round(Sheets("xy����").Range("ap6").Value, 4)



'MaxJND
Label3 = Round(Sheets("JND�v�Z").Range("h9").Value, 2)

'��u'v'
Label7 = Round(Sheets("JND�v�Z").Range("j9").Value, 4) + Round(Sheets("JND�v�Z").Range("k9").Value, 4)

'position
Label9 = Sheets("JND�v�Z").Range("l9").Value & "," & Sheets("JND�v�Z").Range("m9").Value & " -> " & Sheets("JND�v�Z").Range("n9").Value & "," & Sheets("JND�v�Z").Range("o9").Value

'9point�P�x����
Label45 = Round(Sheets("xy����").Range("ao13").Value, 3)

'9point��u'v'
Label47 = Round(Sheets("xy����").Range("ao14").Value, 4)


'''Max��u'v'(u'����) UserForm�ɕ\�L'''
'JND
Label11 = Round(Sheets("��u'v'�v�Z").Range("j9").Value, 2)

'Max��u'v'
Label15 = Round(Sheets("��u'v'�v�Z").Range("h9").Value, 4)

'position
Label16 = Sheets("��u'v'�v�Z").Range("k9").Value & "," & Sheets("��u'v'�v�Z").Range("l9").Value & " -> " & Sheets("��u'v'�v�Z").Range("m9").Value & "," & Sheets("��u'v'�v�Z").Range("n9").Value


'''Max��u'v'(v'����) UserForm�ɕ\�L'''
'JND
Label19 = Round(Sheets("��u'v'�v�Z").Range("j10").Value, 2)

'Max��u'v'
Label23 = Round(Sheets("��u'v'�v�Z").Range("h10").Value, 4)

'position
Label24 = Sheets("��u'v'�v�Z").Range("k10").Value & "," & Sheets("��u'v'�v�Z").Range("l10").Value & " -> " & Sheets("��u'v'�v�Z").Range("m10").Value & "," & Sheets("��u'v'�v�Z").Range("n10").Value


'''''''''''����摜�̕ۑ�''''''''''''''

'   Dim strName As String
'    Dim TabIndex As Long
'    Dim ObjectIndex As Long

'    TabIndex = 1 'CA2000 �[���J���[
'  ObjectIndex = 1 'CA2000 �P�x�摜�I��
'    strName = TextBox2 & MX & "_" & TextBox1 & "Ey.bmp" '�摜��file��
'    Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '�摜�ۑ�
    
 '   TabIndex = 1 'CA2000 �[���J���[
 '   ObjectIndex = 2 'CA2000 Cx�摜�I��
 '  strName = TextBox2 & MX & "_" & TextBox1 & "Cx.bmp" '�摜��file��
 '   Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '�摜�ۑ�

 '   TabIndex = 1 'CA2000 �[���J���[
 '   ObjectIndex = 3 'CA2000 Cy�摜�I��
 '   strName = TextBox2 & MX & "_" & TextBox1 & "Cy.bmp" '�摜��file��
 '   Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '�摜�ۑ�

 '   Set Image1.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Ey.bmp") '���[�U�[�t�@�[���ɋP�x���z�\��
 '   Set Image2.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Cx.bmp") '���[�U�[�t�@�[����Cx���z�\��
 '   Set Image3.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Cy.bmp") '���[�U�[�t�@�[����Cy���z�\��
    
    
'''''''''����f�[�^�̕ۑ�'''''''''''''

'�yCA2000�z.xls�t�@�C�����N������
    Workbooks.Open Filename:= _
        "D:\CA2000data\�yCA2000_08model�z.xls"

'�yCA2000�z.xls�t�@�C���ɐ�����������f�[�^���R�s�[
ThisWorkbook.Sheets("xy����").Range("AN2:AP10").Copy ActiveWorkbook.Worksheets(1).Range("Q10")
ThisWorkbook.Sheets("xy����").Range("G55:AD56").Copy ActiveWorkbook.Worksheets(1).Range("A4")
ThisWorkbook.Sheets("xy����").Range("E1:AL50").Copy ActiveWorkbook.Worksheets(1).Range("AC1")
ThisWorkbook.Sheets("xy����").Range("AN53:AP68").Copy ActiveWorkbook.Worksheets(1).Range("Q21") '�ǉ�071217
ThisWorkbook.Sheets("xy����").Range("AN69:AN72").Copy ActiveWorkbook.Worksheets(1).Range("Q38") '�ǉ�071217
ThisWorkbook.Sheets("xy����").Range("AN73:AN76").Copy ActiveWorkbook.Worksheets(1).Range("Q43") '�ǉ�071217


'ActiveWorkbook.Sheets(1).Range("A6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Ey.bmp").Select
'ThisWorkbook.Sheets("xy����").Range("G6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Cx.bmp").Select
'ThisWorkbook.Sheets("xy����").Range("M6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Cy.bmp").Select

'�yCA2000�z.xls�t�@�C���𖼑O��ύX���ĕۑ�
ActiveWorkbook.SaveAs TextBox2 & MX & "_" & TextBox1 & ".xls"

'�ۑ�����Excel�����
Workbooks(MX & "_" & TextBox1 & ".xls").Close Savechanges:=False

    Set SpotCond = Nothing
    Set myappl = Nothing
    
    Exit Sub
      
'''''''�G���[����'''''''''
    
�G���[������:

 MsgBox "�ECA-S20w�̑����ʂ��N�����Ă��܂��񂩁H" & vbCrLf & "�E496Spot���w�肵�Ă��܂����H"
 
           
End Sub


'''''''''''CA2000�����Z���p�v���O����'''''''''''''''''


Private Sub CommandButton2_Click()

Dim CSEy As Double 'CS1000�Z���P�x
Dim CSCx As Double 'CS1000�Z��Cx
Dim CSCy As Double 'CS1000�Z��Cy

CSEy = TextBox3 'CS1000�Z���P�x�擾
CSCx = TextBox4 'CS1000�Z��Cx�擾
CSCy = TextBox5 'CS1000�Z��Cy�擾

On Error GoTo �G���[������

 Set myappl = CreateObject("CAS20W.Application") 'CA2000�\�t�g�Ƃ̃����N


''''''''�Z���O����''''''''''

    Dim Ret As Long
    Ret = myappl.Measure() 'CA2000 ����v���O����
    
    Do
    'Ret = MsgBox("Cancel?", MsgBoxStyle.OKCancel, "")
    'If Ret = MsgBoxResult.OK Then
    'If myappl.MeasureCancel() = 0 Then
    'Exit Do
    'End If
    'Else
    
    If myappl.PollingMeasure() = 0 Then
    Exit Do
    End If
    'End If
    'Loop Until Ret = MsgBoxResult.OK
    Loop Until Ret = -1
    
    

''''''''�Z���O�P�x�E�F�x�l�擾''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Dim SpotCond As Object  'CA2000 �X�|�b�g����f�[�^�擾�v���O����
    
    Dim CalEy As Double '�Z���O�P�x
    Dim CalCx As Double '�Z���OCx
    Dim CalCy As Double '�Z���OCy
    
    Ret = myappl.SelectData()

    Set SpotCond = myappl.GetSpotCondition() 'CA2000 Spot�ݒ�������擾
    
    Dim SpotCount As Long
    SpotCount = SpotCond.GetSpotCount() 'CA2000 Spot�����擾
    
    Dim SpotData(100000) As Single 'CA2000 Spot�����̔z��m��
    
    
    Element = 3  'Spot �P�x�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
    
      CalEy = SpotData(500) '�Z���O�P�x�擾


    Element = 4  'Spot Cx�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCx = SpotData(500) '�Z���OCx�擾


    Element = 5  'Spot Cy�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCy = SpotData(500) '�Z���OCy�擾

   
    
'''''''''�Z���l����''''''''''

Dim Condition As Object

Dim UserCalibData As Object

 Set Condition = myappl.GetMeasurementCondition() 'CA2000 ��������̎擾

 Set UserCalibData = Condition.GetUserCalibrationData() '���[�U�[�Z���W���I�u�W�F�N�g�̎擾

UserCalibData.CalibrationType = 2 'CA2000 �P�_�Z���I��


UserCalibData.WLv_before = CalEy '�Z���O�P�x
UserCalibData.WLv_after = CSEy 'CS1000�Z���P�x

UserCalibData.Wx_before = CalCx '�Z���OCx
UserCalibData.Wx_after = CSCx 'CS1000�Z��Cx

UserCalibData.Wy_before = CalCy '�Z���OCy
UserCalibData.Wy_after = CSCy 'CS1000�Z��Cy

Ret = Condition.SetUserCalibrationData(UserCalibData)



''''''''�Z������''''''''

   Ret = myappl.Measure() 'CA2000 ����v���O����
    
    Do
    'Ret = MsgBox("Cancel?", MsgBoxStyle.OKCancel, "")
    'If Ret = MsgBoxResult.OK Then
    'If myappl.MeasureCancel() = 0 Then
    'Exit Do
    'End If
    'Else
    
    If myappl.PollingMeasure() = 0 Then
    Exit Do
    End If
    'End If
    'Loop Until Ret = MsgBoxResult.OK
    Loop Until Ret = -1
    
    

'''''''''�Z����P�x�E�F�x�m�F'''''''''''
 
    Ret = myappl.SelectData()

  
    Element = 3 'Spot �P�x�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
      CalEy = SpotData(500)
      Label37 = Round(CalEy, 0) '�Z����P�x�擾
                  
    
    Element = 4 'Spot Cx�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCx = SpotData(500)
    Label38 = Round(CalCx, 4) '�Z����Cx�擾
    
    
    Element = 5 'Spot Cy�f�[�^�̎擾
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCy = SpotData(500)
    Label39 = Round(CalCy, 4) '�Z����Cy�擾


    Set SpotCond = Nothing
    Set myappl = Nothing
    
    MsgBox "�Z������"  '����ē����b�Z�[�W
    
    Exit Sub
    
    
        
'''''''�G���[����'''''''''
    
�G���[������:

 MsgBox "�ECA-S20w�̑����ʂ��N�����Ă��܂��񂩁H"
    
    
End Sub

