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


On Error GoTo エラー処理へ


 Set myappl = CreateObject("CAS20W.Application") 'CA2000ソフトとのリンク


'''''''''''測定用プログラム''''''''''''''

    Dim Ret As Long
    Ret = myappl.Measure() 'CA2000 測定プログラム
    
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
    


''''''''''496Spotデータ取得プログラム''''''''''''

    Dim SpotCond As Object
    Dim x, y As Long

    Ret = myappl.SelectData()

    Set SpotCond = myappl.GetSpotCondition() 'CA2000 Spot設定条件を取得
    
    Dim SpotCount As Long
    SpotCount = SpotCond.GetSpotCount() 'CA2000 Spot個数を取得
    
    Dim SpotData(1000000) As Single 'CA2000 Spot個数分の配列確保
    
        
 '''''''CA2000 496Cxデータ取得'''''''
    
    Element = 4 'Spot Cxデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    x = 1
    y = 1
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(y + 1, x + 6).Value = ""  'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(y + 1, x + 6).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(y + 1, x + 6).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(y + 1, x + 6).Value = SpotData(i) '取得したCxデータをExcelに入力
            
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
            Sheets("xy入力").Cells(i - 494, 41).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 494, 41).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 494, 41).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 494, 41).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
    
    
'''''''CA2000 25point&SJpoint Cxデータ取得 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(i - 452, 41).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 452, 41).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 452, 41).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 452, 41).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    
    
    
 '''''''CA2000 496Cyデータ取得'''''''
    
        Element = 5 'Spot Cyデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    x = 1
    y = 1
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(y + 17, x + 6).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(y + 17, x + 6).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(y + 17, x + 6).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(y + 17, x + 6).Value = SpotData(i) '取得したCyデータをExcelに入力
            
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
            Sheets("xy入力").Cells(i - 494, 42).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 494, 42).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 494, 42).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 494, 42).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
    
 '''''''CA2000 25point&SJpoint Cyデータ取得 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(i - 452, 42).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 452, 42).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 452, 42).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 452, 42).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    

'''''''CA2000 496輝度データ取得'''''''

       Element = 3 'Spot 輝度データの取得
       Ret = myappl.GetSpotData(Element, SpotData)
       x = 1
       y = 1
    
    
        For i = 0 To 495
           
        If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(y + 33, x + 6).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(y + 33, x + 6).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(y + 33, x + 6).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(y + 33, x + 6).Value = SpotData(i) '取得した輝度データをExcelに入力
            
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
            Sheets("xy入力").Cells(i - 494, 40).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 494, 40).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 494, 40).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 494, 40).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
      
'''''''CA2000 25point&SJpoint輝度データ取得 071217'''''''
          
    For i = 505 To 530
    
            If SpotData(i) < -3E+38 Then
            'Over-error pixel
            Sheets("xy入力").Cells(i - 452, 40).Value = "" 'エラー値を空白でExcel入力
            
        ElseIf SpotData(i) < -2E+38 And Value >= -3E+38 Then
            'Under-error pixel
            Sheets("xy入力").Cells(i - 452, 40).Value = "" 'エラー値を空白でExcel入力

            
        ElseIf SpotData(i) < -1E+38 Then
            Sheets("xy入力").Cells(i - 452, 40).Value = "" 'エラー値を空白でExcel入力
           
        Else
            Sheets("xy入力").Cells(i - 452, 40).Value = SpotData(i) '取得したCxデータをExcelに入力
            
        End If
        
    Next i
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim AAA, BBB, CCC, DDD, EEE, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double
Dim FFF, GGG, HHH, III, JJJ, KKK As Double

    
'''''''''''''''''''''''''''''''''496MaxJND計算'''''''''''''''''''''''''''''''''''''''''
      
'''''対象色度CxCy入力'''''
      
       Sheets("JND計算").Range("H9").Value = 0 '前回MaxJND消去
       Sheets("JND計算").Range("I17:J512").Value = 0 '前回対象色度CxCy消去
       
       AAA = 16 '行数
       BBB = 31 '列数

     For AA = 7 To BBB + 6
       For CCC = 2 To AAA + 1
       
       CC = AAA * (AA - 6) + CCC - 1 + (16 - AAA)
       
        Sheets("JND計算").Cells(CC, 9).Value = Sheets("xy入力").Cells(CCC, AA).Value 'x 対象色度Cx
        Sheets("JND計算").Cells(CC, 10).Value = Sheets("xy入力").Cells(CCC + 16, AA).Value 'y 対象色度Cy
        Sheets("JND計算").Cells(CC, 12).Value = CCC - 1 'y 対象(行)入力
        Sheets("JND計算").Cells(CC, 13).Value = AA - 6 'y 対象(列)入力
        
       Next CCC
     Next AA

      
'''''基準CxCy入力とMaxJND値・point入力'''''

Dim CounterB, CounterC As Long

      EE = 0
      
      For CounterB = 7 To BBB + 6
       For CounterC = 2 To AAA + 1
   
       
        DD = Sheets("JND計算").Range("H9").Value '現状MaxJND
  
                
        If DD < EE Then '現状MaxJND VS 追加MaxJND 大小比較

        Sheets("JND計算").Range("H9").Value = Sheets("JND計算").Range("K14").Value '496MaxJND入力
        Sheets("JND計算").Range("l9").Value = Sheets("JND計算").Range("l16").Value 'MaxJND 基準(行)入力
        Sheets("JND計算").Range("m9").Value = Sheets("JND計算").Range("m16").Value 'MaxJND 基準(列)入力
        Sheets("JND計算").Range("n9").Value = Sheets("JND計算").Range("l14").Value 'MaxJND 対象(行)入力
        Sheets("JND計算").Range("o9").Value = Sheets("JND計算").Range("m14").Value 'MaxJND 対象(列)入力
        
        End If
            

        Sheets("JND計算").Cells(16, 9).Value = Sheets("xy入力").Cells(CounterC, CounterB).Value 'x 基準色度Cx入力
        Sheets("JND計算").Cells(16, 10).Value = Sheets("xy入力").Cells(CounterC + 16, CounterB).Value 'y 基準色度Cy入力
        Sheets("JND計算").Cells(16, 12).Value = CounterC - 1 'x 基準point入力
        Sheets("JND計算").Cells(16, 13).Value = CounterB - 6 'y 基準point入力
        
        
        EE = Sheets("JND計算").Range("K14").Value '追加MaxJND
 
                  
       Next CounterC
      Next CounterB
      
      
'''''Max point目印'''''

        Sheets("xy入力").Range("E2:AK33").Interior.ColorIndex = 0 '前回目印消去
  
        FF = Sheets("JND計算").Range("l9").Value + 1 'MaxJND 基準(行)
        GG = Sheets("JND計算").Range("m9").Value + 6 'MaxJND 基準(列)
        HH = FF + 16
        II = Sheets("JND計算").Range("n9").Value + 1 'MaxJND 対象(行)
        JJ = Sheets("JND計算").Range("o9").Value + 6 'MaxJND 対象(列)
        KK = II + 16
        
        
        Sheets("xy入力").Cells(FF, GG).Interior.ColorIndex = 3 'x 基準point目印
        Sheets("xy入力").Cells(HH, GG).Interior.ColorIndex = 3 'y 基準point目印
        
        Sheets("xy入力").Cells(II, JJ).Interior.ColorIndex = 3 'x 対象point目印
        Sheets("xy入力").Cells(KK, JJ).Interior.ColorIndex = 3 'y 対象point目印
        
        
        Sheets("JND計算").Range("p9").Value = Sheets("u'v'変換").Cells(FF, GG).Value 'x 基準色度Cu'入力
        Sheets("JND計算").Range("q9").Value = Sheets("u'v'変換").Cells(HH, GG).Value 'y 基準色度Cv'入力
        
        Sheets("JND計算").Range("r9").Value = Sheets("u'v'変換").Cells(II, JJ).Value 'x 対象色度Cu'入力
        Sheets("JND計算").Range("s9").Value = Sheets("u'v'変換").Cells(KK, JJ).Value 'y 対象色度Cv'入力
        
        Sheets("⊿u'v'計算").Range("h16").Value = Sheets("JND計算").Range("p9").Value '⊿u'v'計算シートに基準Cu'代入
        Sheets("⊿u'v'計算").Range("i16").Value = Sheets("JND計算").Range("q9").Value '⊿u'v'計算シートに基準Cv'代入
        Sheets("⊿u'v'計算").Range("h17").Value = Sheets("JND計算").Range("r9").Value '⊿u'v'計算シートに対象Cu'代入
        Sheets("⊿u'v'計算").Range("i17").Value = Sheets("JND計算").Range("s9").Value '⊿u'v'計算シートに対象Cv'代入
        
        Sheets("JND計算").Range("j9:k9").Value = Sheets("⊿u'v'計算").Range("j17:k17").Value '⊿u'v'計算結果入力
        
        
''''''CA2000データ整理用入力''''''''

        Sheets("xy入力").Range("aa56").Value = Sheets("xy入力").Cells(FF, GG).Value 'x 基準色度Cx入力
        Sheets("xy入力").Range("ab56").Value = Sheets("xy入力").Cells(HH, GG).Value 'y 基準色度Cy入力
        Sheets("xy入力").Range("ac56").Value = Sheets("xy入力").Cells(II, JJ).Value 'x 対象色度Cx入力
        Sheets("xy入力").Range("ad56").Value = Sheets("xy入力").Cells(KK, JJ).Value 'y 対象色度Cy入力
      

    
'''''''''''''''''''''''''''''''''496Max⊿u'v計算'''''''''''''''''''''''''''''''''''''''''


      Sheets("⊿u'v'計算").Range("H9:H10").Value = 0 '前回Max⊿u'v'消去
      Sheets("⊿u'v'計算").Range("h17:i512").Value = 0 '前回色度u'v'消去
  
  
       AAA = 16 '行数
       BBB = 31 '列数
       
       
'''''対象色度u'v'入力'''''

      For AA = 7 To BBB + 6
       For CCC = 2 To AAA + 1
        CC = AAA * (AA - 6) + CCC - 1 + (16 - AAA)
        Sheets("⊿u'v'計算").Cells(CC, 8).Value = Sheets("u'v'変換").Cells(CCC, AA).Value 'u 対象色度u'入力
        Sheets("⊿u'v'計算").Cells(CC, 9).Value = Sheets("u'v'変換").Cells(CCC + 16, AA).Value 'v 対象色度v'入力
        Sheets("⊿u'v'計算").Cells(CC, 12).Value = CCC - 1 'u 対象(行)入力
        Sheets("⊿u'v'計算").Cells(CC, 13).Value = AA - 6 'v 対象(列)入力
       Next CCC
      Next AA
      
      
'''''基準u'v'入力とMax⊿u'v'値・point入力'''''
        
        
       EE = 0
       EEE = 0
       
  For CounterB = 7 To BBB + 6
  
   For CounterC = 2 To AAA + 1
   
       CC = AAA * (CounterB - 6) + CounterC - 1 + (16 - AAA)
      
        DDD = Sheets("⊿u'v'計算").Range("H9").Value 'u'>v' 現状Max⊿u'v'
        DD = Sheets("⊿u'v'計算").Range("H10").Value 'u'<=v' 現状Max⊿u'v'
        

   ''' u'<=v' Max⊿u'v'値・point入力'''
        
        If DD < EE Then

        Sheets("⊿u'v'計算").Range("H10").Value = Sheets("⊿u'v'計算").Range("K14").Value '496Max⊿u'v'入力
        Sheets("⊿u'v'計算").Range("k10").Value = Sheets("⊿u'v'計算").Range("l16").Value 'Max⊿u'v' 基準(行)入力
        Sheets("⊿u'v'計算").Range("l10").Value = Sheets("⊿u'v'計算").Range("m16").Value 'Max⊿u'v' 基準(列)入力
        Sheets("⊿u'v'計算").Range("m10").Value = Sheets("⊿u'v'計算").Range("l14").Value 'Max⊿u'v' 対象(行)入力
        Sheets("⊿u'v'計算").Range("n10").Value = Sheets("⊿u'v'計算").Range("m14").Value 'Max⊿u'v' 対象(列)入力
        
        End If
                
                
        
   ''' u'>v' Max⊿u'v'値・point入力'''
        
        If DDD < EEE Then

        Sheets("⊿u'v'計算").Range("H9").Value = Sheets("⊿u'v'計算").Range("g14").Value '496Max⊿u'v'入力
        Sheets("⊿u'v'計算").Range("k9").Value = Sheets("⊿u'v'計算").Range("l16").Value 'Max⊿u'v' 基準(行)入力
        Sheets("⊿u'v'計算").Range("l9").Value = Sheets("⊿u'v'計算").Range("m16").Value 'Max⊿u'v' 基準(列)入力
        Sheets("⊿u'v'計算").Range("m9").Value = Sheets("⊿u'v'計算").Range("h14").Value 'Max⊿u'v' 対象(行)入力
        Sheets("⊿u'v'計算").Range("n9").Value = Sheets("⊿u'v'計算").Range("i14").Value 'Max⊿u'v' 対象(列)入力
        
        End If
        
        
   '''基準u'v'入力'''
        
        Sheets("⊿u'v'計算").Cells(16, 8).Value = Sheets("u'v'変換").Cells(CounterC, CounterB).Value 'x 基準色度u'入力
        Sheets("⊿u'v'計算").Cells(16, 9).Value = Sheets("u'v'変換").Cells(CounterC + 16, CounterB).Value 'y 基準色度v'入力
        Sheets("⊿u'v'計算").Cells(16, 12).Value = CounterC - 1 'x 基準(行)入力
        Sheets("⊿u'v'計算").Cells(16, 13).Value = CounterB - 6 'y 基準(列)入力
        
        EEE = Sheets("⊿u'v'計算").Range("g14").Value ' u'>v' 追加Max⊿u'v'
        EE = Sheets("⊿u'v'計算").Range("k14").Value ' u'<=v' 追加Max⊿u'v'

    Next CounterC
   Next CounterB
   
   '''''Max point目印'''''
  
        Sheets("u'v'変換").Range("E2:Ak33").Interior.ColorIndex = 0 '前回目印消去
        
  ''' u'>v' '''
        FF = Sheets("⊿u'v'計算").Range("k9").Value + 1 'Max⊿u'v' 基準(行)
        GG = Sheets("⊿u'v'計算").Range("l9").Value + 6 'Max⊿u'v' 基準(列)
        HH = FF + 16
        II = Sheets("⊿u'v'計算").Range("m9").Value + 1 'Max⊿u'v' 対象(行)
        JJ = Sheets("⊿u'v'計算").Range("n9").Value + 6 'Max⊿u'v' 対象(列)
        KK = II + 16
        
        Sheets("xy入力").Cells(FF, GG).Interior.ColorIndex = 4 ' u' 基準point目印
        Sheets("xy入力").Cells(HH, GG).Interior.ColorIndex = 4 ' v' 基準point目印
        
        Sheets("xy入力").Cells(II, JJ).Interior.ColorIndex = 4 ' u' 対象point目印
        Sheets("xy入力").Cells(KK, JJ).Interior.ColorIndex = 4 ' v' 対象point目印
        
        
        
 ''' u'<=v' '''
        FFF = Sheets("⊿u'v'計算").Range("k10").Value + 1 'Max⊿u'v' 基準(行)
        GGG = Sheets("⊿u'v'計算").Range("l10").Value + 6 'Max⊿u'v' 基準(列)
        HHH = FFF + 16
        III = Sheets("⊿u'v'計算").Range("m10").Value + 1 'Max⊿u'v' 対象(行)
        JJJ = Sheets("⊿u'v'計算").Range("n10").Value + 6 'Max⊿u'v' 対象(列)
        KKK = III + 16
        
        Sheets("xy入力").Cells(FFF, GGG).Interior.ColorIndex = 5 ' u' 基準point目印
        Sheets("xy入力").Cells(HHH, GGG).Interior.ColorIndex = 5 ' v' 基準point目印
        
        Sheets("xy入力").Cells(III, JJJ).Interior.ColorIndex = 5 ' u' 対象point目印
        Sheets("xy入力").Cells(KKK, JJJ).Interior.ColorIndex = 5 ' v' 対象point目印
        
        
'''''Max⊿u'v'⇒JND変換'''''

  ''' u'>v' '''

        Sheets("⊿u'v'計算").Range("o9").Value = Sheets("xy入力").Cells(FF, GG).Value ' u' 基準pointCx入力
        Sheets("⊿u'v'計算").Range("p9").Value = Sheets("xy入力").Cells(HH, GG).Value ' v' 基準pointCy入力
        Sheets("⊿u'v'計算").Range("q9").Value = Sheets("xy入力").Cells(II, JJ).Value ' u' 対象pointCx入力
        Sheets("⊿u'v'計算").Range("r9").Value = Sheets("xy入力").Cells(KK, JJ).Value ' v' 対象pointCy入力

        Sheets("JND計算").Range("I16").Value = Sheets("⊿u'v'計算").Range("o9").Value 'JND計算シートに基準Cx代入
        Sheets("JND計算").Range("j16").Value = Sheets("⊿u'v'計算").Range("p9").Value 'JND計算シートに基準Cy代入
        Sheets("JND計算").Range("I17").Value = Sheets("⊿u'v'計算").Range("q9").Value 'JND計算シートに対象Cx代入
        Sheets("JND計算").Range("j17").Value = Sheets("⊿u'v'計算").Range("r9").Value 'JND計算シートに対象Cy代入
        
        Sheets("⊿u'v'計算").Range("j9").Value = Sheets("JND計算").Range("k17").Value 'JND計算結果入力

 ''' u'<=v' '''
        Sheets("⊿u'v'計算").Range("o10").Value = Sheets("xy入力").Cells(FFF, GGG).Value ' u' 基準pointCx入力
        Sheets("⊿u'v'計算").Range("p10").Value = Sheets("xy入力").Cells(HHH, GGG).Value ' v' 基準pointCy入力
        Sheets("⊿u'v'計算").Range("q10").Value = Sheets("xy入力").Cells(III, JJJ).Value ' u' 対象pointCx入力
        Sheets("⊿u'v'計算").Range("r10").Value = Sheets("xy入力").Cells(KKK, JJJ).Value ' v' 対象pointCy入力
        
        Sheets("JND計算").Range("I16").Value = Sheets("⊿u'v'計算").Range("o10").Value 'JND計算シートに基準Cx代入
        Sheets("JND計算").Range("j16").Value = Sheets("⊿u'v'計算").Range("p10").Value 'JND計算シートに基準Cy代入
        Sheets("JND計算").Range("I17").Value = Sheets("⊿u'v'計算").Range("q10").Value 'JND計算シートに対象Cx代入
        Sheets("JND計算").Range("j17").Value = Sheets("⊿u'v'計算").Range("r10").Value 'JND計算シートに対象Cy代入
        
        Sheets("⊿u'v'計算").Range("j10").Value = Sheets("JND計算").Range("k17").Value 'JND計算結果入力
        
        
''''''''''CA2000データ整理用CxCy入力'''''''''''''''''
        
        Sheets("xy入力").Range("l56").Value = Sheets("xy入力").Cells(FF, GG).Value ' u' 基準pointCx入力
        Sheets("xy入力").Range("m56").Value = Sheets("xy入力").Cells(HH, GG).Value ' v' 基準pointCy入力
        Sheets("xy入力").Range("n56").Value = Sheets("xy入力").Cells(II, JJ).Value ' u' 対象pointCx入力
        Sheets("xy入力").Range("o56").Value = Sheets("xy入力").Cells(KK, JJ).Value ' v' 対象pointCy入力
        
        Sheets("xy入力").Range("s56").Value = Sheets("xy入力").Cells(FFF, GGG).Value ' u' 基準pointCx入力
        Sheets("xy入力").Range("t56").Value = Sheets("xy入力").Cells(HHH, GGG).Value ' v' 基準pointCy入力
        Sheets("xy入力").Range("u56").Value = Sheets("xy入力").Cells(III, JJJ).Value ' u' 対象pointCx入力
        Sheets("xy入力").Range("v56").Value = Sheets("xy入力").Cells(KKK, JJJ).Value ' v' 対象pointCy入力
  
'''''''''''測定結果表示プログラム'''''''''''''''''

Dim MX As Double

Sheets("xy入力").Cells(55, 8).Value = TextBox1 'File名(TextBox1)をExcelに入力
MX = Format(Now, "yyyymmddhhnnss") '日時

Sheets("xy入力").Cells(55, 7).Value = MX '測定日時をExcelに入力

'⊿u'v'(u'方向)データの整理
Sheets("xy入力").Range(Sheets("xy入力").Cells(55, 9), Sheets("xy入力").Cells(55, 15)).Value = Sheets("⊿u'v'計算").Range("h9:n9").Value

'⊿u'v'(v'方向)データの整理
Sheets("xy入力").Range(Sheets("xy入力").Cells(55, 16), Sheets("xy入力").Cells(55, 22)).Value = Sheets("⊿u'v'計算").Range("h10:n10").Value

'MaxJNDデータの整理
Sheets("xy入力").Range(Sheets("xy入力").Cells(55, 23), Sheets("xy入力").Cells(55, 30)).Value = Sheets("JND計算").Range("h9:o9").Value


'''MaxJND UserFormに表記'''

'中心値

Label37 = Round(Sheets("xy入力").Range("an6").Value, 0)
Label38 = Round(Sheets("xy入力").Range("ao6").Value, 4)
Label39 = Round(Sheets("xy入力").Range("ap6").Value, 4)



'MaxJND
Label3 = Round(Sheets("JND計算").Range("h9").Value, 2)

'⊿u'v'
Label7 = Round(Sheets("JND計算").Range("j9").Value, 4) + Round(Sheets("JND計算").Range("k9").Value, 4)

'position
Label9 = Sheets("JND計算").Range("l9").Value & "," & Sheets("JND計算").Range("m9").Value & " -> " & Sheets("JND計算").Range("n9").Value & "," & Sheets("JND計算").Range("o9").Value

'9point輝度ムラ
Label45 = Round(Sheets("xy入力").Range("ao13").Value, 3)

'9point⊿u'v'
Label47 = Round(Sheets("xy入力").Range("ao14").Value, 4)


'''Max⊿u'v'(u'方向) UserFormに表記'''
'JND
Label11 = Round(Sheets("⊿u'v'計算").Range("j9").Value, 2)

'Max⊿u'v'
Label15 = Round(Sheets("⊿u'v'計算").Range("h9").Value, 4)

'position
Label16 = Sheets("⊿u'v'計算").Range("k9").Value & "," & Sheets("⊿u'v'計算").Range("l9").Value & " -> " & Sheets("⊿u'v'計算").Range("m9").Value & "," & Sheets("⊿u'v'計算").Range("n9").Value


'''Max⊿u'v'(v'方向) UserFormに表記'''
'JND
Label19 = Round(Sheets("⊿u'v'計算").Range("j10").Value, 2)

'Max⊿u'v'
Label23 = Round(Sheets("⊿u'v'計算").Range("h10").Value, 4)

'position
Label24 = Sheets("⊿u'v'計算").Range("k10").Value & "," & Sheets("⊿u'v'計算").Range("l10").Value & " -> " & Sheets("⊿u'v'計算").Range("m10").Value & "," & Sheets("⊿u'v'計算").Range("n10").Value


'''''''''''測定画像の保存''''''''''''''

'   Dim strName As String
'    Dim TabIndex As Long
'    Dim ObjectIndex As Long

'    TabIndex = 1 'CA2000 擬似カラー
'  ObjectIndex = 1 'CA2000 輝度画像選択
'    strName = TextBox2 & MX & "_" & TextBox1 & "Ey.bmp" '画像のfile名
'    Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '画像保存
    
 '   TabIndex = 1 'CA2000 擬似カラー
 '   ObjectIndex = 2 'CA2000 Cx画像選択
 '  strName = TextBox2 & MX & "_" & TextBox1 & "Cx.bmp" '画像のfile名
 '   Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '画像保存

 '   TabIndex = 1 'CA2000 擬似カラー
 '   ObjectIndex = 3 'CA2000 Cy画像選択
 '   strName = TextBox2 & MX & "_" & TextBox1 & "Cy.bmp" '画像のfile名
 '   Ret = myappl.SaveObjectAsBMPFile(strName, TabIndex, ObjectIndex) '画像保存

 '   Set Image1.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Ey.bmp") 'ユーザーファームに輝度分布表示
 '   Set Image2.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Cx.bmp") 'ユーザーファームにCx分布表示
 '   Set Image3.Picture = LoadPicture(TextBox2 & MX & "_" & TextBox1 & "Cy.bmp") 'ユーザーファームにCy分布表示
    
    
'''''''''測定データの保存'''''''''''''

'【CA2000】.xlsファイルを起動する
    Workbooks.Open Filename:= _
        "D:\CA2000data\【CA2000_08model】.xls"

'【CA2000】.xlsファイルに整理した測定データをコピー
ThisWorkbook.Sheets("xy入力").Range("AN2:AP10").Copy ActiveWorkbook.Worksheets(1).Range("Q10")
ThisWorkbook.Sheets("xy入力").Range("G55:AD56").Copy ActiveWorkbook.Worksheets(1).Range("A4")
ThisWorkbook.Sheets("xy入力").Range("E1:AL50").Copy ActiveWorkbook.Worksheets(1).Range("AC1")
ThisWorkbook.Sheets("xy入力").Range("AN53:AP68").Copy ActiveWorkbook.Worksheets(1).Range("Q21") '追加071217
ThisWorkbook.Sheets("xy入力").Range("AN69:AN72").Copy ActiveWorkbook.Worksheets(1).Range("Q38") '追加071217
ThisWorkbook.Sheets("xy入力").Range("AN73:AN76").Copy ActiveWorkbook.Worksheets(1).Range("Q43") '追加071217


'ActiveWorkbook.Sheets(1).Range("A6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Ey.bmp").Select
'ThisWorkbook.Sheets("xy入力").Range("G6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Cx.bmp").Select
'ThisWorkbook.Sheets("xy入力").Range("M6").Pictures.Insert(TextBox2 & MX & "_" & TextBox1 & "Cy.bmp").Select

'【CA2000】.xlsファイルを名前を変更して保存
ActiveWorkbook.SaveAs TextBox2 & MX & "_" & TextBox1 & ".xls"

'保存したExcelを閉じる
Workbooks(MX & "_" & TextBox1 & ".xls").Close Savechanges:=False

    Set SpotCond = Nothing
    Set myappl = Nothing
    
    Exit Sub
      
'''''''エラー処理'''''''''
    
エラー処理へ:

 MsgBox "・CA-S20wの測定画面が起動していませんか？" & vbCrLf & "・496Spotを指定していますか？"
 
           
End Sub


'''''''''''CA2000自動校正用プログラム'''''''''''''''''


Private Sub CommandButton2_Click()

Dim CSEy As Double 'CS1000校正輝度
Dim CSCx As Double 'CS1000校正Cx
Dim CSCy As Double 'CS1000校正Cy

CSEy = TextBox3 'CS1000校正輝度取得
CSCx = TextBox4 'CS1000校正Cx取得
CSCy = TextBox5 'CS1000校正Cy取得

On Error GoTo エラー処理へ

 Set myappl = CreateObject("CAS20W.Application") 'CA2000ソフトとのリンク


''''''''校正前測定''''''''''

    Dim Ret As Long
    Ret = myappl.Measure() 'CA2000 測定プログラム
    
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
    
    

''''''''校正前輝度・色度値取得''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Dim SpotCond As Object  'CA2000 スポット測定データ取得プログラム
    
    Dim CalEy As Double '校正前輝度
    Dim CalCx As Double '校正前Cx
    Dim CalCy As Double '校正前Cy
    
    Ret = myappl.SelectData()

    Set SpotCond = myappl.GetSpotCondition() 'CA2000 Spot設定条件を取得
    
    Dim SpotCount As Long
    SpotCount = SpotCond.GetSpotCount() 'CA2000 Spot個数を取得
    
    Dim SpotData(100000) As Single 'CA2000 Spot個数分の配列確保
    
    
    Element = 3  'Spot 輝度データの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
    
      CalEy = SpotData(500) '校正前輝度取得


    Element = 4  'Spot Cxデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCx = SpotData(500) '校正前Cx取得


    Element = 5  'Spot Cyデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCy = SpotData(500) '校正前Cy取得

   
    
'''''''''校正値入力''''''''''

Dim Condition As Object

Dim UserCalibData As Object

 Set Condition = myappl.GetMeasurementCondition() 'CA2000 測定条件の取得

 Set UserCalibData = Condition.GetUserCalibrationData() 'ユーザー校正係数オブジェクトの取得

UserCalibData.CalibrationType = 2 'CA2000 １点校正選択


UserCalibData.WLv_before = CalEy '校正前輝度
UserCalibData.WLv_after = CSEy 'CS1000校正輝度

UserCalibData.Wx_before = CalCx '校正前Cx
UserCalibData.Wx_after = CSCx 'CS1000校正Cx

UserCalibData.Wy_before = CalCy '校正前Cy
UserCalibData.Wy_after = CSCy 'CS1000校正Cy

Ret = Condition.SetUserCalibrationData(UserCalibData)



''''''''校正測定''''''''

   Ret = myappl.Measure() 'CA2000 測定プログラム
    
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
    
    

'''''''''校正後輝度・色度確認'''''''''''
 
    Ret = myappl.SelectData()

  
    Element = 3 'Spot 輝度データの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
      CalEy = SpotData(500)
      Label37 = Round(CalEy, 0) '校正後輝度取得
                  
    
    Element = 4 'Spot Cxデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCx = SpotData(500)
    Label38 = Round(CalCx, 4) '校正後Cx取得
    
    
    Element = 5 'Spot Cyデータの取得
    Ret = myappl.GetSpotData(Element, SpotData)
    
    CalCy = SpotData(500)
    Label39 = Round(CalCy, 4) '校正後Cy取得


    Set SpotCond = Nothing
    Set myappl = Nothing
    
    MsgBox "校正完了"  '測定案内メッセージ
    
    Exit Sub
    
    
        
'''''''エラー処理'''''''''
    
エラー処理へ:

 MsgBox "・CA-S20wの測定画面が起動していませんか？"
    
    
End Sub

