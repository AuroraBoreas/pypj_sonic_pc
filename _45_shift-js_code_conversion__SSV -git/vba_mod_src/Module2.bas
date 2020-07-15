Attribute VB_Name = "Module2"
Sub uv21()

Dim AAA, BBB, CCC, DDD, AA, BB As Long
Dim CC, DD, EE, EEE As Double

   
  ''''''''''''''''''''''''''21Max⊿u'v''''''''''''''''''''''''''''''''''''''''''''''''
   
   Sheets("⊿u'v'計算").Range("h17:i512").Value = 0 '前回対象色度CxCy消去
   Sheets("⊿u'v'計算").Range("o14:r454").Value = 0 '前回結果消去
   AA = Sheets("⊿u'v'計算").Range("B3").Value 'point数
   BB = AA + 1
     
'''''対象色度u'v'入力'''''

For Counter0 = 2 To BB

   Sheets("⊿u'v'計算").Cells(Counter0 + 15, 8).Value = Sheets("u'v'変換").Cells(Counter0, 2).Value 'u 対象色度u'入力
   Sheets("⊿u'v'計算").Cells(Counter0 + 15, 9).Value = Sheets("u'v'変換").Cells(Counter0, 3).Value 'v 対象v'色度入力
   Sheets("⊿u'v'計算").Cells(Counter0 + 15, 12).Value = Counter0 - 1 '対象point入力

Next Counter0


'''''基準色度u'v'入力と全データ表記'''''

  For Counter1 = 2 To BB
  
    CC = 14 + AA * (Counter1 - 2)
    DD = CC + AA - 1
    

   Sheets("⊿u'v'計算").Range("h16").Value = Sheets("u'v'変換").Cells(Counter1, 2).Value 'u 基準色度u'
   Sheets("⊿u'v'計算").Range("i16").Value = Sheets("u'v'変換").Cells(Counter1, 3).Value 'v 基準色度v'
   Sheets("⊿u'v'計算").Range("L16").Value = Counter1 - 1 '基準point入力

 
   Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(CC, 15), Sheets("⊿u'v'計算").Cells(DD, 15)).Value = Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(17, 10), Sheets("⊿u'v'計算").Cells(AA + 17, 10)).Value ' u'>v' ⊿u'v'表記
   Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(CC, 16), Sheets("⊿u'v'計算").Cells(DD, 16)).Value = Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(17, 11), Sheets("⊿u'v'計算").Cells(AA + 17, 11)).Value ' u'<=v' ⊿u'v'表記
   Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(CC, 17), Sheets("⊿u'v'計算").Cells(DD, 17)).Value = Sheets("⊿u'v'計算").Range("L16").Value '基準point表記
   Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(CC, 18), Sheets("⊿u'v'計算").Cells(DD, 18)).Value = Sheets("⊿u'v'計算").Range(Sheets("⊿u'v'計算").Cells(17, 12), Sheets("⊿u'v'計算").Cells(AA + 17, 12)).Value '対象point表記

  Next Counter1
 
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
End Sub
Sub uv496()

Dim AAA, BBB, CCC, DDD, EEE, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double
Dim FFF, GGG, HHH, III, JJJ, KKK As Double

    
'''''''''''''''''''''''''''''''''496Max⊿u'v'''''''''''''''''''''''''''''''''''''''''


      Sheets("⊿u'v'計算").Range("H9:H10").Value = 0 '前回Max⊿u'v'消去
      Sheets("⊿u'v'計算").Range("h17:i512").Value = 0 '前回色度u'v'消去
  
  
       AAA = Sheets("⊿u'v'計算").Range("B8").Value '行数
       BBB = Sheets("⊿u'v'計算").Range("B9").Value '列数
       
       
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
        
        Sheets("u'v'変換").Cells(FF, GG).Interior.ColorIndex = 3 ' u' 基準point目印
        Sheets("u'v'変換").Cells(HH, GG).Interior.ColorIndex = 3 ' v' 基準point目印
        
        Sheets("u'v'変換").Cells(II, JJ).Interior.ColorIndex = 7 ' u' 対象point目印
        Sheets("u'v'変換").Cells(KK, JJ).Interior.ColorIndex = 7 ' v' 対象point目印
        
        
        
 ''' u'<=v' '''
        FFF = Sheets("⊿u'v'計算").Range("k10").Value + 1 'Max⊿u'v' 基準(行)
        GGG = Sheets("⊿u'v'計算").Range("l10").Value + 6 'Max⊿u'v' 基準(列)
        HHH = FFF + 16
        III = Sheets("⊿u'v'計算").Range("m10").Value + 1 'Max⊿u'v' 対象(行)
        JJJ = Sheets("⊿u'v'計算").Range("n10").Value + 6 'Max⊿u'v' 対象(列)
        KKK = III + 16
        
        Sheets("u'v'変換").Cells(FFF, GGG).Interior.ColorIndex = 5 ' u' 基準point目印
        Sheets("u'v'変換").Cells(HHH, GGG).Interior.ColorIndex = 5 ' v' 基準point目印
        
        Sheets("u'v'変換").Cells(III, JJJ).Interior.ColorIndex = 8 ' u' 対象point目印
        Sheets("u'v'変換").Cells(KKK, JJJ).Interior.ColorIndex = 8 ' v' 対象point目印
        
        
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
        
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
