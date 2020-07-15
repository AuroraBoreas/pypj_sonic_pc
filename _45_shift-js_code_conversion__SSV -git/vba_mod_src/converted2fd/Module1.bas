Attribute VB_Name = "Module1"
Sub JND496()

Dim AAA, BBB, CCC, DDD, AA, BB, CC As Long
Dim DD, EE, FF, GG, HH, II, JJ, KK As Double

    
'''''''''''''''''''''''''''''''''496MaxJND'''''''''''''''''''''''''''''''''''''''''
      
'''''対象色度CxCy入力'''''
      
       Sheets("JND計算").Range("H9").Value = 0 '前回MaxJND消去
       Sheets("JND計算").Range("I17:J512").Value = 0 '前回対象色度CxCy消去
       
       AAA = Sheets("JND計算").Range("B8").Value '行数
       BBB = Sheets("JND計算").Range("B9").Value '列数

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

      EE = 0
      
      For CounterB = 7 To BBB + 6
       For CounterC = 2 To AAA + 1
   
       
        DD = Sheets("JND計算").Range("H9").Value '現状MaxJND
  
                
        If DD < EE Then '現状MaxJND VS 追加MaxJND 大小比較

        Sheets("JND計算").Range("H9").Value = Sheets("JND計算").Range("K14").Value '496MaxJND入力
        Sheets("JND計算").Range("j9").Value = Sheets("JND計算").Range("l16").Value 'MaxJND 基準(行)入力
        Sheets("JND計算").Range("k9").Value = Sheets("JND計算").Range("m16").Value 'MaxJND 基準(列)入力
        Sheets("JND計算").Range("l9").Value = Sheets("JND計算").Range("l14").Value 'MaxJND 対象(行)入力
        Sheets("JND計算").Range("m9").Value = Sheets("JND計算").Range("m14").Value 'MaxJND 対象(列)入力
        
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
  
        FF = Sheets("JND計算").Range("j9").Value + 1 'MaxJND 基準(行)
        GG = Sheets("JND計算").Range("k9").Value + 6 'MaxJND 基準(列)
        HH = FF + 16
        II = Sheets("JND計算").Range("l9").Value + 1 'MaxJND 対象(行)
        JJ = Sheets("JND計算").Range("m9").Value + 6 'MaxJND 対象(列)
        KK = II + 16
        
        
        Sheets("xy入力").Cells(FF, GG).Interior.ColorIndex = 3 'x 基準point目印
        Sheets("xy入力").Cells(HH, GG).Interior.ColorIndex = 3 'y 基準point目印
        
        Sheets("xy入力").Cells(II, JJ).Interior.ColorIndex = 7 'x 対象point目印
        Sheets("xy入力").Cells(KK, JJ).Interior.ColorIndex = 7 'y 対象point目印
        
       
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

