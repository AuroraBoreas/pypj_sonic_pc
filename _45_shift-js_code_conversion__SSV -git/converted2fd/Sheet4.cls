VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub OptionButton1_Click()

        
        Range("o5:AA42").Select
        Application.CutCopyMode = False
        Selection.Clear
       
        Range("Q5").Select
        ActiveCell.FormulaR1C1 = "100 IRE"
        Range("U5").Select
        ActiveCell.FormulaR1C1 = "50 IRE"
        
        
        ActiveSheet.CheckBox1.Value = False
        ActiveSheet.CheckBox2.Value = False
        ActiveSheet.CheckBox3.Value = False
        ActiveSheet.CheckBox4.Value = False
        ActiveSheet.CheckBox5.Value = False
        ActiveSheet.CheckBox6.Value = False
        
        Range("A1").Select
        
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(220, 230, 242)
 
 
 ActiveSheet.ChartObjects("グラフ 2").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(220, 230, 242)
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
        
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
 ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        
        

 End Sub
 
 
Private Sub OptionButton2_Click()

 Range("o5:AA42").Select
        Application.CutCopyMode = False
        Selection.Clear
       
        Range("R5").Select
        ActiveCell.FormulaR1C1 = "100 IRE"
        Range("X5").Select
        ActiveCell.FormulaR1C1 = "50 IRE"
        
        
        ActiveSheet.CheckBox1.Value = False
        ActiveSheet.CheckBox2.Value = False
        ActiveSheet.CheckBox3.Value = False
        ActiveSheet.CheckBox4.Value = False
        ActiveSheet.CheckBox5.Value = False
        ActiveSheet.CheckBox6.Value = False
        
        Range("A1").Select
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(240, 230, 242)
 
 
 ActiveSheet.ChartObjects("グラフ 2").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(240, 230, 242)
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
        
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
 
 
 
 ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With

        
        
 End Sub
 
Private Sub OptionButton3_Click()

  
        Range("o5:AA42").Select
        Application.CutCopyMode = False
        Selection.Clear
       
        Range("Q5").Select
        ActiveCell.FormulaR1C1 = "100 IRE"
        
        ActiveSheet.CheckBox1.Value = False
        ActiveSheet.CheckBox2.Value = False
        ActiveSheet.CheckBox3.Value = False
        ActiveSheet.CheckBox4.Value = False
        ActiveSheet.CheckBox5.Value = False
        ActiveSheet.CheckBox6.Value = False

        Range("A1").Select
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(240, 240, 240)
 
 
 ActiveSheet.ChartObjects("グラフ 2").Activate
 ActiveChart.ChartArea.Interior.Color = rgb(240, 240, 240)
 
 ActiveSheet.ChartObjects("グラフ 1").Activate
        
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
         ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
 
 ActiveSheet.ChartObjects("グラフ 2").Activate
        
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
 End Sub
 
Private Sub checkbox1_click()

 Select Case CheckBox1.Value
 
   Case True:
   
    If OptionButton1.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(2).XValues = "='data (2)'!$AI$3:$AI$10"
         ActiveChart.SeriesCollection(2).Values = "='data (2)'!$AJ$3:$AJ$10"
         ActiveChart.SeriesCollection(2).Select
         With Selection
        .MarkerBackgroundColorIndex = 5 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 5 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(2).XValues = "='data (2)'!$ar$3:$ar$10"
         ActiveChart.SeriesCollection(2).Values = "='data (2)'!$as$3:$as$10"
         ActiveChart.SeriesCollection(2).Select
         With Selection
        .MarkerBackgroundColorIndex = 5 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 5 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
    
        Range("al8:As10").Select
        Selection.Copy
        Range("o8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
    
        
    If OptionButton2.Value = True Then

        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(2).XValues = "='data (2)'!$g$3:$g$26"
        ActiveChart.SeriesCollection(2).Values = "='data (2)'!$h$3:$h$26"
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = 5 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 5 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(2).XValues = "='data (2)'!$p$3:$p$26"
        ActiveChart.SeriesCollection(2).Values = "='data (2)'!$q$3:$q$26"
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = 5 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 5 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
    
        Range("au7:bf11").Select
        Selection.Copy
        Range("o7").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
         SkipBlanks:=False, Transpose:=False
        Range("A3").Select
    
    End If
    
    If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(2).XValues = "='data (2)'!$y$3:$y$10"
         ActiveChart.SeriesCollection(2).Values = "='data (2)'!$z$3:$z$10"
         ActiveChart.SeriesCollection(2).Select
         With Selection
        .MarkerBackgroundColorIndex = 5 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 5 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh8:bk10").Select
        Selection.Copy
        Range("o8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
        
  
        
   Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(2).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o7:AA11").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select

     End Select

   
End Sub

Private Sub checkbox2_click()
  

Select Case CheckBox2.Value

    Case True:
    
     If OptionButton1.Value = True Then
    
    
         
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(3).XValues = "='data (2)'!$AI$30:$AI$37"
         ActiveChart.SeriesCollection(3).Values = "='data (2)'!$AJ$30:$AJ$37"
         ActiveChart.SeriesCollection(3).Select
         With Selection
        .MarkerBackgroundColorIndex = 6 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 6 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(3).XValues = "='data (2)'!$ar$30:$ar$37"
         ActiveChart.SeriesCollection(3).Values = "='data (2)'!$As$30:$As$37"
         ActiveChart.SeriesCollection(3).Select
         With Selection
        .MarkerBackgroundColorIndex = 6 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 6 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         
        Range("al14:As16").Select
        Selection.Copy
        Range("o14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
    
    End If
    
    If OptionButton2.Value = True Then
     
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(3).XValues = "='data (2)'!$g$30:$g$53"
        ActiveChart.SeriesCollection(3).Values = "='data (2)'!$h$30:$h$53"
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = 6 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 6 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(3).XValues = "='data (2)'!$p$30:$p$53"
        ActiveChart.SeriesCollection(3).Values = "='data (2)'!$q$30:$q$53"
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = 6 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 6 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
             
        Range("au13:bf17").Select
        Selection.Copy
        Range("o13").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
     End If
    
    If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(3).XValues = "='data (2)'!$y$30:$y$37"
         ActiveChart.SeriesCollection(3).Values = "='data (2)'!$z$30:$z$37"
         ActiveChart.SeriesCollection(3).Select
         With Selection
        .MarkerBackgroundColorIndex = 6 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 6 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh14:bk16").Select
        Selection.Copy
        Range("o14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
     
        
    Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(3).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o13:AA17").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select

      End Select
      
End Sub

Private Sub checkbox3_click()

Select Case CheckBox3.Value

    Case True:
      
      If OptionButton1.Value = True Then
      
    
        
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(4).XValues = "='data (2)'!$AI$57:$AI$64"
         ActiveChart.SeriesCollection(4).Values = "='data (2)'!$AJ$57:$AJ$64"
         ActiveChart.SeriesCollection(4).Select
         With Selection
        .MarkerBackgroundColorIndex = 7 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 7 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(4).XValues = "='data (2)'!$ar$57:$ar$64"
         ActiveChart.SeriesCollection(4).Values = "='data (2)'!$As$57:$As$64"
         ActiveChart.SeriesCollection(4).Select
         With Selection
        .MarkerBackgroundColorIndex = 7 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 7 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
         
        Range("al20:As22").Select
        Selection.Copy
        Range("o20").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
        
    If OptionButton2.Value = True Then
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(4).XValues = "='data (2)'!$g$57:$g$80"
        ActiveChart.SeriesCollection(4).Values = "='data (2)'!$h$57:$h$80"
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = 7 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 7 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(4).XValues = "='data (2)'!$p$57:$p$80"
        ActiveChart.SeriesCollection(4).Values = "='data (2)'!$q$57:$q$80"
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = 7 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 7 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        Range("Au19:bf23").Select
        Selection.Copy
        Range("o19").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
    
    If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(4).XValues = "='data (2)'!$y$57:$y$64"
         ActiveChart.SeriesCollection(4).Values = "='data (2)'!$z$57:$z$64"
         ActiveChart.SeriesCollection(4).Select
         With Selection
        .MarkerBackgroundColorIndex = 7 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 7 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh20:bk22").Select
        Selection.Copy
        Range("o20").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
     
    
    Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(4).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o19:AA23").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select

    End Select

End Sub

Private Sub checkbox4_click()

Select Case CheckBox4.Value

    Case True:
    
    If OptionButton1.Value = True Then
      
       
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(5).XValues = "='data (2)'!$AI$84:$AI$91"
         ActiveChart.SeriesCollection(5).Values = "='data (2)'!$AJ$84:$AJ$91"
         ActiveChart.SeriesCollection(5).Select
         With Selection
        .MarkerBackgroundColorIndex = 8 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 8 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(5).XValues = "='data (2)'!$ar$84:$ar$91"
         ActiveChart.SeriesCollection(5).Values = "='data (2)'!$As$84:$As$91"
         ActiveChart.SeriesCollection(5).Select
         With Selection
        .MarkerBackgroundColorIndex = 8 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 8 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
             
        Range("Al26:As28").Select
        Selection.Copy
        Range("o26").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
    
    If OptionButton2.Value = True Then
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(5).XValues = "='data (2)'!$g$84:$g$107"
        ActiveChart.SeriesCollection(5).Values = "='data (2)'!$h$84:$h$107"
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = 8 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 8 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(5).XValues = "='data (2)'!$p$84:$p$107"
        ActiveChart.SeriesCollection(5).Values = "='data (2)'!$q$84:$q$107"
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = 8 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 8 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
                 
        Range("Au25:bf29").Select
        Selection.Copy
        Range("o25").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
        End If
        
     If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(5).XValues = "='data (2)'!$y$84:$y$91"
         ActiveChart.SeriesCollection(5).Values = "='data (2)'!$z$84:$z$91"
         ActiveChart.SeriesCollection(5).Select
         With Selection
        .MarkerBackgroundColorIndex = 8 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 8 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh26:bk28").Select
        Selection.Copy
        Range("o26").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
        
        
    Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(5).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o25:AA29").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select

    End Select
     
     
End Sub

Private Sub checkbox5_click()

  Select Case CheckBox5.Value

    Case True:
    
    If OptionButton1.Value = True Then
    
     
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(6).XValues = "='data (2)'!$AI$111:$AI$118"
         ActiveChart.SeriesCollection(6).Values = "='data (2)'!$AJ$111:$AJ$118"
         ActiveChart.SeriesCollection(6).Select
         With Selection
        .MarkerBackgroundColorIndex = 9 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 9 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(6).XValues = "='data (2)'!$ar$111:$ar$118"
         ActiveChart.SeriesCollection(6).Values = "='data (2)'!$As$111:$As$118"
         ActiveChart.SeriesCollection(6).Select
         With Selection
        .MarkerBackgroundColorIndex = 9 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 9 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
                      
        Range("Al32:As34").Select
        Selection.Copy
        Range("o32").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
     End If
      
      
        If OptionButton2.Value = True Then
    
         ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(6).XValues = "='data (2)'!$g$110:$g$134"
        ActiveChart.SeriesCollection(6).Values = "='data (2)'!$h$110:$h$134"
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = 9 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 9 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(6).XValues = "='data (2)'!$p$110:$p$134"
        ActiveChart.SeriesCollection(6).Values = "='data (2)'!$q$110:$q$134"
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = 9 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 9 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
                  
        Range("Au31:bf35").Select
        Selection.Copy
        Range("o31").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
        End If
        
    If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(6).XValues = "='data (2)'!$y$111:$y$118"
         ActiveChart.SeriesCollection(6).Values = "='data (2)'!$z$111:$z$118"
         ActiveChart.SeriesCollection(6).Select
         With Selection
        .MarkerBackgroundColorIndex = 9 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 9 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh32:bk34").Select
        Selection.Copy
        Range("o32").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
          
        
        
    Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(6).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o31:AA35").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select


  End Select

End Sub

Private Sub checkbox6_click()

  Select Case CheckBox6.Value

    Case True:
    
    If OptionButton1.Value = True Then
    
     
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(7).XValues = "='data (2)'!$AI$138:$AI$145"
         ActiveChart.SeriesCollection(7).Values = "='data (2)'!$AJ$138:$AJ$145"
         ActiveChart.SeriesCollection(7).Select
         With Selection
        .MarkerBackgroundColorIndex = 10 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 10 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
         ActiveSheet.ChartObjects("グラフ 2").Activate
         ActiveChart.SeriesCollection(7).XValues = "='data (2)'!$ar$138:$ar$145"
         ActiveChart.SeriesCollection(7).Values = "='data (2)'!$As$138:$As$145"
         ActiveChart.SeriesCollection(7).Select
         With Selection
        .MarkerBackgroundColorIndex = 10 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 10 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
        
                      
        Range("Al38:AS40").Select
        Selection.Copy
        Range("o38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
     End If
      
      
     If OptionButton2.Value = True Then
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(7).XValues = "='data (2)'!$g$138:$g$161"
        ActiveChart.SeriesCollection(7).Values = "='data (2)'!$h$138:$h$161"
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = 10 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 10 'マーカの外枠の色
        .MarkerStyle = xlSquare 'マーカの種類
        End With

        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.PlotArea.Select
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(7).XValues = "='data (2)'!$p$138:$p$161"
        ActiveChart.SeriesCollection(7).Values = "='data (2)'!$q$138:$q$161"
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = 10 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 10 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
                  
        Range("Au37:bf41").Select
        Selection.Copy
        Range("o37").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
        End If
        
    If OptionButton3.Value = True Then

    
         ActiveSheet.ChartObjects("グラフ 1").Activate
         ActiveChart.SeriesCollection(7).XValues = "='data (2)'!$y$138:$y$145"
         ActiveChart.SeriesCollection(7).Values = "='data (2)'!$z$138:$z$145"
         ActiveChart.SeriesCollection(7).Select
         With Selection
        .MarkerBackgroundColorIndex = 10 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = 10 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
         End With
    
        Range("bh38:bk40").Select
        Selection.Copy
        Range("o32").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Range("A3").Select
        
    End If
          
        
        
    Case False:
    
        ActiveSheet.ChartObjects("グラフ 1").Activate
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare  'マーカの種類
        End With
        
        ActiveSheet.ChartObjects("グラフ 2").Activate
        ActiveChart.SeriesCollection(7).Select
        With Selection
        .MarkerBackgroundColorIndex = xlNone 'マーカの中の色　中抜きは[xlNone]
        .MarkerForegroundColorIndex = xlNone 'マーカの外枠の色
        .MarkerStyle = xlSquare    'マーカの種類
        End With
        
        Range("o37:AA41").Select
        Application.CutCopyMode = False
        Selection.Clear
        Range("A3").Select


  End Select

End Sub
