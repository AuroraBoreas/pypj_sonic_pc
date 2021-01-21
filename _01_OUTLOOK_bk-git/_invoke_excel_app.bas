Attribute VB_Name = "mod1"
Option Explicit

Sub open_excel()
    Dim xlApp As Excel.Application
    Dim myWB As Excel.Workbook
    
    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.Application.Visible = True
    Set myWB = xlApp.Workbooks.Add
    
    xlApp.Left = 430.75
    xlApp.Top = 1
    xlApp.Width = 528
    xlApp.Height = 540

End Sub

