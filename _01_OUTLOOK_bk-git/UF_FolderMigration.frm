VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_FolderMigration 
   Caption         =   "UserForm1"
   ClientHeight    =   4130
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   7980
   OleObjectBlob   =   "UF_FolderMigration.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_FolderMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objNewPSTFolder As Outlook.Folder

Private Sub CommandButton1_Click()
''''''''''''''''''''''''''CopyFolderStructure''''''''''''''''''''''''''
    Dim OldName As String: OldName = Me.txb_OldName
    Dim NewName As String: NewName = Me.txb_NewName
    Dim NewAddress As String: NewAddress = Me.txb_NewAddrs
    
    Dim objFolders As Outlook.Folders
    Dim objFolder As Outlook.Folder

    'Get the folders of the source Outlook PST file
    Set objFolders = Outlook.Application.Session.Folders(OldName).Folders
    
    'Create the new pst file in your desired local folder and name
    Outlook.Application.Session.AddStore NewAddress
    Set objNewPSTFolder = Session.Folders.GetLast()
    
    For Each objFolder In objFolders
        CreateFolder objFolder
    Next
    'Revise its name
    Outlook.Session.Folders.GetLast.Name = NewName
    MsgBox "Completed!", vbOKOnly + vbInformation, "Copy Folder Structure"
    Unload Me
End Sub

Private Sub CreateFolder(objFolder As Outlook.Folder)
    Dim NewName As String: NewName = Me.txb_NewName
    Dim objSubFolder As Outlook.Folder
    'Only copy the mail folder
    On Error Resume Next
    If (objFolder.DefaultItemType = olMailItem) Then
       'New Outlook PST file auto includes the "Deleted Items" folder, so skip it
       'Skip the useless mail folders - "Conversation Action Settings" and "Quick Step Settings"
       If (objFolder.Name <> Me.chb1.Caption) And _
          (objFolder.Name <> Me.chb2.Caption) And _
          (objFolder.Name <> Me.chb3.Caption) And _
          (objFolder.Name <> Me.chb4.Caption) And _
          (objFolder.Name <> Me.chb5.Caption) And _
          (objFolder.Name <> Me.chb6.Caption) Then
          'Create the new folder
          objNewPSTFolder.Folders.Add objFolder.Name
          Set objNewPSTFolder = objNewPSTFolder.Folders.Item(objFolder.Name)
          
          For Each objSubFolder In objFolder.Folders
              CreateFolder objSubFolder
          Next
          
          Set objNewPSTFolder = objNewPSTFolder.Parent
       End If
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub txb_NewName_Change()
    With Me
        .txb_NewAddrs = "D:\Reference_02_archived\email_archived" & "\" & .txb_NewName & ".pst"
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Folder Structure Migration Tool " & " v1, by Z.Liang, 2017"
        .txb_OldName = "2017"
        .txb_NewName = "2018"
        .txb_NewAddrs = "D:\Reference_02_archived\email_archived" & "\" & .txb_NewName & ".pst"
    End With
End Sub
