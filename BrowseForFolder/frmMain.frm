VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BrowseForFolder  (SHELL)  Demo"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' 2003 By Paul Veenstra
' You may use this code and redistribute it.
' There's no copyright.
'
Private Sub cmdBrowse_Click()

' Make reference to:
'    Microsoft Shell Controls and Automation
'
' According to the MSDN the return object for
' BrowseForFOlder is Folder, but this is not the case.
' On newer systems the return object is Folder2.
' Therefore i've made objFolder an object.

Dim objShell As New Shell           ' Reference to the Shell32 Lib.
Dim objFolder As Object             ' Container for the Folder Class.
Dim objItem As FolderItem           ' Container for the FolderItem Class.
Const BIF_NEWDIALOGSTYLE = &H40     ' Constand that shows create folder in the
                                    ' browse for folder dialog.

On Error GoTo errHand               ' Set the error handling in case we cancle
                                    ' the dialog.
                                    

' Show the dialog and pass the Folder object.
Set objFolder = objShell.BrowseForFolder(Me.hWnd, _
                                         "Select a Folder", _
                                         BIF_NEWDIALOGSTYLE)

Set objItem = objFolder.Self        ' Set the current selected Folder reference.
txtPath = objItem.Path              ' Show the path in de textbox.

On Error GoTo 0                     ' Turn off error handling.

Exit Sub
errHand:                            ' Just trapping, no handling.
End Sub
