VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateTable 
   Caption         =   "Create Table"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   OleObjectBlob   =   "frmCreateTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ReferenceCreateTable"
Option Explicit

Private Sub cmbOK_Click()
    MsgBox Me.refRange.text
End Sub
