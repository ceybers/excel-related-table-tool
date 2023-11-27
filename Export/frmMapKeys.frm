VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMapKeys 
   Caption         =   "Map Related Table"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "frmMapKeys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMapKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MapKeys"
Option Explicit

Private Type TFrmMapKeysView
    Model As clsMapKeysViewModel
    IsCancelled As Boolean
End Type

Private this As TFrmMapKeysView

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    this.Model.Source = Me.refSource.text
    this.Model.Destination = Me.refDestination.text
    Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Hide
End Sub

Public Function ShowDialog(ByVal viewModel As Object) As Boolean
    Set this.Model = viewModel
    
    If this.Model.Source <> vbNullString Then
        Me.refSource.text = this.Model.Source
    End If
    
    If this.Model.Source <> vbNullString Then
        Me.refDestination.text = this.Model.Destination
    End If
    
    Show
    ShowDialog = Not this.IsCancelled
End Function

