VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectColumn 
   Caption         =   "Select Column"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "frmSelectColumn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SelectColumn"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents Model As clsSelectColumnViewModel
Attribute Model.VB_VarHelpID = -1

Private Type TFrmSelectColumnView
    IsCancelled As Boolean
End Type

Private this As TFrmSelectColumnView

Private Sub chkReplaceDestColName_Click()
    Model.ReplaceColumnHeader = Me.chkReplaceDestColName.Value
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbClearSearch_Click()
    Me.txtSearch.text = vbNullString
    Me.txtSearch.SetFocus
End Sub

Private Sub cmbOK_Click()
    Hide
End Sub

Private Sub lbColumns_Change()
    If Me.lbColumns.ListCount = 0 Then Exit Sub
    Model.TrySelect Me.lbColumns
End Sub

Private Sub lbColumns_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Model.SelectedColumn Is Nothing Then
        Hide
    End If
End Sub

Private Sub lbColumns_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii = 13) And (Not Model.SelectedColumn Is Nothing) Then
        Hide
    End If
End Sub

Private Sub Model_CollectionChanged()
    Me.lbColumns.Clear
    Dim v As Variant
    For Each v In Model.Columns
        Me.lbColumns.AddItem v
    Next v
End Sub

Private Sub Model_ItemSelected()
    Me.cmbOK.Enabled = True
End Sub

Private Sub txtSearch_Change()
    Model.SearchCriteria = Me.txtSearch & "*"
    If Model.Columns.Count > 0 Then
        Me.lbColumns.selected(0) = True
    End If
End Sub

Private Sub UserForm_Activate()
    Model.SearchCriteria = vbNullString
    Me.cmbOK.Enabled = False
    Me.cmbClearSearch.Picture = Application.CommandBars.GetImageMso("Delete", 16, 16)
    Model.ReplaceColumnHeader = Me.chkReplaceDestColName.Value
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
    Set Model = viewModel
    Show
    ShowDialog = Not this.IsCancelled
End Function
