VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectColumns 
   Caption         =   "Select Columns"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "frmSelectColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SelectColumn")
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents Model As clsSelectColumnViewModel
Attribute Model.VB_VarHelpID = -1
Private Const SELECT_ALL As String = "(Select all)"

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

Private Sub lvColumns_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.key = SELECT_ALL Then
        DoSelectAll Item.Checked
    Else
        Model.TrySelect Item.key, Item.Checked
        Me.cmbOK.Enabled = CanContinue
        CheckSelectAll
    End If
End Sub

Private Sub CheckSelectAll()
    Dim allSelected As Boolean
    Dim li As ListItem
    
    allSelected = True
    For Each li In Me.lvColumns.ListItems
        If li.key <> SELECT_ALL Then
            If li.Checked = False Then
                allSelected = False
            End If
        End If
    Next
    
    Me.lvColumns.ListItems(SELECT_ALL).Checked = allSelected
End Sub

Private Sub DoSelectAll(ByVal state As Boolean)
    Dim li As ListItem
    For Each li In Me.lvColumns.ListItems
        If li.key <> SELECT_ALL Then
            Model.TrySelect li.key, state
        End If
    Next li
    Model_CollectionChanged
End Sub

Private Sub Model_CollectionChanged()
    Dim lc As ListColumn
    Dim chkCheck As Variant
    Dim li As ListItem
    
    Me.lvColumns.ListItems.Clear
    Me.lvColumns.ColumnHeaders.Clear
    Me.lvColumns.ColumnHeaders.Add Index:=1, text:="Name", Width:=120
    Me.lvColumns.ColumnHeaders.Add Index:=2, text:="Address", Width:=56

    Me.lvColumns.ListItems.Add key:=SELECT_ALL, text:=SELECT_ALL
    
    For Each lc In Model.Columns
        Set li = Me.lvColumns.ListItems.Add(key:=lc.name, text:=lc.name)
        Call li.ListSubItems.Add(text:="Column " & GetListColumnColumnAddress(lc))
        For Each chkCheck In Model.SelectedColumns
            If lc.name = chkCheck Then
                Me.lvColumns.ListItems(lc.name).Checked = True
            End If
        Next chkCheck
    Next lc
    
    CheckSelectAll
End Sub

Private Sub Model_ItemSelected()
    Me.cmbOK.Enabled = True
End Sub

Private Sub txtSearch_Change()
    Model.SearchCriteria = Me.txtSearch & "*"
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

Private Function CanContinue() As Boolean
    Dim li As ListItem
    For Each li In Me.lvColumns.ListItems
        If li.Checked = True Then
            CanContinue = True
            Exit Function
        End If
    Next li
End Function

Private Function GetListColumnColumnAddress(ByRef lc As ListColumn) As String
    GetListColumnColumnAddress = Split(lc.range.EntireColumn.Address(False, False), ":")(0)
End Function

