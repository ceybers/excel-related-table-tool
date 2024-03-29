VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelectTableViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SelectTable")
Option Explicit
Option Compare Text

Public Event CollectionChanged()
Public Event ItemSelected()

Private Type TClsSelectTableViewModel
    SelectedTable As ListObject
    ActiveTable As ListObject
    Criteria As String
    Tables As Collection
End Type

Private this As TClsSelectTableViewModel

Public Property Get Tables() As Collection
    Set Tables = this.Tables
End Property

Public Property Get SelectedTable() As ListObject
    Set SelectedTable = this.SelectedTable
End Property

Public Property Set SelectedTable(ByVal RHS As ListObject)
    Set this.SelectedTable = RHS
    RaiseEvent ItemSelected
End Property

Public Property Get ActiveTable() As ListObject
    Set ActiveTable = this.ActiveTable
End Property

Public Property Set ActiveTable(ByVal RHS As ListObject)
    Set this.ActiveTable = RHS
    Load
End Property

Public Property Get Criteria() As String
    Criteria = this.Criteria
End Property

Public Property Let Criteria(ByVal RHS As String)
    this.Criteria = RHS
    Load
End Property

Private Sub Load()
    Set this.Tables = GetAllTables
    RaiseEvent CollectionChanged
End Sub

Private Function GetAllTables() As Collection
    Set GetAllTables = New Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            For Each lo In ws.ListObjects
                If this.Criteria = vbNullString Or lo.name Like this.Criteria Then
                    GetAllTables.Add lo, lo.range.Address(External:=True)
                End If
            Next lo
        Next ws
    Next wb
End Function

Public Sub TrySelect(ByVal key As String)
    Dim lo As ListObject
    For Each lo In this.Tables
        If lo.range.Address(External:=True) = key Then
            Set SelectedTable = lo
            Exit Sub
        End If
    Next lo
    Set SelectedTable = Nothing
End Sub

Private Sub Class_Initialize()
    Load
End Sub

