VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelectColumnViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SelectColumn"
Option Explicit
Option Compare Text

Public Event CollectionChanged()
Public Event ItemSelected()

Private Type TClsSelectColumnViewModel
    ListObject As ListObject
    SelectedColumn As ListColumn
    SelectedColumns As Collection
    Columns As Collection
    SearchCriteria As String
    ReplaceColumnHeader As Boolean
End Type

Private this As TClsSelectColumnViewModel

' Properties
Public Property Get SelectedColumns() As Collection
    Set SelectedColumns = this.SelectedColumns
End Property

Public Property Get ReplaceColumnHeader() As Boolean
    ReplaceColumnHeader = this.ReplaceColumnHeader
End Property

Public Property Let ReplaceColumnHeader(ByVal RHS As Boolean)
    this.ReplaceColumnHeader = RHS
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = this.ListObject
End Property

Public Property Set ListObject(ByVal RHS As ListObject)
    Set this.ListObject = RHS
    PopulateColumns
End Property

Public Property Get SelectedColumn() As ListColumn
    Set SelectedColumn = this.SelectedColumn
End Property

Public Property Set SelectedColumn(ByVal RHS As ListColumn)
    Set this.SelectedColumn = RHS
    RaiseEvent ItemSelected
End Property

Public Property Get Columns() As Collection
    Set Columns = this.Columns
End Property

Public Property Set Columns(ByVal RHS As Collection)
    Set this.Columns = RHS
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = this.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal RHS As String)
    this.SearchCriteria = RHS
    PopulateColumns
End Property

' Methods
Public Function IsValid() As Boolean
    IsValid = Not this.SelectedColumn Is Nothing
End Function

Private Sub PopulateColumns()
    Dim lc As ListColumn
    
    If this.Columns Is Nothing Then
        Set this.Columns = New Collection
    Else
        ClearCollection this.Columns
    End If
    
    For Each lc In this.ListObject.ListColumns
        If this.SearchCriteria = vbNullString Then
            this.Columns.Add lc, lc.name
        ElseIf lc.name Like this.SearchCriteria Then
            this.Columns.Add lc, lc.name
        End If
    Next lc

    RaiseEvent CollectionChanged
End Sub

Public Sub SelectItem(ByVal key As String)
    MsgBox key
End Sub

Private Sub ClearCollection(ByVal coll As Collection)
    Dim i As Long
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
End Sub

Public Sub TrySelect(ByVal key As String, Optional ByVal selected As Boolean = True)
    Dim lc As ListColumn

    For Each lc In this.ListObject.ListColumns
        If lc.name = key Then
            Set SelectedColumn = lc
        End If
    Next lc
    
    If selected Then
        this.SelectedColumns.Add SelectedColumn
    Else
        Set this.SelectedColumns = RemoveFromCollection(this.SelectedColumns, SelectedColumn)
        Set SelectedColumn = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    Set this.SelectedColumns = New Collection
End Sub

Private Function RemoveFromCollection(ByRef coll As Collection, RHS As Variant) As Collection
    Set RemoveFromCollection = New Collection
    Dim v As Variant
    For Each v In coll
        If Not v Is RHS Then
            RemoveFromCollection.Add v
        End If
    Next v
End Function

