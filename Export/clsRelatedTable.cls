VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelatedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Const NAMED_RANGE_FOR_KEY As String = "relTblKeyColumn"
Private Const NAMED_RANGE_FOR_TABLE As String = "relTblExternalRange"
Private vmMapKeys As clsMapKeysViewModel
Private vmColSel As clsSelectColumnViewModel

Private ws As Worksheet

' Properties
Public Property Set Worksheet(ByRef RHS As Variant)
    If RHS.ListObjects.Count = 0 Then
        MsgBox "There are no tables on this worksheet!", vbCritical + vbOKOnly, "Cannot map table"
    End If
    
    Set ws = RHS
    
    LoadSheetIntoMapKeysViewModel
    LoadListObjectIntoColSelViewModel
End Property

' Ctor
Private Sub Class_Initialize()
    Set vmMapKeys = New clsMapKeysViewModel
    Set vmColSel = New clsSelectColumnViewModel
End Sub

' Methods
Private Sub LoadSheetIntoMapKeysViewModel()
    If HasNamedRange(NAMED_RANGE_FOR_KEY) Then
        vmMapKeys.Destination = GetNamedRange(NAMED_RANGE_FOR_KEY).Address(External:=True)
    End If
    
    If HasNamedRange(NAMED_RANGE_FOR_TABLE) Then
        vmMapKeys.Source = GetNamedRange(NAMED_RANGE_FOR_TABLE).Address(External:=True)
    End If
End Sub

Private Sub LoadListObjectIntoColSelViewModel()
    If HasNamedRange(NAMED_RANGE_FOR_TABLE) Then
        Set vmColSel.ListObject = GetNamedRange(NAMED_RANGE_FOR_TABLE).ListObject
    End If
End Sub

Public Sub ShowMapKeysDialog()
    With New frmMapKeys
        If .ShowDialog(vmMapKeys) Then
            If vmMapKeys.IsValid Then
                SaveViewModel
            End If
        End If
    End With
End Sub

Public Sub TryLinkColumn()
    Dim targetColumn As ListColumn
    Set targetColumn = TryGetListColumnFromSelection(Selection)
    If targetColumn Is Nothing Then Exit Sub
    
    MsgBox "check if this worksheet has named ranges"
    
    If DoesListColumnHaveData(targetColumn) Then
        If vbNo = MsgBox("Replace existing data?", vbExclamation + vbYesNo + vbDefaultButton2, "Link Column") Then
            Exit Sub
        End If
    End If
    
    With New frmSelectColumn
        If .ShowDialog(vmColSel) Then
            If vmColSel.IsValid Then
                DoLinkColumn targetColumn
            End If
        End If
    End With
End Sub

Private Sub DoLinkColumn(ByRef targetColumn As ListColumn)
    Dim formula As String
    Dim localKey As String
    Dim externalKey As String
    Dim externalCol As String
    Dim srcCell As range
    
    localKey = GetNamedRange(NAMED_RANGE_FOR_KEY).Offset(-1, 0).Cells(1, 1).Value2
    externalKey = TryGetListColumnFromSelection(GetNamedRange(NAMED_RANGE_FOR_TABLE)).DataBodyRange.Address(External:=True)
    externalCol = vmColSel.SelectedColumn.DataBodyRange.Address(External:=True)
    Set srcCell = vmColSel.SelectedColumn.DataBodyRange.Cells(1, 1)
    
    formula = "=IFERROR(INDEX(" & externalCol & ",MATCH([@[" & localKey & "]], " & externalKey & ", 0)),"""")"
    
    With targetColumn.DataBodyRange
        .formula = formula
        .NumberFormat = srcCell.NumberFormat
        .HorizontalAlignment = srcCell.HorizontalAlignment
    End With
    
    If vmColSel.ReplaceColumnHeader Then
        targetColumn.name = vmColSel.SelectedColumn.name
    End If
End Sub

Private Function DoesListColumnHaveData(ByRef lc As ListColumn) As Boolean
    If lc Is Nothing Then Exit Function
    
    Dim rangeCount As Long
    Dim emptyCount As Long

    rangeCount = lc.DataBodyRange.Cells.Count
    emptyCount = CountBlanks(lc.DataBodyRange)

    DoesListColumnHaveData = rangeCount <> emptyCount
End Function

Private Function CountBlanks(ByRef rng As range) As Long
    Dim v As Variant
    Dim i As Long
    Dim j As Long
    
    v = rng.Value2
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            If VarType(v(i, j)) = vbEmpty Then
                CountBlanks = CountBlanks + 1
            End If
        Next j
    Next i
End Function

Private Function TryGetListColumnFromSelection(ByRef rng As range) As ListColumn
    If rng Is Nothing Then Exit Function
    If rng.Cells(1, 1).ListObject Is Nothing Then Exit Function
    If rng.Columns.Count <> 1 Then Exit Function
    
    Dim lo As ListObject
    Dim lc As ListColumn
    Set lo = rng.Cells(1, 1).ListObject
    For Each lc In lo.ListColumns
        If lc.range.Column = rng.Column Then
            Set TryGetListColumnFromSelection = lc
            Exit Function
        End If
    Next lc
End Function

Public Sub SaveViewModel()
    AddNamedRange NAMED_RANGE_FOR_KEY, vmMapKeys.DestinationColumn.DataBodyRange
    AddNamedRange NAMED_RANGE_FOR_TABLE, vmMapKeys.SourceColumn.DataBodyRange
End Sub

Private Function AddNamedRange(ByVal rangeName As String, ByRef rng As range) As range
    ws.Names.Add name:=rangeName, RefersTo:=("=" & rng.Address(External:=True)), Visible:=True
End Function

Private Function GetNamedRange(ByVal rangeName As String) As range
    If HasNamedRange(rangeName) = False Then Exit Function
    ' This will fail for external table's range if the file is not open.
    ' Need to try and open it
    On Error GoTo Catch
    Set GetNamedRange = ws.Names(ws.name & "!" & rangeName).RefersToRange
    GoTo Finally
    
Catch:
    If Err.Number = 1004 Then
        MsgBox "Please open the file first!"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
Finally:
    
End Function

Private Function HasNamedRange(ByVal rangeName As String) As Boolean
    Dim nm As name
    For Each nm In ws.Names
        If nm.name = (ws.name & "!" & rangeName) Then
            HasNamedRange = True
            Exit Function
        End If
    Next nm
End Function
