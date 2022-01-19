Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Private Const NAMED_RANGE_FOR_KEY As String = "relTblKeyColumn"
Private Const NAMED_RANGE_FOR_TABLE As String = "relTblExternalRange"
Private viewModel As clsMapKeysViewModel
Private vmColumn As clsSelectColumnViewModel
Private targetColumn As ListColumn

Public Sub Main()
    Dim view As frmMapKeys
    Set view = New frmMapKeys
    
    Set viewModel = New clsMapKeysViewModel
    With viewModel
        'If Not view.ShowDialog(.Self) Then Exit Sub
        viewModel.PopulateTestValues
        If .IsValid Then KeepWorking
    End With
End Sub

Public Sub Main2(ByRef lc As ListColumn)
    Set targetColumn = lc
    
    Dim view As frmMapKeys
    Set view = New frmMapKeys
    
    Set viewModel = New clsMapKeysViewModel
    With viewModel
        'If Not view.ShowDialog(.Self) Then Exit Sub
        viewModel.PopulateTestValues
        If .IsValid Then KeepWorking
    End With
End Sub

Private Sub KeepWorking()
    'Dim wb As Workbook
    'Set wb = viewModel.SourceColumn.Parent.Parent.Parent
    'Debug.Print wb.FullName
    SetNameForSourceKeyColumn
    SetNameForExternalTable
    DisplaySelectColumnDialog
End Sub

Public Sub DisplaySelectColumnDialog()
    Dim nm As name
    Dim rng As range
    Dim lo As ListObject
    
    ' TODO Error trap this
    Set nm = ActiveSheet.Names(NAMED_RANGE_FOR_TABLE)
    'Debug.Print nm.RefersToRange.Address(external:=True)
    
    Set rng = nm.RefersToRange
    Set lo = rng.Cells(1, 1).ListObject
    
    Set vmColumn = New clsSelectColumnViewModel
    Set vmColumn.ListObject = lo
    
    Dim view As frmSelectColumn
    Set view = New frmSelectColumn
    
    Call frmSelectColumn.ShowDialog(vmColumn)
    'vmColumn.TrySelect "Airport"
    
    If vmColumn.IsValid Then
        ApplyFormulaToColumn
    End If
End Sub

Public Sub ApplyFormulaToColumn()
    'Debug.Print vmColumn.SelectedColumn.name & " was selected"
    Dim tgt As ListColumn
    Set tgt = ActiveSheet.ListObjects(1).ListColumns(2)
    If Not targetColumn Is Nothing Then
        Set tgt = targetColumn
    End If
    'tgt.DataBodyRange.Formula = "=2"
    
    Dim wb As Workbook
    Set wb = viewModel.SourceColumn.parent.parent.parent
    Dim localKey As String
    Dim externalKey As String
    Dim externalCol As String
    Dim lc As ListColumn
    Dim fx As String
    Set lc = vmColumn.SelectedColumn
    
    localKey = wb.Names(NAMED_RANGE_FOR_KEY).RefersToRange.Value
    externalKey = wb.Names(NAMED_RANGE_FOR_TABLE).RefersToRange.Address(external:=True)
    externalCol = lc.range.Address(external:=True)
    ' TODO Escape structured table reference charas
    'fx = "=XLOOKUP([@[" & localKey & "]], " & externalKey & ", " & externalCol & ", """")"
    fx = "=IFERROR(INDEX(" & externalCol & ",MATCH([@[" & localKey & "]], " & externalKey & ", 0)),"""")"
    tgt.DataBodyRange.Formula = fx
    tgt.DataBodyRange.NumberFormat = lc.DataBodyRange.Cells(1, 1).NumberFormat
    tgt.DataBodyRange.HorizontalAlignment = lc.DataBodyRange.Cells(1, 1).HorizontalAlignment
End Sub

Private Sub SetNameForExternalTable()
    Dim ws As Worksheet
    Dim rng As range
    Dim nm As name
    
    Set ws = viewModel.SourceColumn.parent.parent ' Names are stored in this worksheet, not the external table
    Set rng = viewModel.DestinationColumn.range
    
    If HasKeyNamedRange(ws, NAMED_RANGE_FOR_TABLE) Then
        Set nm = ws.Names(NAMED_RANGE_FOR_TABLE)
        If nm.RefersToRange.Address(external:=True) = rng.Address(external:=True) Then
            ' OK
        Else
            MsgBox "Replace existing external related table?", vbYesNo + vbDefaultButton2 + vbQuestion, "Map Key Columns"
        End If
    Else
        'AddKeyNamedRange rng, NAMED_RANGE_FOR_TABLE
        ws.Names.Add name:=NAMED_RANGE_FOR_TABLE, RefersTo:=("=" & rng.Address(external:=True)), Visible:=True
    End If
End Sub

Private Sub SetNameForSourceKeyColumn()
    Dim ws As Worksheet
    Dim rng As range
    Dim nm As name
    
    Set ws = viewModel.SourceColumn.parent.parent
    Set rng = viewModel.SourceColumn.range.Cells(1, 1)
    
    If HasKeyNamedRange(ws, NAMED_RANGE_FOR_KEY) Then
        Set nm = ws.Names(NAMED_RANGE_FOR_KEY)
        If nm.RefersToRange.Address = rng.Address Then
            ' OK
        Else
            MsgBox "Replace existing key column?", vbYesNo + vbDefaultButton2 + vbQuestion, "Map Key Columns"
        End If
    Else
        AddKeyNamedRange rng, NAMED_RANGE_FOR_KEY
    End If
End Sub

Private Sub AddKeyNamedRange(ByRef rng As range, ByVal mname As String)
    Dim ws As Worksheet
    Dim nm As name
    
    Set ws = rng.parent
    ws.Names.Add name:=mname, RefersTo:=("=" & rng.Address(external:=True)), Visible:=True
End Sub

Private Function HasKeyNamedRange(ByRef ws As Worksheet, ByVal Criteria As String) As Boolean
    Dim nm As name
    For Each nm In ws.Names
        If nm.name = (ws.name & "!" & Criteria) Then
            HasKeyNamedRange = True
            Exit Function
        End If
    Next nm
End Function

Public Sub UnhideNames()
    Dim nm As name
    For Each nm In ActiveSheet.Names
        nm.Visible = True
    Next nm
End Sub
