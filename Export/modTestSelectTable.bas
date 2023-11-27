Attribute VB_Name = "modTestSelectTable"
'@IgnoreModule
'@Folder "SelectTable"
Option Explicit

Public Sub TestSelectTable()
    Dim lo As ListObject
    Set lo = TrySelectTable
    
    If Not lo Is Nothing Then
        Debug.Print "You chose '" & lo.range.Address(External:=True) & "'"
    End If
End Sub

Private Function TrySelectTable() As ListObject
    Dim vm As clsSelectTableViewModel
    Dim view As frmSelectTable
        
    Set vm = New clsSelectTableViewModel
    Set view = New frmSelectTable
    
    Set vm.ActiveTable = Selection.ListObject
    
    If frmSelectTable.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set TrySelectTable = vm.SelectedTable
            
        End If
    End If
    
    Set view = Nothing
    Set vm = Nothing
End Function

