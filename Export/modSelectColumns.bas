Attribute VB_Name = "modSelectColumns"
'@Folder("SelectColumn")
Option Explicit

Public Sub AAATest()
    Dim vm As New clsSelectColumnViewModel
    Dim view As New frmSelectColumns
    Dim lo As ListObject
    
    Set lo = ActiveWorkbook.Worksheets(1).ListObjects(1)
    Set vm.ListObject = lo
    
    If view.ShowDialog(vm) Then
        If vm.IsValid Then
            Debug.Print "IsValid true"
            PrintResult vm
        Else
            Debug.Print "IsValid false"
        End If
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub

Private Sub PrintResult(ByRef vm As clsSelectColumnViewModel)
    Dim v As Variant
    For Each v In vm.SelectedColumns
        Debug.Print v.name
    Next v
End Sub
