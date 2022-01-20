Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Public Sub Main()
    MapTable
End Sub

Public Sub MapTable()
    With New clsRelatedTable
        Set .Worksheet = Selection.parent
        .ShowMapKeysDialog
    End With
End Sub
