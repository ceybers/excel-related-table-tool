Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Private Const COMMAND_BAR As String = "List Range Popup"

Public Sub Main()
    MapTable
End Sub

Public Sub MapTable()
    With New clsRelatedTable
        Set .Worksheet = Selection.parent
        .ShowMapKeysDialog
    End With
    RegisterCommandBarButton
End Sub

Public Sub RegisterCommandBarButton()
    KillMacro
    AddCommandBarButton COMMAND_BAR, "Link to Related Column", "RelateColumn", 526
    AddCommandBarButton COMMAND_BAR, "Insert Related Columns", "RelateColumns", 530
End Sub

Private Sub KillMacro()
    TryDeleteCommandBarControl COMMAND_BAR, "Link to Related Column"
    TryDeleteCommandBarControl COMMAND_BAR, "Insert Related Columns"
End Sub

Public Sub RelateColumn()
    With New clsRelatedTable
        Set .Worksheet = Selection.parent
        .TryLinkColumn
    End With
End Sub

Public Sub RelateColumns()
    With New clsRelatedTable
        Set .Worksheet = Selection.parent
        .TryLinkColumns
    End With
End Sub

Private Sub AddCommandBarButton(ByVal commandBarName As String, ByVal caption As String, ByVal action As String, ByVal faceId As Long)
    With Application.CommandBars(commandBarName).Controls.Add(Temporary:=True)
        .caption = caption
        .Style = msoButtonIconAndCaption
        .OnAction = action
        .faceId = faceId
    End With
End Sub

Private Sub TryDeleteCommandBarControl(ByVal commandBarName As String, ByVal controlName As String)
    Dim cb As CommandBarControl
    Set cb = TryGetCommandBarControl(commandBarName, controlName)
    If Not cb Is Nothing Then
        cb.Delete
    End If
End Sub

Private Function TryGetCommandBarControl(ByVal commandBarName As String, ByVal controlName As String) As CommandBarControl
    Dim cb As CommandBarControl
    For Each cb In Application.CommandBars(commandBarName).Controls
        If cb.caption = controlName Then
            Set TryGetCommandBarControl = cb
            Exit Function
        End If
    Next cb
End Function

