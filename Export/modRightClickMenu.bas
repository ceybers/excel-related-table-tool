Attribute VB_Name = "modRightClickMenu"
'@Folder("VBAProject")
Option Explicit

Private Const CONTROL_NAME As String = "Link to Related Column"
Private Const COMMAND_BAR As String = "List Range Popup"

Public Sub CreateMacro()
    KillMacro
   
    With Application.CommandBars(COMMAND_BAR).Controls.Add(Temporary:=True)
        .Caption = CONTROL_NAME
        .Style = msoButtonIconAndCaption
        .OnAction = "RelateColumn"
        .FaceId = 523 '216
    End With
End Sub

Private Sub KillMacro()
    On Error Resume Next
    Application.CommandBars(COMMAND_BAR).Controls(CONTROL_NAME).Delete
End Sub

Public Sub RelateColumn()
    With New clsRelatedTable
        Set .Worksheet = Selection.parent
        .TryLinkColumn
    End With
End Sub
