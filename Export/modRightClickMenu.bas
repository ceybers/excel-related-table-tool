Attribute VB_Name = "modRightClickMenu"
'@Folder("VBAProject")
Option Explicit

Private Const CONTROL_NAME As String = "RelateTable"
Private Const COMMAND_BAR As String = "List Range Popup"

Public Sub CreateMacro()
    Dim cBut
    Call KillMacro
   
    Set cBut = Application.CommandBars(COMMAND_BAR).Controls.Add(Temporary:=True)
    With cBut
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
    Dim lo As ListObject
    Dim lc As ListColumn
    Set lo = Selection.ListObject
    For Each lc In lo.ListColumns
        If lc.range.Column = Selection.Cells(1, 1).Column Then
            DoWork lc
        End If
    Next lc
End Sub

Private Sub DoWork(ByRef lc As ListColumn)
    modMain.Main2 lc
End Sub

