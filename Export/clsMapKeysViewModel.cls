VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMapKeysViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MapKeys"
Option Explicit

Private Type TClsMapKeysViewModel
    Source As String
    Destination As String
    sourceColumn As ListColumn
    DestinationColumn As ListColumn
End Type

Private this As TClsMapKeysViewModel

' Properties
Public Property Get sourceColumn() As ListColumn
    Set sourceColumn = this.sourceColumn
End Property

Public Property Get DestinationColumn() As ListColumn
    Set DestinationColumn = this.DestinationColumn
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal RHS As String)
    this.Source = RHS
    Set this.sourceColumn = GetListColumnFromRangeText(RHS)
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property

Public Property Let Destination(ByVal RHS As String)
    this.Destination = RHS
    Set this.DestinationColumn = GetListColumnFromRangeText(RHS)
End Property

Public Property Get Self() As clsMapKeysViewModel
    Set Self = Me
End Property

' Private Functions
Public Function IsValid() As Boolean
    If (sourceColumn Is Nothing) Then Exit Function
    If (DestinationColumn Is Nothing) Then Exit Function
    IsValid = True
End Function

Private Function GetListColumnFromRangeText(ByVal rngText As String) As ListColumn
    On Error GoTo Catch
    Dim rng As range
    Dim lo As ListObject
    Dim lc As ListColumn
    
    If rngText = vbNullString Then Exit Function
    
    ' This will fail if rngText is not a valid range
    Set rng = range(rngText)
    
    ' Fail quietly if selection was not a single column
    If rng.Columns.Count <> 1 Then Exit Function
    
    ' Fail quietly if range was not inside a table
    If rng.Cells(1, 1).ListObject Is Nothing Then Exit Function
    
    Set lo = rng.Cells(1, 1).ListObject
    For Each lc In lo.ListColumns
        If lc.range.Column = rng.Cells(1, 1).Column Then
            Set GetListColumnFromRangeText = lc
            Exit Function
        End If
    Next lc
    
    GoTo Finally
    
Catch:
    If Err.Number = 1004 Then
        Debug.Print "rngText is invalid! '" & rngText & "'"
    Else
        Debug.Print "Error: " & Err.Number
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Function
Finally:
    
End Function

Public Sub PopulateTestValues()
    Source = "[Destination.xlsm]Sheet1!$A$1:$A$6"
    Destination = "[Source.xlsx]Sheet1!$E$1:$E$51"
End Sub
