VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMappingOptionsViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'@Folder "MappingOptions"
Option Explicit

Public AddNewKeys As Boolean
Public ClearOldKeys As Boolean
Public RemoveOldKeys As Boolean
Public MapMatchedKeys As Boolean

Public HighlightAdditions As Boolean
Public HighlightRemovals As Boolean
Public HighlightMatches As Boolean

Public ClearDestinationColumns As Boolean
Public ClearReplaceEmptyOnly As Boolean
Public MapNonEmptyOnly As Boolean

Public RemoveDuplicatesInDestination As Boolean
