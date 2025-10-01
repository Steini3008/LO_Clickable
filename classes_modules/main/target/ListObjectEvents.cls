VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListObjectEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "ListObjectEvents" defines an event "CellWithinRelevantListObjectClicked" and raises this event, if a cell was double-clicked within
' a specific listobject.
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Event CellWithinRelevantListObjectClicked(rngClicked As Range)


' Fields

Private pRelevantListObject As ListObject

Private WithEvents pWorksheetContainingRelevantListObject As Worksheet
Attribute pWorksheetContainingRelevantListObject.VB_VarHelpID = -1


' Constructor

Private Sub Class_Initialize()

End Sub


' Getters and Setters

Public Property Set RelevantListObject(nRelevantListObject As ListObject)

    Set pRelevantListObject = nRelevantListObject
    
    Set pWorksheetContainingRelevantListObject = nRelevantListObject.Parent
    
End Property


' Events

Private Sub pWorksheetContainingRelevantListObject_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
' If a click happened in "pWorksheetContainingRelevantListObject", this method checks if it happened within "pRelevantListObject"
' If so, Raises event "CellWithinRelevantListObjectClicked"
    
    If isClickWithinRelevantListObject(Target) Then
    
        RaiseEvent CellWithinRelevantListObjectClicked(Target)
        
        Cancel = True
        
    End If
    
End Sub


' Helpers

Private Function isClickWithinRelevantListObject(rngClicked As Range) As Boolean
' Returns True, only if "rngClicked" is within "pRelevantListObject"
    
    isClickWithinRelevantListObject = LO_Range_Interaction.IsRangeWithinListObject(rngClicked, pRelevantListObject)
    
End Function
