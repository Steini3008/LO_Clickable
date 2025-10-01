Attribute VB_Name = "LO_Range_Interaction"
Option Explicit
Option Private Module


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module "LO_Range_Interaction" provides simple functions to make interactions between "Range" and "ListObject" easier
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Public API

Public Function FindRowFromCell(rngFromRow As Range) As ListRow
' Returns the row, which "rngFromRow" contains

    With rngFromRow.ListObject
        
        Set FindRowFromCell = .ListRows(rngFromRow.Row - .HeaderRowRange.Row)
        
    End With
    
End Function


Public Function IsRangeWithinListObject(rng As Range, lo As ListObject) As Boolean
' Returns True, only if "rng" is within "lo"

    Dim loRange As ListObject
    
    Set loRange = rng.ListObject
    
    ' "rng" is not even in any listobject
    If loRange Is Nothing Then Exit Function
    
    IsRangeWithinListObject = IsSameListObject(loRange, lo)
    
End Function


Public Function IsSameListObject(loA As ListObject, loB As ListObject) As Boolean
' Returns True, only if "loA" and "loB" are objects to the same listobject

    IsSameListObject = (uniqueIdentifierForListObject(loA) = uniqueIdentifierForListObject(loB))
    
End Function


' Helpers

Private Function uniqueIdentifierForListObject(lo As ListObject) As String
' Builds the full path for "lo", containing the workbook-path and worksheet-name

    Dim ws As Worksheet
    
    Dim wb As Workbook
    
    Set ws = lo.Parent
    
    Set wb = ws.Parent
    
    uniqueIdentifierForListObject = wb.FullName & "\" & ws.Name & "\" & lo.Name
    
End Function
