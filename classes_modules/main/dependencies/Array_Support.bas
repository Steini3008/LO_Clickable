Attribute VB_Name = "Array_Support"
Option Explicit
Option Private Module


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module "Array_Support" provides useful functions to deal with type "Variant()"
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

Public Function FromCollection(c As Collection) As Variant()
' Returns an array from "c"
' Special: Casts every item from "c" to "String" as it is currently only used within the context of "AutoFilters"
    
    Dim arr() As Variant
    
    Dim i As Integer
    
    
    ReDim arr(0 To c.Count - 1) As Variant
    
    For i = 1 To c.Count
        
        arr(i - 1) = CStr(c.Item(i))
        
    Next i
    
    FromCollection = arr
    
End Function


' Helpers

