VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterWithValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "AutoFilterWithValue" provides an implementation of "AutoFilterObject" based on if the actual value is matching the one provided value
' It implement "UpdateableAutoFilter" meaning the value can come from the row-context
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


Implements AutoFilterObject


Implements UpdateableAutoFilter


' Fields

Private pListObject As ListObject

Private pFilterColumn As ListColumn

Private pFilterValue As String


' Getters and Setters

Public Property Set ListObject(nListObject As ListObject)
    
    Set pListObject = nListObject
    
End Property


Public Property Get FilterColumn() As ListColumn
    
    Set FilterColumn = pFilterColumn
    
End Property


Public Property Let FilterColumn(nFilterColumn As Variant)

    Set pFilterColumn = LO_Support.LO_Finder.FindColumn(pListObject, nFilterColumn)
    
End Property


Public Property Let FilterValue(nFilterValue As Variant)

    pFilterValue = CStr(nFilterValue)
    
End Property


Private Property Set AutoFilterObject_ListObject(nListObject As ListObject)
    
    Set ListObject = nListObject
    
End Property


Private Property Get AutoFilterObject_FilterColumn() As ListColumn
    
    Set AutoFilterObject_FilterColumn = FilterColumn
    
End Property


Private Property Let AutoFilterObject_FilterColumn(nFilterColumn As Variant)
    
    FilterColumn = nFilterColumn
    
End Property


' Public API


Public Sub ApplyFilter()
    
    With pListObject
        .Range.AutoFilter Field:=pFilterColumn.Index, Criteria1:=pFilterValue
    End With
    
End Sub


Public Sub UpdateValue(v As Variant)
    
    FilterValue = v
    
End Sub


Private Sub AutoFilterObject_ApplyFilter()
    
    ApplyFilter
    
End Sub


Private Sub UpdateableAutoFilter_UpdateValue(v As Variant)
    
    UpdateValue v
    
End Sub


' Helpers



