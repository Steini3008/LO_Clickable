VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterWithFixedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "AutoFilterWithFixedValues" provides an implementation of "AutoFilterObject" based on if the actual value is contained within the listed values
' It is not implementing "UpdateableAutoFilter" meaning it is unaffected by the row-context
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   "Collection_Support"
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


Implements AutoFilterObject


' Fields

Private pListObject As ListObject

Private pFilterColumn As ListColumn

Private pFilterValues As Collection


' Constructor

Private Sub Class_Initialize()
    
    Set pFilterValues = New Collection
    
End Sub


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

Public Sub SetFilterValues(ParamArray paFilterValues() As Variant)
' "paFilterValues" need to be from type String, as other types (e.g. Integer/Long) produce issues
    
    Set pFilterValues = Collection_Support.Collection_Creator.Flattened(paFilterValues)
    
End Sub


Public Sub AppendFilterValues(ParamArray paFilterValues() As Variant)
    
    Collection_Support.Collection_Manipulator.AppendFlattened pFilterValues, Collection_Support.Collection_Creator.Flattened(paFilterValues)
    
End Sub


Public Sub ApplyFilter()
    
    With pListObject
        .Range.AutoFilter Field:=pFilterColumn.Index, Operator:=xlFilterValues, Criteria1:=Array_Support.FromCollection(pFilterValues)
    End With
    
End Sub


Private Sub AutoFilterObject_ApplyFilter()
    
    ApplyFilter
    
End Sub


' Helpers

