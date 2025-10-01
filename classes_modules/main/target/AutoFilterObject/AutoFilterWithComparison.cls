VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterWithComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "AutoFilterWithComparison" provides an implementation of "AutoFilterObject" based on a comparison (e.g. Equals, GreaterThan, ...) to a value
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


' Fields

Private pListObject As ListObject

Private pFilterColumn As ListColumn

Private pComparisonTypeMapping As Scripting.Dictionary ' Key: ComparisonType, Value: String

Private pComparisonType As String

Private pComparisonValue As Variant


' Constructor

Private Sub Class_Initialize()
    
    setupComparisonTypeMapping
    
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


Public Property Let Comparison(nComparison As ComparisonType)
    
    pComparisonType = pComparisonTypeMapping.Item(nComparison)
    
End Property


Public Property Let ComparisonValue(nComparisonValue As Variant)

    pComparisonValue = nComparisonValue
    
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
        .Range.AutoFilter Field:=pFilterColumn.Index, Criteria1:=pComparisonType & pComparisonValue
    End With
    
End Sub


Private Sub AutoFilterObject_ApplyFilter()
    
    ApplyFilter
    
End Sub


' Helpers

Private Sub setupComparisonTypeMapping()

    Set pComparisonTypeMapping = New Scripting.Dictionary ' Key: ComparisonType, Value: String
    
    With pComparisonTypeMapping
        .Add Key:=ComparisonType.Equals, Item:="="
        .Add Key:=ComparisonType.NotEquals, Item:="<>"
        .Add Key:=ComparisonType.GreaterThan, Item:=">"
        .Add Key:=ComparisonType.GreaterThanOrEqualTo, Item:=">="
        .Add Key:=ComparisonType.LessThan, Item:="<"
        .Add Key:=ComparisonType.LessThanOrEqualTo, Item:="<="
    End With
    
End Sub
