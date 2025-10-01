VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LOAutoFilterInteractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "LOAutoFilterInteractor" provides methods for the visible auto-filter-interactions with one "ListObject"
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Fields

Private pListObject As ListObject

Private pAutoFilters As Collection ' Of AutoFilterObject


' Constructor

Private Sub Class_Initialize()

    Set pAutoFilters = New Collection ' Of AutoFilterObject

End Sub


' Getters and Setters

Public Property Set ListObject(nListObject As ListObject)

    Set pListObject = nListObject
    
End Property


' Public API

Public Sub Unfilter()
' Removes all auto-filters from "pListObject"
    
    activateAutoFilter
    
    removeAnyExistingAutoFilters
    
End Sub


Public Sub Focus()
' Sets the focus to the worksheet containing "pListObject"

    pListObject.Parent.Activate
    
End Sub


Public Sub ClearAutoFilters()
    
    Set pAutoFilters = New Collection ' Of AutoFilterObject
    
End Sub


Public Sub AddAutoFilter(flter As AutoFilterObject)
    
    If Not LO_Range_Interaction.IsSameListObject(pListObject, flter.FilterColumn.Parent) Then
    
        Err.Raise Number:=9, Description:="Filter must be for listobject '" & pListObject.Name & "'"
    
    ElseIf alreadyContainsFilterForColumn(flter.FilterColumn) Then
        
        Err.Raise Number:=9, Description:="Already contains a filter for listobject '" & flter.FilterColumn.Name & "'"
    
    End If
    
    pAutoFilters.Add Item:=flter
    
End Sub


Public Sub ApplyAutoFilters()
' Applies all auto-filters from "pAutoFilters"

    Dim flter As AutoFilterObject
    
    For Each flter In pAutoFilters
        flter.ApplyFilter
    Next flter
    
End Sub


Public Function FindAutoFilter(col As Variant) As AutoFilterObject
' Returns the auto-filter within "pAutoFilters" according to column "col"

    Dim lc As ListColumn
    
    Dim flter As AutoFilterObject
    
    Set lc = LO_Support.FindColumn(pListObject, col)
    
    For Each flter In pAutoFilters
        
        If flter.FilterColumn.Name = lc.Name Then
            Set FindAutoFilter = flter
        End If
        
    Next flter

End Function


' Helpers

Private Sub activateAutoFilter()
' Activates the autofilter for "pListObject", if it is not already set

    With pListObject
        If .AutoFilter Is Nothing Then .Range.AutoFilter
    End With
    
End Sub


Private Sub removeAnyExistingAutoFilters()
' Removes any auto-filters from "pListObject", if existing

    With pListObject.AutoFilter
        If .Filters.Count > 0 Then
            .ShowAllData
        End If
    End With
    
End Sub


Private Function alreadyContainsFilterForColumn(lc As ListColumn) As Boolean

    alreadyContainsFilterForColumn = (Not FindAutoFilter(lc) Is Nothing)

End Function
