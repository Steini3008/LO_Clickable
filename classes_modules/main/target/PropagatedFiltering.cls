VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropagatedFiltering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "PropagatedFiltering" uses data from one "ListObject" and by that filters another "ListObject"
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

Private pSource As ListObject

Private pTarget As ListObject

Private pColumnMapping As LO_Support.ColumnMapper

Private WithEvents pListObjectEventHandler As ListObjectEvents
Attribute pListObjectEventHandler.VB_VarHelpID = -1

Private pRowContext As ListRow

Private pAutoFilterInteractor As LOAutoFilterInteractor


' Constructor

Private Sub Class_Initialize()
    
    Set pColumnMapping = LO_Support.Factory.CreateColumnMapper
    
    Set pListObjectEventHandler = New ListObjectEvents
    
    Set pAutoFilterInteractor = New LOAutoFilterInteractor
    
End Sub


' Getters and Setters

Public Property Set Source(nSource As ListObject)
    
    setAllSourceListObjects nSource
    
End Property


Public Property Set Target(nTarget As ListObject)
    
    setAllTargetListObjects nTarget
    
End Property


Public Property Set RowContext(lr As ListRow)

    If Not LO_Range_Interaction.IsSameListObject(pSource, lr.Parent) Then
        Err.Raise Number:=9, Description:="ListRow must be from listobject '" & pSource.Name & "'"
    End If
    
    Set pRowContext = lr
    
End Property


' Events

Private Sub pListObjectEventHandler_CellWithinRelevantListObjectClicked(rngClicked As Range)
' Event "CellWithinRelevantListObjectClicked" sets the relevant listrow from the clicked cell "rngClicked" and triggers the filtering

    Set RowContext = LO_Range_Interaction.FindRowFromCell(rngClicked)
    
    FilterTarget
    
End Sub


' Public API

Public Sub FilterTarget()
' Filters "pTarget" based on the provided auto-filters within "pAutoFilterInteractor"
' Uses the data from "pRowContext" over the mapping "pColumnMapping"

    updateValuesFromAutoFilters
    
    useAutoFilters
    
    focusOnFilteredWorksheet

End Sub


Public Sub ClearColumnMapping()
' Delegates to "pColumnMapping"

    pColumnMapping.ClearMappingColumns
    
End Sub


Public Function AddColumnMappings(colsSource As Variant, colsTarget As Variant) As Boolean
' Delegates to "pColumnMapping"

    AddColumnMappings = pColumnMapping.AddMappingColumnPairs(colsSource, colsTarget)
    
End Function


Public Function SetColumnMappings(colsSource As Variant, colsTarget As Variant) As Boolean
' Delegates to "pColumnMapping"

    SetColumnMappings = pColumnMapping.SetMappingColumnPairs(colsSource, colsTarget)
    
End Function


Public Sub AddAutoFilter(flter As AutoFilterObject)
' Adds "flter" to "pAutoFilterInteractor"
' If you define and inject an auto-filter on your own, make sure that it is defined on listobject "pTarget"
' as it otherwise throws an error

    pAutoFilterInteractor.AddAutoFilter flter
    
End Sub


Public Sub AddAutoFilterWithValue(filterCol As Variant)
' Adds an "empty" auto-filter to "pAutoFilterInteractor" based on a filter-value for column "filterCol"
' Empty in this context means, that the filter-value will come from "pRowContext"

    Dim flterWithValue As New AutoFilterWithValue
    
    With flterWithValue
        
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
    End With
    
    AddAutoFilter flterWithValue
    
End Sub


Public Sub AddAutoFilterWithFixedValues(filterCol As Variant, ParamArray filterValues() As Variant)
' Adds an "fully set-up" auto-filter to "pAutoFilterInteractor" based on filter-values for column "filterCol"
' Fully set-up as "AutoFilterWithFixedValues" is not updateable

    Dim flterWithValues As New AutoFilterWithFixedValues
    
    With flterWithValues
        
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
        .SetFilterValues filterValues
        
    End With
    
    AddAutoFilter flterWithValues
    
End Sub


Public Sub AddAutoFilterWithDynamicValues(filterCol As Variant, ParamArray filterValues() As Variant)
' Adds an "empty" auto-filter to "pAutoFilterInteractor" based on filter-values for column "filterCol"
' Empty in this context means, that one additional filter-value will come from "pRowContext"

    Dim flterWithValues As New AutoFilterWithDynamicValues
    
    With flterWithValues
        
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
        .SetFilterValues filterValues
        
    End With
    
    AddAutoFilter flterWithValues
    
End Sub


Public Sub AddAutoFilterWithComparison(filterCol As Variant, compType As ComparisonType)
' Adds an "empty" auto-filter to "pAutoFilterInteractor" based on comparison for column "filterCol" with comparison-type "compType"
' Empty in this context means, that the comparison-value will come from "pRowContext"
    
    Dim flterWithComparison As New AutoFilterWithComparison
    
    With flterWithComparison
    
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
        .Comparison = compType
        
    End With
    
    AddAutoFilter flterWithComparison
    
End Sub


' Helpers

Private Sub updateValuesFromAutoFilters()
' Iterates over auto-filters within "pAutoFilterInteractor" and sets/updates values

    Dim flterTarget As AutoFilterObject
    
    For Each flterTarget In pAutoFilterInteractor.AutoFilters
        
        updateValueFromAutoFilter flterTarget
        
    Next flterTarget
    
End Sub


Private Sub updateValueFromAutoFilter(flter As AutoFilterObject)
' For auto-filter "flter", it sets/updates its value with the content from the mapped column "lcSource" within "pRowContext"

    Dim lcSource As ListColumn
    
    Set lcSource = pColumnMapping.MappingColumnForB(flter.FilterColumn)
    
    updateAutoFilterWithValue flter, LO_Support.LO_Finder.FindValue(pRowContext, lcSource)

End Sub


Private Sub updateAutoFilterWithValue(flter As AutoFilterObject, v As Variant)
' For auto-filter "flter", it sets/updates its value with "v"

    If TypeOf flter Is UpdateableAutoFilter Then
        
        Dim updateableFilter As UpdateableAutoFilter
        
        Set updateableFilter = flter
        
        updateableFilter.UpdateValue v
        
    End If

End Sub


Private Sub useAutoFilters()
' Delegates to "pAutoFilterInteractor" by first clearing and then applying filters

    With pAutoFilterInteractor
        
        .Unfilter
        
        .ApplyAutoFilters
        
    End With
    
End Sub


Private Sub focusOnFilteredWorksheet()

    pTarget.Parent.Activate
    
End Sub


Private Sub setAllSourceListObjects(lo As ListObject)

    Set pSource = lo
    
    Set pColumnMapping.ListObjectA = lo
    
    Set pListObjectEventHandler.RelevantListObject = lo
    
End Sub


Private Sub setAllTargetListObjects(lo As ListObject)

    Set pTarget = lo
    
    Set pColumnMapping.ListObjectB = lo
    
    Set pAutoFilterInteractor.ListObject = lo
    
End Sub

