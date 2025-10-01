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

    Dim lcSource As ListColumn
    
    Dim lcTarget As ListColumn
    
    Dim flterTarget As AutoFilterObject
    
    
    For Each lcSource In pColumnMapping.MappingColumnsA
    
        Set lcTarget = pColumnMapping.MappingColumnForA(lcSource)
        
        Set flterTarget = pAutoFilterInteractor.FindAutoFilter(lcTarget)
        
        If Not flterTarget Is Nothing Then
            
            If TypeOf flterTarget Is AutoFilterWithValues Then
                
                Dim flterTargetValues As AutoFilterWithValues
                
                Set flterTargetValues = flterTarget
                
                With flterTargetValues
                    
                    .SetFilterValues CStr(LO_Support.LO_Finder.FindValue(pRowContext, lcSource))
                    
                End With
                
            ElseIf TypeOf flterTarget Is AutoFilterWithComparison Then
                
                Dim flterTargetComparison As AutoFilterWithComparison
                
                Set flterTargetComparison = flterTarget
                
                With flterTargetComparison
                    
                    .ComparisonValue = LO_Support.LO_Finder.FindValue(pRowContext, lcSource)
                    
                End With
                
            End If
            
        End If
    
    Next lcSource
    
    
    With pAutoFilterInteractor
        
        .Unfilter
        
        .ApplyAutoFilters
        
        .Focus
        
    End With
    
    
    pTarget.Parent.Activate ' Set focus to filtered worksheet

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


Public Sub AddAutoFilterWithValues(filterCol As Variant)
' Adds an "empty" auto-filter based on filter-values for column "filterCol"
' Empty in this context means, that the filter-value will come from "pRowContext"

    Dim flterWithValues As New AutoFilterWithValues
    
    With flterWithValues
        
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
    End With
    
    pAutoFilterInteractor.AddAutoFilter flter:=flterWithValues
    
End Sub


Public Sub AddAutoFilterWithComparison(filterCol As Variant, compType As ComparisonType)
' Adds an "empty" auto-filter based on comparison for column "filterCol" with comparison-type "compType"
' Empty in this context means, that the comparison-value will come from "pRowContext"
    
    Dim flterWithComparison As New AutoFilterWithComparison
    
    With flterWithComparison
    
        Set .ListObject = pTarget
        
        .FilterColumn = filterCol
        
        .Comparison = compType
        
    End With
    
    pAutoFilterInteractor.AddAutoFilter flter:=flterWithComparison
    
End Sub


' Helpers

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

