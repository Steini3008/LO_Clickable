Attribute VB_Name = "Factory"
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module "Factory" exposes the creation of several types of this VBA-Project to other VBA-Projects
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

Public Function CreateListObjectEvents() As ListObjectEvents

    Set CreateListObjectEvents = New ListObjectEvents
    
End Function


Public Function CreatePropagatedFiltering() As PropagatedFiltering

    Set CreatePropagatedFiltering = New PropagatedFiltering
    
End Function
