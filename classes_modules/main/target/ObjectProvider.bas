Attribute VB_Name = "ObjectProvider"
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module "ObjectProvider" exposes objects, so they are not getting destroyed immediately after a procedure-call.
'
' An object of type "ListObjectEvents" needs to exist in the global space to make sure, that
' its events (e.g. "CellWithinRelevantListObjectClicked") are getting raised as expected.
' "PropagatedFiltering" has a member from type "ListObjectEvents", who reacts on event "Worksheet_BeforeDoubleClick"
' by raising its own event "CellWithinRelevantListObjectClicked", which will be published and is subscribed to by "PropagatedFiltering"
' So the object "FilterPropagator" keeps an instance of type "ListObjectEvents" alive.
'
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


Public FilterPropagator As PropagatedFiltering
