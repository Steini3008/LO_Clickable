Attribute VB_Name = "ObjectProvider"
Option Explicit


' An object of type "ListObjectEvents" needs to exist in the global space to make sure, that
' events are getting raised as expected.
' "PropagatedFiltering" has a member from type "ListObjectEvents", who reacts on event "Worksheet_BeforeDoubleClick"
' by raising its own event "CellWithinRelevantListObjectClicked", which will be published and is subscribed to
' by "PropagatedFiltering"


Public FilterPropagator As PropagatedFiltering
