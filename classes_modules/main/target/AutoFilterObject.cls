VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Interface "AutoFilterObject" provides an API to apply an autofilter on one listcolumn on a listobject
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Getters and Setters

Public Property Set ListObject(nListObject As ListObject)

End Property


Public Property Get FilterColumn() As ListColumn

End Property


Public Property Let FilterColumn(nFilterColumn As Variant)

End Property


' Public API

Public Sub ApplyFilter()

End Sub
