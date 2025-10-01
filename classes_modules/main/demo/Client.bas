Attribute VB_Name = "Client"
Option Explicit


Private Property Get loPeopleData() As ListObject

    Set loPeopleData = wsPeople.ListObjects("PeopleData")

End Property


Private Property Get loCountryInfo() As ListObject

    Set loCountryInfo = wsCountry.ListObjects("LandesInformationen")

End Property


Public Sub Main_PropagatedFiltering()
    
    Set ObjectProvider.FilterPropagator = New PropagatedFiltering
    
    With ObjectProvider.FilterPropagator
        
        Set .Source = loPeopleData
        
        Set .Target = loCountryInfo
        
        .AddColumnMappings Array("Country", "Age", "Gender"), Array("Land", "Alter", "Geschlecht")
        
        .AddAutoFilterWithComparison "Land", ComparisonType.Equals
        
        .AddAutoFilterWithValues "Alter"
        
    End With
    
End Sub
