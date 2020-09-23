Attribute VB_Name = "Economics"
Public TaxR As Byte
Public TaxC As Byte
Public TaxI As Byte
Public Sub DoEconomics()
    CountStuff
    GetTaxes
    PayMaintenance
End Sub

Public Sub CountStuff()
Dim tType As Integer
    CityInfo.Buildings.PowerPlants = 0
    CityInfo.Buildings.Roads = 0
    CityInfo.Buildings.Bridges = 0
    CityInfo.JobsC = 0
    CityInfo.JobsI = 0
    
    Countpeople
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            tType = BoardData(X, Y).Build
            If tType = 0 Then GoTo NextX
            If BoardData(X, Y).Power = "0" Then GoTo NextX 'Hvis ikke størm
            
            Select Case tType
            Case 1 'residential
            Case 2
                
            Case 3
                
            Case 5 'Powerplant
                CityInfo.Buildings.PowerPlants = CityInfo.Buildings.PowerPlants + 1
            Case 10 'Roads
                Select Case BoardData(X, Y).BuildType
                Case 19 To 24: CityInfo.Buildings.Bridges = CityInfo.Buildings.Bridges + 1
                Case Else: CityInfo.Buildings.Roads = CityInfo.Buildings.Roads + 1
                End Select
            End Select
NextX:
        Next X
    Next Y
    
    'fix Those on more than one square
    CityInfo.Buildings.PowerPlants = CityInfo.Buildings.PowerPlants / 4
End Sub
Public Sub Countpeople()
Dim tType As Integer
    CityInfo.Inhabitants = 0
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            tType = BoardData(X, Y).Build
            If tType = 0 Then GoTo NextX
            If BoardData(X, Y).Power = "0" Then GoTo NextX 'Hvis ikke størm
            If Not BoardData(X, Y).Build = 1 Then GoTo NextX 'Hvis ikke Resident
            
            Select Case BoardData(X, Y).BuildType
            Case 1 To 10: CityInfo.Inhabitants = CityInfo.Inhabitants + infSmallRes.Inhab
            Case 11 To 20: CityInfo.Inhabitants = CityInfo.Inhabitants + infMedRes.Inhab
            End Select
NextX:
        Next X
    Next Y
End Sub

Public Sub GetTaxes()
Dim Taxes As Currency
Dim IncomeR As Currency
Dim IncomeC As Currency
Dim IncomeI As Currency

TaxR = 30
    IncomeR = CityInfo.Inhabitants * ((TaxR * 40) / 100)
    
    Taxes = Taxes + IncomeR + IncomeC + IncomeI
    
    CityInfo.Money = CityInfo.Money + Taxes
End Sub
Public Sub PayMaintenance()
Dim lPrice As Currency
    'Powerplant
    lPrice = lPrice + CityInfo.Buildings.PowerPlants * infPowerPlant.Maint
    'Roads
    lPrice = lPrice + CityInfo.Buildings.Roads * infRoads.Maint
    'Bridges
    lPrice = lPrice + CityInfo.Buildings.Bridges * infBridge.Maint
    CityInfo.Money = CityInfo.Money - lPrice
End Sub
