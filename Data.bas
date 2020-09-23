Attribute VB_Name = "Data"
Public Type ItemInfo
 Price As Integer
 ItemName As String
 Maint As Long
 Inhab As Integer
 Jobs As Integer
 Type As String
End Type

Public Type Buildings
 PowerPlants As Integer
 PoliceStat As Integer
 FireStat As Integer
 Roads As String
 Bridges As Integer
End Type

Public Type CityInf
 CityName As String
 Inhabitants As Long
 Buildings As Buildings
 Money As Long
 MayorName As String
 JobsC As Long
 JobsI As Long
 Week As Byte
 Month As Byte
 Year As Integer
 Season As Single
End Type

Public cityinfo As CityInf
Public infPowerlines As ItemInfo
Public infSmallRes As ItemInfo
Public infMedRes As ItemInfo
Public infSmallCom As ItemInfo
Public infMedCom As ItemInfo
Public infPowerPlant As ItemInfo
Public infRoads As ItemInfo
Public infBridge As ItemInfo
Public infTrees As ItemInfo
Public infSmallPark As ItemInfo
Public infBigPark As ItemInfo
Public infDemo As ItemInfo


Public Sub SetData()
    
    infPowerlines.ItemName = "Power Lines"
    infPowerlines.Price = 5
    
    'Small res
    infSmallRes.ItemName = "Small Residental"
    infSmallRes.Price = 50
    infSmallRes.Inhab = 4
    infSmallRes.Type = "r"
    
    'Medeum res
    infMedRes.ItemName = "Medium Residental"
    infMedRes.Price = 300
    infMedRes.Inhab = 18
    infMedRes.Type = "r"
    
    'Small com
    infSmallCom.ItemName = "Small Commercial"
    infSmallCom.Price = 70
    infSmallCom.Jobs = 3
    infSmallCom.Type = "c"
    
    'Medeum com
    infMedCom.ItemName = "Medium Commercial"
    infMedCom.Price = 500
    infMedCom.Jobs = 20
    infMedCom.Type = "c"
    
    ' The Powerplant
    infPowerPlant.ItemName = "Power Plant"
    infPowerPlant.Price = 2000
    infPowerPlant.Jobs = 100
    infPowerPlant.Type = "i"
    infPowerPlant.Maint = 70
    
    'The Road
    infRoads.ItemName = "Road"
    infRoads.Price = 10
    infRoads.Maint = 1
    
    'The Bridge
    infBridge.ItemName = "Bridge"
    infBridge.Price = 70
    infBridge.Maint = 10
    
    'Trees
    infTrees.ItemName = Vegitation
    infTrees.Price = 1
    
    'Parks
    infSmallPark.ItemName = "Small Park"
    infSmallPark.Maint = 5
    infSmallPark.Price = 20
    infBigPark.ItemName = "Big Park"
    infBigPark.Maint = 20
    infBigPark.Price = 100
    
    'Demolsih
    infDemo.Price = 2
End Sub
Public Sub CallDefaults()

    Bredde = 20
    Hoyde = 100
    WBredde = 16
    WHoyde = 16
    WStartX = 1
    WStartY = 1
    
    Form1.HScroll.Max = Bredde - WBredde
    Form1.VScroll.Max = Hoyde - WHoyde
    

End Sub

Public Sub NewGame(MapPath, CityName, MayorName, Diff)
    'Default values
    With cityinfo
     .Inhabitants = 0
     .Buildings.FireStat = 0
     .Buildings.PoliceStat = 0
     .Buildings.Roads = 0
     .Buildings.PowerPlants = 0
     .JobsC = 0
     .JobsI = 0
     .MayorName = MayorName
     Select Case Diff
     Case 1: .Money = 3500
     Case 2: .Money = 3000
     Case 3: .Money = 2500
     End Select
     .CityName = CityName
     .Week = 1
     .Month = 1
     .Year = 1975
     .Season = 4
    End With
    
    Loadmap MapPath
    
    Form1.VScroll.Max = Hoyde - WHoyde
    Form1.HScroll.Max = Bredde - WBredde
    
    WStartX = 1
    WStartY = 1
    Form1.VScroll.Value = 1
    Form1.HScroll.Value = 1
    
    UpdateData
    Dirty = False
    MainPause = False
    Form1.lblPause.FontBold = False
    
    PaintGround
    PaintlMap
    
End Sub
