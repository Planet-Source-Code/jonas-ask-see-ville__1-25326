Attribute VB_Name = "LoadSave"
Option Explicit
Public Sub LoadGame(GameName)
Dim FileNum
Dim TempIn As String
Dim TempData As String
Dim NowPos As Integer
Dim Lengde As Integer
Dim Lengde2 As Integer
Dim L As Long
Dim a, X, Y As Integer
Dim temp
    FileNum = FreeFile
    On Error GoTo Error1
    
    Screen.MousePointer = 11
    NowPos = 1
    
    Open GameName For Input As FileNum
    Input #FileNum, TempIn
    

    'BYNAVN
    Lengde2 = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    Lengde = Mid(TempIn, NowPos, Lengde2)
    NowPos = NowPos + Lengde2
    cityinfo.CityName = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'INNBYGERE
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    cityinfo.Inhabitants = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'PENGER
    Lengde2 = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    Lengde = Mid(TempIn, NowPos, Lengde2)
    NowPos = NowPos + Lengde2
    cityinfo.Money = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'UKER
    cityinfo.Week = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    
    'MÅNEDER
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    cityinfo.Month = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'ÅR
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    cityinfo.Year = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'ÅRSTID
    cityinfo.Season = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    
    'ORDFØRER
    Lengde2 = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    Lengde = Mid(TempIn, NowPos, Lengde2)
    NowPos = NowPos + Lengde2
    cityinfo.MayorName = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'Bredde
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    Bredde = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'Hoyde
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    Hoyde = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'WstartX
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    WStartX = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    'WstartY
    Lengde = Mid(TempIn, NowPos, 1)
    NowPos = NowPos + 1
    WStartY = Mid(TempIn, NowPos, Lengde)
    NowPos = NowPos + Lengde
    
    ReDim BoardData(1 To Bredde, 1 To Hoyde)
    
    For L = 1 To Bredde * Hoyde
    
        Dim TempS As TileData
        Input #FileNum, TempIn
        NowPos = 0
        
        With TempS
            
            'POWER
            NowPos = NowPos + 1
            Select Case Mid(TempIn, NowPos, 1)
            Case "h": .Power = "1"
            Case "n": .Power = "0"
            Case "d": .Power = ""
            End Select
            NowPos = NowPos + 1
            
            'X
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            X = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            
            'BUILD
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            .Build = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            
            'BUILDTYPE
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            .BuildType = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            
            'CHILDREN
            For a = 1 To 3
                Lengde = Mid(TempIn, NowPos, 1)
                NowPos = NowPos + 1
                .Child(a).X = Mid(TempIn, NowPos, Lengde)
                NowPos = NowPos + Lengde
                Lengde = Mid(TempIn, NowPos, 1)
                NowPos = NowPos + 1
                temp = Mid(TempIn, NowPos, Lengde)
                .Child(a).Y = temp
                NowPos = NowPos + Lengde
            Next a
            
            'PARENT
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            .mParent.X = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            .mParent.Y = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            
            'SIZE
            .Size = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            
            'TER
            .Ter = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            
            'TERTYPE
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            .TerType = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
            
            'Y
            Lengde = Mid(TempIn, NowPos, 1)
            NowPos = NowPos + 1
            Y = Mid(TempIn, NowPos, Lengde)
            NowPos = NowPos + Lengde
                        
            
        End With
        With BoardData(X, Y)
            .Build = TempS.Build
            .BuildType = TempS.BuildType
            For a = 1 To 3
            .Child(a).X = TempS.Child(a).X
            .Child(a).Y = TempS.Child(a).Y
            Next
            .mParent.X = TempS.mParent.X
            .mParent.Y = TempS.mParent.Y
            .Power = TempS.Power
            .Size = TempS.Size
            .Ter = TempS.Ter
            .TerType = TempS.TerType
        End With
    Next L
    
    Close FileNum
    
    
    SetSeason (cityinfo.Season)
    UpdateData
    PaintGround
    Dirty = False
    MainPause = False
    Form1.lblPause.FontBold = False
    PaintMap 1, MiniMap
    Screen.MousePointer = 1
    
Exit Sub
Error1:
MsgBox Err.Description & vbNewLine & "The Game is not been loaded", vbOKOnly, GameTitle
Close FileNum
Screen.MousePointer = 1
End Sub

Public Sub SaveGame(GameName)
Dim Path As String
Dim FileNum
Dim Lengde
Dim Lengde2
Dim TempData
Dim Line1 As String
Dim ThisLine As String
    Path = GameName
    
    If Dir(Path) <> "" Then Kill Path
    
    Screen.MousePointer = 11
    
    'BYNAVN
    Lengde = Len(cityinfo.CityName)
    Lengde2 = Len(Lengde)
    Line1 = Line1 & Lengde2 & Lengde & cityinfo.CityName
    
    'INNBYGERE
    TempData = cityinfo.Inhabitants
    Lengde = Len(TempData)
    Line1 = Line1 & Lengde & cityinfo.Inhabitants
    
    'PENGER
    TempData = cityinfo.Money
    Lengde = Len(TempData)
    Lengde2 = Len(Lengde)
    Line1 = Line1 & Lengde2 & Lengde & cityinfo.Money
    
    'UKER
    Line1 = Line1 & cityinfo.Week
    
    'MÅNEDER
    Lengde = Len(cityinfo.Month)
    Line1 = Line1 & Lengde & cityinfo.Month
    
    'ÅR
    TempData = cityinfo.Year
    Lengde = Len(TempData)
    Line1 = Line1 & Lengde & cityinfo.Year
    
    'ÅRSTID
    Line1 = Line1 & cityinfo.Season
    
    'ORDFØRER
    Lengde = Len(cityinfo.MayorName)
    Lengde2 = Len(Lengde)
    Line1 = Line1 & Lengde2 & Lengde & cityinfo.MayorName
    
    'BREDDE
    Lengde = Len(Bredde)
    Line1 = Line1 & Lengde & Bredde
    
    'HOYDE
    Lengde = Len(Hoyde)
    Line1 = Line1 & Lengde & Hoyde
    
    'WstartX
    Lengde = Len(WStartX)
    Line1 = Line1 & Lengde & WStartX
    
    'WstartY
    Lengde = Len(WStartY)
    Line1 = Line1 & Lengde & WStartY
    
    FileNum = FreeFile
    Open Path For Output As FileNum
    Print #FileNum, Line1
    
    Dim X, Y, a As Integer
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            With BoardData(X, Y)
            ThisLine = ""
            
            'POWER
            Select Case .Power
            Case "": ThisLine = ThisLine & "d"
            Case "1": ThisLine = ThisLine & "h"
            Case "0": ThisLine = ThisLine & "n"
            End Select
            
            'X
            TempData = X
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & X
            
            'BUILD
            TempData = .Build
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & .Build
            
            'BUILDTYPE
            TempData = .BuildType
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & .BuildType
            
            'CHILDREN
            For a = 1 To 3
             TempData = .Child(a).X
             Lengde = Len(TempData)
             ThisLine = ThisLine & Lengde & .Child(a).X
             TempData = .Child(a).Y
             Lengde = Len(TempData)
             ThisLine = ThisLine & Lengde & .Child(a).Y
            Next a
            
            'PARENT
            TempData = .mParent.X
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & .mParent.X
            TempData = .mParent.Y
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & .mParent.Y
            
            'SIZE
            ThisLine = ThisLine & .Size
            
            'TER
            ThisLine = ThisLine & .Ter
            
            'TERTYPE
            TempData = .TerType
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & .TerType
            
            'Y
            TempData = Y
            Lengde = Len(TempData)
            ThisLine = ThisLine & Lengde & Y
            
            Print #FileNum, ThisLine
            
            End With
        Next X
    Next Y
    
    
    Dirty = False
    Close FileNum
    Screen.MousePointer = 1
End Sub
