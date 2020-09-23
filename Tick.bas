Attribute VB_Name = "Tick"
Public Type TickData
 WeekC As Integer
 UpdateGUIC As Integer
End Type
Public MainTick As TickData


Public Sub OneTick()
    If MainPause = True Then Exit Sub
    If MainPause Then
        Form1.lblPause.FontBold = True
    Else
        Form1.lblPause.FontBold = False
    End If
    
    
    DoPower
    
    Dirty = True
    
    With MainTick
        .WeekC = .WeekC + 1
        .UpdateGUIC = .UpdateGUIC + 1
        
        
        
        If .WeekC = 5 Then ' A WEEK
            .WeekC = 0
            CityInfo.Week = CityInfo.Week + 1
            
            
            
            If CityInfo.Week = 5 Then ' A MONTH
                DoLandvalues
                
                
                
                CityInfo.Week = 1
                CityInfo.Month = CityInfo.Month + 1
                If CityInfo.Month = 3 Then SetSeason 1
                If CityInfo.Month = 6 Then SetSeason 2
                If CityInfo.Month = 9 Then SetSeason 3
                If CityInfo.Month = 12 Then SetSeason 4
                If CityInfo.Month = 13 Then ' A YEAR
                    
                    
                    
                    CityInfo.Month = 1
                    CityInfo.Year = CityInfo.Year + 1
                    DoEconomics
                End If
            End If
        End If
        
        If .UpdateGUIC = 2 Then
            .UpdateGUIC = 0
            PaintGround
            Countpeople
        End If
        
    End With
    UpdateData
End Sub
