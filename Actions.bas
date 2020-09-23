Attribute VB_Name = "Actions"
Public Type Directions
 o As Boolean
 n As Boolean
 h As Boolean
 v As Boolean
End Type

Public Sub Demolish(X, Y)
    Select Case BoardData(X, Y).Size
    Case 1
        Demolish1x1 X, Y
    Case 2
        Demolish2x2 X, Y
    Case 3
        'Demolish3x3 X, Y
    End Select
End Sub
Public Sub BuildNewResSml(mX, mY, btype)
    Build1x1 mX, mY, 1, btype, "1"
End Sub

Public Sub BuildNewResMed(mX, mY, btype)
    Build2x2 mX, mY, 1, btype, "1"
End Sub
Public Sub BuildNewComSml(mX, mY, btype)
    Build1x1 mX, mY, 2, btype, "1"
End Sub
Public Sub BuildNewComMed(mX, mY, btype)
    Build1x1 mX, mY, 2, btype, "1"
End Sub
Public Sub BuildNewPlant(mX, mY, btype)
    Build2x2 mX, mY, 5, 1, "1"
End Sub
Public Sub BuildPark1(mX, mY, btype)
    Build1x1 mX, mY, 4, btype, ""
End Sub
Public Sub BuildPark2(mX, mY, btype)
    Build2x2 mX, mY, 4, btype, ""
End Sub

Public Sub BuildTree(mX, mY)
    If Not BoardData(mX, mY).Build = 0 Then Exit Sub
    If BoardData(mX, mY).Ter = 0 Then Exit Sub
    If Not FixMoney(infTrees.Price) Then Exit Sub
    
    BoardData(mX, mY).Build = 0
    BoardData(mX, mY).Size = 1
    
    Select Case BoardData(mX, mY).BuildType
    Case 0
        BoardData(mX, mY).BuildType = 0 + RndTall(1, 6)
    Case 1 To 4
        BoardData(mX, mY).BuildType = 4 + RndTall(1, 4)
    Case 5 To 8
        BoardData(mX, mY).BuildType = 8 + RndTall(1, 4)
    Case 9 To 12
        BoardData(mX, mY).BuildType = 12 + RndTall(1, 4)
    Case 13 To 16
        Exit Sub
    End Select
    
End Sub

Public Sub Enquire(X, Y)
Dim tX As Integer, tY As Integer
    With frmInfo
        .picIcon.Cls
        If Not BoardData(X, Y).mParent.X = 0 And Not BoardData(X, Y).mParent.Y = 0 Then
            If BoardData(X, Y).Build = 10 Then GoTo AsNormal 'THIS WILL NOT HAPPEN TO A BRIDGE
            tX = X
            tY = Y
            X = BoardData(tX, tY).mParent.X
            Y = BoardData(tX, tY).mParent.Y
        End If
AsNormal:

        Select Case BoardData(X, Y).Size 'Tilpass sizen på bilde
        Case 1, 0: .picIcon.Top = 300: .picIcon.Left = 1500 + 300
        Case 2: .picIcon.Top = 150: .picIcon.Left = 1500 + 150
        Case 3: .picIcon.Top = 0: .picIcon.Left = 1500 + 0
        End Select
        
        
        Select Case BoardData(X, Y).Build
        Case 1 'Residental
            Select Case BoardData(X, Y).BuildType
            Case 1 To 10
                .lblName = infSmallRes.ItemName
                .lblMaint = "Yearly Maintenance: " & infSmallRes.Maint
            Case 11 To 20
                .lblName = infMedRes.ItemName
                .lblMaint = "Yearly Maintenance: " & infMedRes.Maint
            End Select
        Case 2 'Commerce
            Select Case BoardData(X, Y).BuildType
            Case 1 To 10
                .lblName = infSmallCom.ItemName
                .lblMaint = "Yearly Maintenance: " & infSmallCom.Maint
            Case 11 To 20
                .lblName = infMedCom.ItemName
                .lblMaint = "Yearly Maintenance: " & infMedCom.Maint
            End Select
        Case 4 'Parks
            Select Case BoardData(X, Y).BuildType
            Case 1 To 5
                .lblName = infSmallPark.ItemName
                .lblMaint = "Yearly Maintenance: " & infSmallPark.Maint
            Case 6 To 10
                .lblName = infBigPark.ItemName
                .lblMaint = "Yearly Maintenance: " & infBigPark.Maint
            End Select
        Case 5 'PowerPlants
            .lblName = infPowerPlant.ItemName
            .lblMaint = "Yearly Maintenance: " & infPowerPlant.Maint
        Case 9 'Powerlines
            .lblName = infPowerlines.ItemName
            .lblMaint = "Yearly Maintenance: " & infPowerlines.Maint
        Case 10 'Roads
            Select Case BoardData(X, Y).BuildType
            Case 19 To 24
                .lblName = infBridge.ItemName
                .lblMaint = "Yearly Maintenance: " & infBridge.Maint
            Case Else
                .lblName = infRoads.ItemName
                .lblMaint = "Yearly Maintenance: " & infRoads.Maint
            End Select
        Case 0 'Empty land
            .lblName = "Empty land"
            BitBlt .picIcon.hDC, 0, 0, Size, Size, BufferG.hDC, (X - WStartX) * Size, (Y - WStartY) * Size, SRCCOPY
            .lblMaint = "Yearly Maintenance: 0"
        Case Else
            Exit Sub
        End Select
        
        
        If BoardData(X, Y).Power = "1" Then
            .lblPower = "Power: Yes"
        Else
            .lblPower = "Power: No"
        End If
        
        If BoardData(X, Y).Ter = 1 Then
            .lblTer = "Terrain: Land"
        ElseIf BoardData(X, Y).Ter = 0 Then
            .lblTer = "Terrain: Water"
        End If
        
        .lblValue = "Landvalue: " & BoardData(X, Y).LandVal
        
        .Show , Form1
        BitBlt .picIcon.hDC, 0, 0, Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, BufferM.hDC, (X - WStartX) * Size, (Y - WStartY) * Size, SRCAND
        BitBlt .picIcon.hDC, 0, 0, Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, BufferS.hDC, (X - WStartX) * Size, (Y - WStartY) * Size, SRCPAINT
    End With
End Sub
Public Sub DemolishBridge(X, Y)
Dim tX, tY
Dim DeltaX As Integer, DeltaY As Integer
Dim sX, sY, gX, gY
Dim StepSize As Integer

    sX = X
    sY = Y
    If Not BoardData(X, Y).mParent.X = 0 And Not BoardData(X, Y).mParent.Y = 0 Then
        sX = BoardData(X, Y).mParent.X
        sY = BoardData(X, Y).mParent.Y
    End If
    gX = BoardData(sX, sY).Child(1).X
    gY = BoardData(sX, sY).Child(1).Y
    
    DeltaX = sX - gX
    DeltaY = sY - gY
    
    
    If sY > gY Then StepSize = -1 Else StepSize = 1
    If DeltaX = 0 Then
        For Y = sY To gY Step StepSize
            BoardData(sX, Y).Build = 0
            BoardData(sX, Y).Power = ""
            BoardData(sX, Y).BuildType = 0
            BoardData(sX, Y).mParent.X = 0
            BoardData(sX, Y).mParent.Y = 0
            BoardData(sX, Y).Size = 0
        Next Y
        If BoardData(sX, sY - 1).Build = 10 Then BoardData(sX, sY - 1).BuildType = getRoadType(sX, sY - 1, "road")
        If BoardData(sX, sY + 1).Build = 10 Then BoardData(sX, sY + 1).BuildType = getRoadType(sX, sY + 1, "road")
        If BoardData(gX, gY - 1).Build = 10 Then BoardData(gX, gY - 1).BuildType = getRoadType(gX, gY - 1, "road")
        If BoardData(gX, gY + 1).Build = 10 Then BoardData(gX, gY + 1).BuildType = getRoadType(gX, gY + 1, "road")
    Else
        For X = sX To gX Step StepSize
            BoardData(X, sY).Build = 0
            BoardData(X, sY).Power = ""
            BoardData(X, sY).BuildType = 0
            BoardData(X, sY).mParent.X = 0
            BoardData(X, sY).mParent.Y = 0
            BoardData(X, sY).Size = 0
        Next X
        If BoardData(sX - 1, sY).Build = 10 Then BoardData(sX - 1, sY).BuildType = getRoadType(sX - 1, sY, "road")
        If BoardData(sX + 1, sY).Build = 10 Then BoardData(sX + 1, sY).BuildType = getRoadType(sX + 1, sY, "road")
        If BoardData(gX - 1, gY).Build = 10 Then BoardData(gX - 1, gY).BuildType = getRoadType(gX - 1, gY, "road")
        If BoardData(gX + 1, gY).Build = 10 Then BoardData(gX + 1, gY).BuildType = getRoadType(gX + 1, gY, "road")
    End If

    BoardData(sX, sY).Child(1).X = 0
    BoardData(sX, sY).Child(1).Y = 0
        If BoardData(sX - 1, sY).Build = 9 Then BoardData(sX - 1, sY).BuildType = getRoadType(sX - 1, sY, "power")
        If BoardData(sX, sY + 1).Build = 9 Then BoardData(sX, sY + 1).BuildType = getRoadType(sX, sY + 1, "power")
        If BoardData(sX + 1, sY).Build = 9 Then BoardData(sX + 1, sY).BuildType = getRoadType(sX + 1, sY, "power")
        If BoardData(sX, sY - 1).Build = 9 Then BoardData(sX, sY - 1).BuildType = getRoadType(sX, sY - 1, "power")
        
        If BoardData(gX - 1, gY).Build = 9 Then BoardData(gX - 1, gY).BuildType = getRoadType(gX - 1, gY, "power")
        If BoardData(gX, gY + 1).Build = 9 Then BoardData(gX, gY + 1).BuildType = getRoadType(gX, gY + 1, "power")
        If BoardData(gX + 1, gY).Build = 9 Then BoardData(gX + 1, gY).BuildType = getRoadType(gX + 1, gY, "power")
        If BoardData(gX, gY - 1).Build = 9 Then BoardData(gX, gY - 1).BuildType = getRoadType(gX, gY - 1, "power")
End Sub


Public Sub Demolish1x1(X, Y)
Dim Itwas As Integer
    If BoardData(X, Y).Build = 0 And BoardData(X, Y).BuildType = 0 Then Exit Sub
    If Not FixMoney(infDemo.Price) Then Exit Sub
    
    Itwas = BoardData(X, Y).Build
    Select Case BoardData(X, Y).BuildType
    Case 19 To 24
        DemolishBridge X, Y
        Exit Sub
    End Select
    BoardData(X, Y).Build = 0
    BoardData(X, Y).BuildType = 0
    BoardData(X, Y).Size = 0
    BoardData(X, Y).Power = ""
    
    If Itwas = 10 Then 'Road
        If BoardData(X - 1, Y).Build = 10 Then BoardData(X - 1, Y).BuildType = getRoadType(X - 1, Y, "road")
        If BoardData(X, Y + 1).Build = 10 Then BoardData(X, Y + 1).BuildType = getRoadType(X, Y + 1, "road")
        If BoardData(X + 1, Y).Build = 10 Then BoardData(X + 1, Y).BuildType = getRoadType(X + 1, Y, "road")
        If BoardData(X, Y - 1).Build = 10 Then BoardData(X, Y - 1).BuildType = getRoadType(X, Y - 1, "road")
    End If
    
    If BoardData(X - 1, Y).Build = 9 Then BoardData(X - 1, Y).BuildType = getRoadType(X - 1, Y, "power")
    If BoardData(X, Y + 1).Build = 9 Then BoardData(X, Y + 1).BuildType = getRoadType(X, Y + 1, "power")
    If BoardData(X + 1, Y).Build = 9 Then BoardData(X + 1, Y).BuildType = getRoadType(X + 1, Y, "power")
    If BoardData(X, Y - 1).Build = 9 Then BoardData(X, Y - 1).BuildType = getRoadType(X, Y - 1, "power")
    
End Sub
Public Sub Demolish2x2(X, Y)
Dim temp As COOR
    If BoardData(X, Y).Build = 0 Then Exit Sub
    
    If Not BoardData(X, Y).mParent.X = 0 And Not BoardData(X, Y).mParent.Y = 0 Then ' IF CHILD
        Demolish BoardData(X, Y).mParent.X, BoardData(X, Y).mParent.Y
        BoardData(X, Y).mParent.X = 0
        BoardData(X, Y).mParent.Y = 0
        Exit Sub
    End If
    
    If Not FixMoney(infDemo.Price * 4) Then Exit Sub
    
    If Not BoardData(X, Y).Child(1).X = 0 Then ' IF PARENT
        temp.X = X
        temp.Y = Y
        Dim lX, lY As Integer
        For a = 1 To 3
        'every Child
            lX = BoardData(temp.X, temp.Y).Child(a).X: lY = BoardData(temp.X, temp.Y).Child(a).Y
            With BoardData(lX, lY)
                .Build = 0
                .Size = 0
                .mParent.X = 0
                .mParent.Y = 0
                .BuildType = 0
                .Power = ""
            End With
            'slette Children fra Parent
            BoardData(temp.X, temp.Y).Child(a).X = 0
            BoardData(temp.X, temp.Y).Child(a).Y = 0
        Next a
        
        X = temp.X
        Y = temp.Y
    End If

    BoardData(X, Y).Build = 0
    BoardData(X, Y).Size = 0
    BoardData(X, Y).BuildType = 0
    BoardData(X, Y).Power = ""
    
    If BoardData(X - 1, Y).Build = 9 Then BoardData(X - 1, Y).BuildType = getRoadType(X - 1, Y, "power"): PaintGround
    If BoardData(X - 1, Y + 1).Build = 9 Then BoardData(X - 1, Y + 1).BuildType = getRoadType(X - 1, Y + 1, "power"): PaintGround
    If BoardData(X, Y + 2).Build = 9 Then BoardData(X, Y + 2).BuildType = getRoadType(X, Y + 2, "power"): PaintGround
    If BoardData(X + 1, Y + 2).Build = 9 Then BoardData(X + 1, Y + 2).BuildType = getRoadType(X + 1, Y + 2, "power"): PaintGround
    If BoardData(X + 2, Y + 1).Build = 9 Then BoardData(X + 2, Y + 1).BuildType = getRoadType(X + 2, Y + 1, "power"): PaintGround
    If BoardData(X + 2, Y).Build = 9 Then BoardData(X + 2, Y).BuildType = getRoadType(X + 2, Y, "power"): PaintGround
    If BoardData(X + 1, Y - 1).Build = 9 Then BoardData(X + 1, Y - 1).BuildType = getRoadType(X + 1, Y - 1, "power"): PaintGround
    If BoardData(X, Y - 1).Build = 9 Then BoardData(X, Y - 1).BuildType = getRoadType(X, Y - 1, "power"): PaintGround
    
End Sub

Public Sub Build1x1(mX, mY, Build, btype, PowerOrNot)
    If Not BoardData(mX, mY).Build = 0 Then Exit Sub
    If BoardData(mX, mY).Ter = 0 Then Exit Sub
    
    Select Case Build
    Case 1: If Not FixMoney(infSmallRes.Price) Then Exit Sub
    Case 2:
    Case 3:
    Case 4: If Not FixMoney(infSmallPark.Price) Then Exit Sub
    End Select
    
    With BoardData(mX, mY)
     .Build = Build
     .BuildType = btype
     .Power = PowerOrNot
     .Size = 1
    End With
    
    If BoardData(mX - 1, mY).Build = 9 Then BoardData(mX - 1, mY).BuildType = getRoadType(mX - 1, mY, "power")
    If BoardData(mX, mY + 1).Build = 9 Then BoardData(mX, mY + 1).BuildType = getRoadType(mX, mY + 1, "power")
    If BoardData(mX + 1, mY).Build = 9 Then BoardData(mX + 1, mY).BuildType = getRoadType(mX + 1, mY, "power")
    If BoardData(mX, mY - 1).Build = 9 Then BoardData(mX, mY - 1).BuildType = getRoadType(mX, mY - 1, "power")
End Sub
Public Sub Build2x2(mX, mY, Build, btype, PowerOrNot)
    'Check array
    If Not BoardData(mX, mY).Build = 0 Then Exit Sub
    If BoardData(mX, mY).Ter = 0 Then Exit Sub

    If Not BoardData(mX + 1, mY).Build = 0 Then Exit Sub
    If BoardData(mX + 1, mY).Ter = 0 Then Exit Sub
    
    If Not BoardData(mX + 1, mY + 1).Build = 0 Then Exit Sub
    If BoardData(mX + 1, mY + 1).Ter = 0 Then Exit Sub
    
    If Not BoardData(mX, mY + 1).Build = 0 Then Exit Sub
    If BoardData(mX, mY + 1).Ter = 0 Then Exit Sub
    'check array end
    
    Select Case Build
    Case 1: If Not FixMoney(infMedRes.Price) Then Exit Sub
    Case 2:
    Case 3:
    Case 4: If Not FixMoney(infBigPark.Price) Then Exit Sub
    Case 5: If Not FixMoney(infPowerPlant.Price) Then Exit Sub
    End Select
    
    With BoardData(mX, mY)
     .Build = Build
     .BuildType = btype
     .Power = PowerOrNot
     .Size = 2
     .Child(1).X = mX + 1
     .Child(1).Y = mY
     .Child(2).X = mX + 1
     .Child(2).Y = mY + 1
     .Child(3).X = mX
     .Child(3).Y = mY + 1
    End With
    With BoardData(mX + 1, mY)
     .Build = Build
     .BuildType = 100
     .Power = PowerOrNot
     .Size = 2
     .mParent.X = mX
     .mParent.Y = mY
    End With
    With BoardData(mX + 1, mY + 1)
     .Build = Build
     .BuildType = 100
     .Power = PowerOrNot
     .Size = 2
     .mParent.X = mX
     .mParent.Y = mY
    End With
    With BoardData(mX, mY + 1)
     .Build = Build
     .BuildType = 100
     .Power = PowerOrNot
     .Size = 2
     .mParent.X = mX
     .mParent.Y = mY
    End With
    
        If BoardData(mX - 1, mY).Build = 9 Then BoardData(mX - 1, mY).BuildType = getRoadType(mX - 1, mY, "power")
        If BoardData(mX, mY + 1).Build = 9 Then BoardData(mX, mY + 1).BuildType = getRoadType(mX, mY + 1, "power")
        If BoardData(mX + 1, mY).Build = 9 Then BoardData(mX + 1, mY).BuildType = getRoadType(mX + 1, mY, "power")
        If BoardData(mX, mY - 1).Build = 9 Then BoardData(mX, mY - 1).BuildType = getRoadType(mX, mY - 1, "power")
        For a = 1 To 3
            Dim lX, lY As Integer
            lX = BoardData(mX, mY).Child(a).X: lY = BoardData(mX, mY).Child(a).Y
            If BoardData(lX - 1, lY).Build = 9 Then BoardData(lX - 1, lY).BuildType = getRoadType(lX - 1, lY, "power")
            If BoardData(lX, lY + 1).Build = 9 Then BoardData(lX, lY + 1).BuildType = getRoadType(lX, lY + 1, "power")
            If BoardData(lX + 1, lY).Build = 9 Then BoardData(lX + 1, lY).BuildType = getRoadType(lX + 1, lY, "power")
            If BoardData(lX, lY - 1).Build = 9 Then BoardData(lX, lY - 1).BuildType = getRoadType(lX, lY - 1, "power")
        Next a

End Sub

Public Sub BuildRoad(X, Y)
    
    If BoardData(X, Y).Build = 9 Then
        Select Case BoardData(X, Y).BuildType
        Case 1, 2
        Case Else: Exit Sub
        End Select
    ElseIf Not BoardData(X, Y).Build = 0 Then: Exit Sub
    End If
    
    If BoardData(X, Y).Ter = 0 Then Exit Sub
    If Not FixMoney(infRoads.Price) Then Exit Sub
    
    If BoardData(X, Y).Build = 9 Then
        BoardData(X, Y).Power = 1
        If BoardData(X, Y).BuildType = 1 Then BoardData(X, Y).BuildType = 18
        If BoardData(X, Y).BuildType = 2 Then BoardData(X, Y).BuildType = 17
    Else
    BoardData(X, Y).Power = ""
    BoardData(X, Y).BuildType = getRoadType(X, Y, "road")
    End If
    BoardData(X, Y).Build = 10
    BoardData(X, Y).Size = 1

    
        If BoardData(X - 1, Y).Build = 10 Then BoardData(X - 1, Y).BuildType = getRoadType(X - 1, Y, "road")
        If BoardData(X, Y + 1).Build = 10 Then BoardData(X, Y + 1).BuildType = getRoadType(X, Y + 1, "road")
        If BoardData(X + 1, Y).Build = 10 Then BoardData(X + 1, Y).BuildType = getRoadType(X + 1, Y, "road")
        If BoardData(X, Y - 1).Build = 10 Then BoardData(X, Y - 1).BuildType = getRoadType(X, Y - 1, "road")
End Sub
Public Sub BuildLine(X, Y)
    
    If BoardData(X, Y).Build = 10 Then
        Select Case BoardData(X, Y).BuildType
        Case 1, 2
        Case Else: Exit Sub
        End Select
    ElseIf Not BoardData(X, Y).Build = 0 Then: Exit Sub
    End If

    If BoardData(X, Y).Ter = 0 Then Exit Sub
    If Not FixMoney(infPowerlines.Price) Then Exit Sub
    
    If BoardData(X, Y).Build = 10 Then
        BoardData(X, Y).Build = 10
        BoardData(X, Y).BuildType = BoardData(X, Y).BuildType + 16
    Else
        BoardData(X, Y).Build = 9
        BoardData(X, Y).BuildType = getRoadType(X, Y, "power")
    End If
    BoardData(X, Y).Size = 1
    BoardData(X, Y).Power = 1
    
        If BoardData(X - 1, Y).Build = 9 Then BoardData(X - 1, Y).BuildType = getRoadType(X - 1, Y, "power")
        If BoardData(X, Y + 1).Build = 9 Then BoardData(X, Y + 1).BuildType = getRoadType(X, Y + 1, "power")
        If BoardData(X + 1, Y).Build = 9 Then BoardData(X + 1, Y).BuildType = getRoadType(X + 1, Y, "power")
        If BoardData(X, Y - 1).Build = 9 Then BoardData(X, Y - 1).BuildType = getRoadType(X, Y - 1, "power")
    
End Sub
Public Sub BuildBridge(sX, sY, gX, gY)
Dim DeltaX As Integer, DeltaY As Integer
Dim StepSize
Dim X As Integer, Y As Integer
    sX = sX + WStartX: gX = gX + WStartX
    sY = sY + WStartY: gY = gY + WStartY
    
    DeltaX = sX - gX
    DeltaY = sY - gY
    
    If Not DeltaX = 0 Xor DeltaY = 0 Then Exit Sub 'Skjekk at vi bare strekker rett
    If BoardData(sX, sY).Ter = 0 Or BoardData(gX, gY).Ter = 0 Then Exit Sub
    
    If Not FixMoney((Abs(DeltaX) + Abs(DeltaY) + 1) * infBridge.Price) Then Exit Sub
    
    
    If DeltaX = 0 Then 'Vertikal
        If sY > gY Then StepSize = -1 Else StepSize = 1
        
        For Y = sY To gY Step StepSize
            If Not BoardData(sX, Y).Build = 0 Then Exit Sub
        Next Y
        
        For Y = sY To gY Step StepSize ' Sett proerties for hele brua
            BoardData(sX, Y).Build = 10
            BoardData(sX, Y).BuildType = 24
            BoardData(sX, Y).Size = 1
            BoardData(sX, Y).Power = 1
            BoardData(sX, Y).mParent.X = sX
            BoardData(sX, Y).mParent.Y = sY
        Next Y
        
        If sY > gY Then
            BoardData(sX, sY).BuildType = 19
            BoardData(gX, gY).BuildType = 20
        Else
            BoardData(sX, sY).BuildType = 20
            BoardData(gX, gY).BuildType = 19
        End If
        If BoardData(sX, sY - 1).Build = 10 Then BoardData(sX, sY - 1).BuildType = getRoadType(sX, sY - 1, "road")
        If BoardData(sX, sY + 1).Build = 10 Then BoardData(sX, sY + 1).BuildType = getRoadType(sX, sY + 1, "road")
        If BoardData(gX, gY - 1).Build = 10 Then BoardData(gX, gY - 1).BuildType = getRoadType(gX, gY - 1, "road")
        If BoardData(gX, gY + 1).Build = 10 Then BoardData(gX, gY + 1).BuildType = getRoadType(gX, gY + 1, "road")
    Else 'Horisontal
        If sX > gX Then StepSize = -1 Else StepSize = 1
        
        For X = sX To gX Step StepSize
            If Not BoardData(X, sY).Build = 0 Then Exit Sub
        Next X
        
        For X = sX To gX Step StepSize ' Sett proerties for hele brua
            BoardData(X, sY).Build = 10
            BoardData(X, sY).BuildType = 23
            BoardData(X, sY).Size = 1
            BoardData(X, sY).Power = 1
            BoardData(X, sY).mParent.X = sX
            BoardData(X, sY).mParent.Y = sY
        Next X
        
        If sX > gX Then
            BoardData(sX, sY).BuildType = 22
            BoardData(gX, gY).BuildType = 21
        Else
            BoardData(sX, sY).BuildType = 21
            BoardData(gX, gY).BuildType = 22
        End If
        If BoardData(sX - 1, sY).Build = 10 Then BoardData(sX - 1, sY).BuildType = getRoadType(sX - 1, sY, "road")
        If BoardData(sX + 1, sY).Build = 10 Then BoardData(sX + 1, sY).BuildType = getRoadType(sX + 1, sY, "road")
        If BoardData(gX - 1, gY).Build = 10 Then BoardData(gX - 1, gY).BuildType = getRoadType(gX - 1, gY, "road")
        If BoardData(gX + 1, gY).Build = 10 Then BoardData(gX + 1, gY).BuildType = getRoadType(gX + 1, gY, "road")
        
    End If
    BoardData(sX, sY).Child(1).X = gX
    BoardData(sX, sY).Child(1).Y = gY
    BoardData(sX, sY).mParent.X = 0: BoardData(sX, sY).mParent.Y = 0
    
        If BoardData(sX - 1, sY).Build = 9 Then BoardData(sX - 1, sY).BuildType = getRoadType(sX - 1, sY, "power")
        If BoardData(sX, sY + 1).Build = 9 Then BoardData(sX, sY + 1).BuildType = getRoadType(sX, sY + 1, "power")
        If BoardData(sX + 1, sY).Build = 9 Then BoardData(sX + 1, sY).BuildType = getRoadType(sX + 1, sY, "power")
        If BoardData(sX, sY - 1).Build = 9 Then BoardData(sX, sY - 1).BuildType = getRoadType(sX, sY - 1, "power")
        
        If BoardData(gX - 1, gY).Build = 9 Then BoardData(gX - 1, gY).BuildType = getRoadType(gX - 1, gY, "power")
        If BoardData(gX, gY + 1).Build = 9 Then BoardData(gX, gY + 1).BuildType = getRoadType(gX, gY + 1, "power")
        If BoardData(gX + 1, gY).Build = 9 Then BoardData(gX + 1, gY).BuildType = getRoadType(gX + 1, gY, "power")
        If BoardData(gX, gY - 1).Build = 9 Then BoardData(gX, gY - 1).BuildType = getRoadType(gX, gY - 1, "power")
End Sub

Function getRoadType(X, Y, mType)
Dim RoadAr(1 To 4)
Dim TypeAr(1 To 4) As Integer
Dim CountRoad As Byte
Dim Dire As Directions
Select Case mType
    Case "road"
        'ROADS
        If BoardData(X, Y).BuildType = 17 Or BoardData(X, Y).BuildType = 18 Then getRoadType = BoardData(X, Y).BuildType: Exit Function
        
        Select Case BoardData(X, Y).BuildType 'hvis vi ser på en BRIDGE
        Case 19 To 24
            getRoadType = BoardData(X, Y).BuildType
            Exit Function
        End Select
        
        RoadAr(1) = BoardData(X - 1, Y).Build
        RoadAr(2) = BoardData(X, Y + 1).Build
        RoadAr(3) = BoardData(X + 1, Y).Build
        RoadAr(4) = BoardData(X, Y - 1).Build
        
        TypeAr(1) = BoardData(X - 1, Y).BuildType
        TypeAr(2) = BoardData(X, Y + 1).BuildType
        TypeAr(3) = BoardData(X + 1, Y).BuildType
        TypeAr(4) = BoardData(X, Y - 1).BuildType
        
        If RoadAr(1) = 10 Then Dire.v = True
        If RoadAr(2) = 10 Then Dire.n = True
        If RoadAr(3) = 10 Then Dire.h = True
        If RoadAr(4) = 10 Then Dire.o = True
        
        For a = 1 To 4 ' slå off hvis bridge
            Select Case TypeAr(a)
            Case 19 To 24
                If a = 1 Then Dire.v = False
                If a = 2 Then Dire.n = False
                If a = 3 Then Dire.h = False
                If a = 4 Then Dire.o = False
            End Select
        Next a
        ' Se om vi har en ONRAMP til BRIDGE her...
        If RoadAr(1) = 10 And TypeAr(1) = 22 Then Dire.v = True
        If RoadAr(2) = 10 And TypeAr(2) = 20 Then Dire.n = True
        If RoadAr(3) = 10 And TypeAr(3) = 21 Then Dire.h = True
        If RoadAr(4) = 10 And TypeAr(4) = 19 Then Dire.o = True
        
        With Dire
            If .v = False And .n = False And .h = False And .o = False Then getRoadType = 12 ' Single
            
            If .v = True And .n = False And .h = False And .o = False Then getRoadType = 13 ' Ender
            If .v = False And .n = True And .h = False And .o = False Then getRoadType = 16
            If .v = False And .n = False And .h = True And .o = False Then getRoadType = 14
            If .v = False And .n = False And .h = False And .o = True Then getRoadType = 15
            
            If .v = False And .n = True And .h = True And .o = False Then getRoadType = 3 'svinger
            If .v = False And .n = False And .h = True And .o = True Then getRoadType = 4
            If .v = True And .n = False And .h = False And .o = True Then getRoadType = 5
            If .v = True And .n = True And .h = False And .o = False Then getRoadType = 6
            
            If .v = True And .n = False And .h = True And .o = False Then getRoadType = 1 'rette
            If .v = False And .n = True And .h = False And .o = True Then getRoadType = 2
            
            If .v = True And .n = True And .h = False And .o = True Then getRoadType = 7 'T kryss
            If .v = True And .n = True And .h = True And .o = False Then getRoadType = 8
            If .v = False And .n = True And .h = True And .o = True Then getRoadType = 9
            If .v = True And .n = False And .h = True And .o = True Then getRoadType = 10
            
            If .v = True And .n = True And .h = True And .o = True Then getRoadType = 11 ' + kryss
        End With
    
    Case "power"
        'POWER LINES
        
        RoadAr(1) = BoardData(X - 1, Y).Power
        RoadAr(2) = BoardData(X, Y + 1).Power
        RoadAr(3) = BoardData(X + 1, Y).Power
        RoadAr(4) = BoardData(X, Y - 1).Power
        
        For a = 1 To 4
            If Not RoadAr(a) = "" Then CountRoad = CountRoad + 1
        Next a
        
        Select Case CountRoad
        Case 0
            getRoadType = 12
        Case 1
            If Not RoadAr(1) = "" Then getRoadType = 13
            If Not RoadAr(2) = "" Then getRoadType = 16
            If Not RoadAr(3) = "" Then getRoadType = 14
            If Not RoadAr(4) = "" Then getRoadType = 15
        Case 2
            If Not RoadAr(2) = "" And Not RoadAr(3) = "" Then getRoadType = 3  'svinger
            If Not RoadAr(3) = "" And Not RoadAr(4) = "" Then getRoadType = 4
            If Not RoadAr(1) = "" And Not RoadAr(4) = "" Then getRoadType = 5
            If Not RoadAr(1) = "" And Not RoadAr(2) = "" Then getRoadType = 6
            If Not RoadAr(2) = "" And Not RoadAr(4) = "" Then getRoadType = 2  'rette
            If Not RoadAr(1) = "" And Not RoadAr(3) = "" Then getRoadType = 1
        Case 3
            If Not RoadAr(1) = "" And Not RoadAr(2) = "" And Not RoadAr(4) = "" Then getRoadType = 7
            If Not RoadAr(1) = "" And Not RoadAr(2) = "" And Not RoadAr(3) = "" Then getRoadType = 8
            If Not RoadAr(2) = "" And Not RoadAr(3) = "" And Not RoadAr(4) = "" Then getRoadType = 9
            If Not RoadAr(3) = "" And Not RoadAr(4) = "" And Not RoadAr(1) = "" Then getRoadType = 10
        Case 4
            getRoadType = 11
        End Select
          
    End Select
    
End Function
Public Function FixMoney(Price) As Boolean
    If cityinfo.Money - Price < 0 Then
        FixMoney = False
        ShowMessage "Our City can not afford this", vbRed
    Else
        FixMoney = True
        cityinfo.Money = cityinfo.Money - Price
    End If
End Function
