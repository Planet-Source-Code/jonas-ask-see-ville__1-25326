Attribute VB_Name = "GUI"
Public Sub PaintMapSmall(PS, picmap As PictureBox)
        
    For Y = WStartY To WStartY + WHoyde + 5
        For X = WStartX To WStartX + WBredde + 5
            DoEvents
            Select Case BoardData(X, Y).Ter
            Case 0: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlue, BF
            Case 1 To 4: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(6, 124, 12), BF
            Case Else: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlack, BF
            End Select
        Next X
    Next Y
    picmap.Cls
    BitBlt picmap.hDC, 0, 0, Bredde * PS, Hoyde * PS, BufferMap.hDC, 0, 0, SRCCOPY
    picmap.Line (WStartX - 2, WStartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picmap.Refresh
End Sub
Public Sub PaintMap(PS, picmap As PictureBox)
    BufferMap.Cls
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            DoEvents
            Select Case BoardData(X, Y).Ter
            Case 0: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlue, BF
            Case 1 To 4: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(6, 124, 12), BF
            Case Else: BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlack, BF
            End Select
            
            'buildings
            Select Case BoardData(X, Y).Build
            Case 1 'residental
                BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbRed, BF
            Case 2 'Commerce
                BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(49, 11, 215), BF
            Case 4
                BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(64, 255, 64), BF
            Case 5 'Power plant
                BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(152, 53, 6), BF
            Case 9 ' Power Lines
                BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(200, 200, 200), BF
            Case 10 'roads
                 BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(100, 100, 100), BF
            Case Else
            End Select
            
            
        Next X
    Next Y

    
    picmap.Cls
    BitBlt picmap.hDC, 0, 0, Bredde * PS, Hoyde * PS, BufferMap.hDC, 0, 0, SRCCOPY
    picmap.Line (WStartX - 2, WStartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picmap.Refresh
End Sub

Public Sub PaintGround()
    Select Case Form1.chkLand.Value
    Case 1
        PaintGroundLValue
    Case 0
        PaintGroundNormal
    End Select
End Sub

Public Sub PaintGroundLValue()
Dim Time1 As Currency
Dim Time2 As Currency
Dim temp As COOR
Dim ValueColor As Long
Dim X, Y As Integer
    Time1 = GetTickCount
    
    BufferOL.Cls
    BufferOLm.Cls
    BufferS.Cls
    BufferM.Cls
    For Y = WStartY To WStartY + WHoyde
        For X = WStartX To WStartX + WBredde
            BitBlt BufferG.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicGround.Item(BoardData(X, Y).TerType).hDC, 0, 0, SRCCOPY
            
            If BoardData(X, Y).Build = 0 Then
                'Få med trærna
                If Not BoardData(X, Y).BuildType = 0 Then
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmTree.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicTree.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
                End If
            Else
                Select Case BoardData(X, Y).LandVal
                Case -100 To -30: ValueColor = RGB(255, 255, 255)
                Case -30 To 0: ValueColor = RGB(211, 211, 255)
                Case 0 To 30: ValueColor = RGB(200, 200, 255)
                Case 30 To 50: ValueColor = RGB(153, 153, 255)
                Case 50 To 70: ValueColor = RGB(115, 117, 255)
                Case 70 To 130: ValueColor = RGB(91, 91, 255)
                Case 130 To 200: ValueColor = RGB(25, 25, 255)
                Case Is > 200: ValueColor = RGB(0, 0, 230)
                End Select
            
            
                BufferOLm.Line ((X - WStartX) * Size, ((Y - WStartY) * Size))-Step(Size, Size), vbBlack, BF
                BufferOL.Line ((X - WStartX) * Size, ((Y - WStartY) * Size))-Step(Size, Size), ValueColor, BF
            End If
        Next X
    Next Y
    
    Board.Cls
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferG.hDC, 0, 0, SRCCOPY
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferM.hDC, 0, 0, SRCAND
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferS.hDC, 0, 0, SRCPAINT
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferOLm.hDC, 0, 0, SRCAND
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferOL.hDC, 0, 0, SRCPAINT
    
    Time2 = GetTickCount
    Form1.lblRate = "TPF: " & (Time2 - Time1)

End Sub

Public Sub PaintGroundNormal()
Dim Time1 As Currency
Dim Time2 As Currency
Dim temp As COOR
Dim X, Y As Integer
    Time1 = GetTickCount
    
    BufferOL.Cls
    BufferOLm.Cls
    BufferS.Cls
    BufferM.Cls
    For Y = WStartY To WStartY + WHoyde
        For X = WStartX To WStartX + WBredde
            a = a + 1
            BitBlt BufferG.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicGround.Item(BoardData(X, Y).TerType).hDC, 0, 0, SRCCOPY
            
            'Det her er for å få med større bygg som bare "stikker inn"
            If X - WStartX = 0 Or Y - WStartY = 0 Then
                If Not BoardData(X, Y).mParent.X = 0 Then 'SE OM VI HAR EN CHILD
                    If BoardData(X, Y).Build = 10 Then GoTo AsNormal 'THIS WILL NOT HAPPEN TO A BRIDGE
                    temp.X = X
                    temp.Y = Y
                    If BoardData(X, Y).Power = "0" Then 'Se å få med strømmen
                        BitBlt BufferOLm.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmPower.hDC, 0, 0, SRCAND
                        BitBlt BufferOL.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicPower.hDC, 0, 0, SRCPAINT
                    End If
                    X = BoardData(temp.X, temp.Y).mParent.X
                    Y = BoardData(temp.X, temp.Y).mParent.Y
                End If
            End If
AsNormal:
            Select Case BoardData(X, Y).Build
            Case 0 'Trees
                If Not BoardData(X, Y).BuildType = 0 Then
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmTree.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicTree.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
                End If
            Case 1 'Residenal buildings
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.PicmBu.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.PicBu.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 2 'Commercial buildings
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.Picmcom.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.PicCom.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 4 'parks
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.PicmPark.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * BoardData(X, Y).Size, Size * BoardData(X, Y).Size, Pictures.PicPark.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 5 'power plant (2X2)
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * 2, Size * 2, Pictures.PicmPlant.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size * 2, Size * 2, Pictures.PicPlant.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 9 'power lines
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmLines.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicLines.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 10 'road
                BitBlt BufferM.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmRoad.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt BufferS.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicRoad.Item(BoardData(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            Case 11

            End Select
            
            If BoardData(X, Y).Power = "0" Then
                BitBlt BufferOLm.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicmPower.hDC, 0, 0, SRCAND
                BitBlt BufferOL.hDC, ((X - WStartX) * Size), ((Y - WStartY) * Size), Size, Size, Pictures.PicPower.hDC, 0, 0, SRCPAINT
            End If
            
            
            
            'Det her er for å få med større bygg som bare "stikker inn"
            If Not temp.X = 0 Or Not temp.Y = 0 Then
                If Not BoardData(X, Y).Child(1).X = 0 Then 'SE OM VI HAR EN PARENT
                    X = temp.X
                    Y = temp.Y
                    temp.X = 0
                    temp.Y = 0
                End If
            End If
        Next X
    Next Y
    ComposeMap
    
    Time2 = GetTickCount
    Form1.lblRate = "TPF: " & (Time2 - Time1)

End Sub

Public Sub ComposeMap()
    Board.Cls
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferG.hDC, 0, 0, SRCCOPY
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferM.hDC, 0, 0, SRCAND
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferS.hDC, 0, 0, SRCPAINT
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferOLm.hDC, 0, 0, SRCAND
    BitBlt Board.hDC, 0, 0, Bredde * Size, Hoyde * Size, BufferOL.hDC, 0, 0, SRCPAINT
End Sub
Public Sub UpdateData()
    Form1.Caption = GameTitle & " - " & cityinfo.CityName & " " & cityinfo.Inhabitants & " inhab. - $" & cityinfo.Money & "  Week: " & cityinfo.Week & "  Month: " & cityinfo.Month & "  Year: " & cityinfo.Year
End Sub

Public Sub ShowMessage(Text, Color)
    CurrentMSG = Text
    MSGTimeLeft = 10
    Form1.lblInfo.ForeColor = Color
    Form1.lblInfo.Caption = Text
    Form1.tmrMSG.Enabled = True
End Sub

Public Sub SetSeason(Num)
    With Pictures
    Select Case Num
    Case 1
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicSpring.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 2
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicSummer.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 3
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicAutumn.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 4
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicWinter.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    End Select
    End With
    cityinfo.Season = Num
End Sub
