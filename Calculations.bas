Attribute VB_Name = "Calculations"
Const effRes1 As Integer = -5
Const effRes2 As Integer = -10
Const effLand As Integer = 3
Const effTree1 As Integer = 5
Const effTree2 As Integer = 7
Const effTree3 As Integer = 9
Const effPark1 As Integer = 30
Const effPark2 As Integer = 40
Const effPlant As Integer = -40
Const effLines As Integer = -5
Const effRoad As Integer = -1
Public Sub DoLandvalues()
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            If BoardData(X, Y).Build = 0 Then GoTo NextX
            BoardData(X, Y).LandVal = 0
            BoardData(X, Y).LandVal = BoardData(X, Y).LandVal + GetLandvalue(X, Y)
            
NextX:
        Next X
    Next Y
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            If BoardData(X, Y).Build = 0 Then
                BoardData(X, Y).LandVal = 0
            End If
            If BoardData(X, Y).LandVal > 200 Then BoardData(X, Y).LandVal = 200
            If BoardData(X, Y).LandVal < -100 Then BoardData(X, Y).LandVal = -100
            ApplyOnArray X, Y
        Next X
    Next Y
End Sub
Sub ApplyOnArray(mX, mY)
    Select Case BoardData(mX, mY).Build
    Case 4 'Parks
        Select Case BoardData(mX, mY).BuildType
        Case 1 To 5: DoArray 1, effPark1, mX, mY
        Case 6 To 10: DoArray 2, effPark1, mX, mY
        End Select
    Case 5 'Powerplant
        Select Case BoardData(mX, mY).BuildType
        Case Is <> 100: DoArray 2, effPlant, mX, mY
        End Select
    End Select
End Sub
Sub DoArray(ASize, AAmount, X, Y)
    Select Case ASize
    Case 1
        For a = -1 To 1
            For b = -1 To 1
                BoardData(X + b, Y + a).LandVal = BoardData(X + b, Y + a).LandVal + AAmount
            Next b
        Next a
    Case 2
        For a = -2 To 3 'akompanser for at vi starter på rute 1x1
            For b = -2 To 3
                If Not BoardData(X + b, Y + a).Build = 0 Then
                    BoardData(X + b, Y + a).LandVal = BoardData(X + b, Y + a).LandVal + AAmount
                End If
            Next b
        Next a
    End Select
        
End Sub
Public Function GetLandvalue(mX, mY)
Dim buildA(1 To 8) As Integer
Dim TypeA(1 To 8) As Integer
Dim LandCount As Integer
    
    If BoardData(mX, mY).Power = "0" Then GetLandvalue = 10: Exit Function
    
    buildA(1) = BoardData(mX - 1, mY).Build
    buildA(2) = BoardData(mX - 1, mY + 1).Build
    buildA(3) = BoardData(mX, mY + 1).Build
    buildA(4) = BoardData(mX + 1, mY + 1).Build
    buildA(5) = BoardData(mX + 1, mY).Build
    buildA(6) = BoardData(mX + 1, mY - 1).Build
    buildA(7) = BoardData(mX, mY - 1).Build
    buildA(8) = BoardData(mX - 1, mY - 1).Build
    TypeA(1) = BoardData(mX - 1, mY).BuildType
    TypeA(2) = BoardData(mX - 1, mY + 1).BuildType
    TypeA(3) = BoardData(mX, mY + 1).BuildType
    TypeA(4) = BoardData(mX + 1, mY + 1).BuildType
    TypeA(5) = BoardData(mX + 1, mY).BuildType
    TypeA(6) = BoardData(mX + 1, mY - 1).BuildType
    TypeA(7) = BoardData(mX, mY - 1).BuildType
    TypeA(8) = BoardData(mX - 1, mY - 1).BuildType
    
    LandCount = 50
    
    For a = 1 To 8
        Select Case buildA(a)
        Case 0 'GROUND
            Select Case TypeA(a)
            Case 0: LandCount = LandCount + effLand
            Case 1 To 4: LandCount = LandCount + effTree1
            Case 5 To 8: LandCount = LandCount + effTree2
            Case 9 To 12: LandCount = LandCount + effTree3
            Case 13 To 16: LandCount = LandCount + effTree4
            End Select
        Case 1 'Residental
            Select Case TypeA(a)
            Case 1 To 10: LandCount = LandCount + effRes1
            Case 11 To 20, 100: LandCount = LandCount + effRes2
            End Select
        Case 9 'Poerlines
            LandCount = LandCount + effLines
        Case 10 'Road
            LandCount = LandCount + effRoad
        End Select
    Next a
    
    If LandCount > 200 Then LandCount = 200
    If LandCount < -100 Then LandCount = -100
    GetLandvalue = LandCount
End Function
