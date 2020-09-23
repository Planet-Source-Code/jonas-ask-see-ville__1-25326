Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Type COOR
 X As Integer
 Y As Integer
End Type

Public Type TileData
 Ter As Integer
 TerType As Integer
 Shoretype As Integer
 Build As Integer
 BuildType As Integer
 Size As Byte
 LandVal As Integer
 Child(1 To 3) As COOR
 mParent As COOR
 Power As String
End Type

Public DragHold As Boolean
Public DragDown As COOR
Public DragUp As COOR

Public Dirty As Boolean
Public MainPause As Boolean

Public CurrentMSG As String
Public MSGTimeLeft As Integer

Public Const GameTitle As String = "City Game 2001"

Public Board As PictureBox
Public BufferM As PictureBox
Public BufferS As PictureBox
Public BufferG As PictureBox
Public BufferOLm As PictureBox
Public BufferOL As PictureBox
Public MiniMap As PictureBox
Public BufferMap As PictureBox

Public Bredde As Integer
Public Hoyde As Integer
Public WBredde As Integer
Public WHoyde As Integer
Public WStartX As Integer
Public WStartY As Integer

Public Const Size As Integer = 20

Public SelItem As String

Public BoardData() As TileData

Public Function RndTall(Min, Max)
    Randomize
    RndTall = Int((Rnd * Max) + Min)
End Function

Public Function GetXY(XY)
    GetXY = Int(XY / Size)
End Function

Public Sub DoPower()
Dim X, Y As Integer
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            If BoardData(X, Y).Power <> "" And BoardData(X, Y).Build <> 5 Then BoardData(X, Y).Power = "0"
        Next X
    Next Y
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            If BoardData(X, Y).Power = "1" Then
                Putpower X, Y
            End If
            
            
        Next X
    Next Y
End Sub

Sub Putpower(X, Y)
    If BoardData(X - 1, Y).Power = "0" Then
        BoardData(X - 1, Y).Power = 1
        Putpower X - 1, Y
    End If
    If BoardData(X + 1, Y).Power = "0" Then
        BoardData(X + 1, Y).Power = 1
        Putpower X + 1, Y
    End If
    If BoardData(X, Y - 1).Power = "0" Then
        BoardData(X, Y - 1).Power = 1
        Putpower X, Y - 1
    End If
    If BoardData(X, Y + 1).Power = "0" Then
        BoardData(X, Y + 1).Power = 1
        Putpower X, Y + 1
    End If
End Sub

Public Sub Loadmap(Path)
Dim Lendge As Integer
Dim TempIn As String
Dim FreeNum
Dim a
    'On Error GoTo Error1
    
    FileNum = FreeFile
    Open Path For Random As FileNum Len = 10
    
    Get FileNum, 1, Bredde
    Get FileNum, 2, Hoyde
    
    ReDim BoardData(1 To Bredde, 1 To Hoyde)
    
    a = 2
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            a = a + 1
            Get FileNum, a, TempIn
            
            BoardData(X, Y).Ter = Mid(TempIn, 1, 1)
            BoardData(X, Y).TerType = Mid(TempIn, 2, 1)
            Lengde = Mid(TempIn, 3, 1)
            BoardData(X, Y).BuildType = Mid(TempIn, 4, Lengde)
            If BoardData(X, Y).BuildType <> 0 Then: BoardData(X, Y).Size = 1
        Next X
    Next Y
    Close FileNum
    
    Exit Sub
    
Error1:
    MsgBox Err.Description, vbCritical, "Error " & Err.Number
End Sub


Public Sub PaintlMap()
    PaintMap 1, MiniMap
End Sub

Public Sub Determin(X, Y, P)
    If X <= 0 Then Exit Sub
    If Y <= 0 Then Exit Sub
    If X >= Bredde Then Exit Sub
    If Y >= Hoyde Then Exit Sub
    
    If MainPause = True And Not SelItem = "enquire" Then Exit Sub
    
    Select Case SelItem
    Case "road"
        BuildRoad X, Y
    Case "demo"
        Demolish X, Y
    Case "res1"
        BuildNewResSml X, Y, RndTall(1, 10) + 0
    Case "res2"
        BuildNewResMed X, Y, RndTall(1, 3) + 10 '10 er så mange små bygg det finnes
    Case "com1"
        BuildNewComSml X, Y, RndTall(1, 2) + 0
    Case "plant"
        BuildNewPlant X, Y, 1
    Case "lines"
        BuildLine X, Y
    Case "trees"
        BuildTree X, Y
    Case "park1"
        BuildPark1 X, Y, RndTall(1, 5)
    Case "park2"
        BuildPark2 X, Y, RndTall(1, 2) + 5
    Case "enquire"
        Enquire X, Y
    End Select
        
    Dirty = True
    
    If P = 1 Then
        PaintGround
        PaintlMap
    End If
End Sub

