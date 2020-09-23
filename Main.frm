VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8805
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5355
      Left            =   60
      Picture         =   "Main.frx":030A
      ScaleHeight     =   5355
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   660
      Width           =   1215
      Begin VB.CommandButton cmdCom1 
         Caption         =   "Com1"
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   3360
         Width           =   855
      End
      Begin VB.PictureBox IconResSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   540
         Picture         =   "Main.frx":15790
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   29
         Top             =   1020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconPwrSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   540
         Picture         =   "Main.frx":15C82
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   28
         Top             =   1740
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconResSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   540
         Picture         =   "Main.frx":16174
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   27
         Top             =   660
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   180
         Picture         =   "Main.frx":16666
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   26
         Top             =   660
         Width           =   300
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   180
         Picture         =   "Main.frx":16B58
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   25
         Top             =   1380
         Width           =   300
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   180
         Picture         =   "Main.frx":1704A
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   24
         Top             =   1740
         Width           =   300
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   180
         Picture         =   "Main.frx":1753C
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   23
         Top             =   2100
         Width           =   300
      End
      Begin VB.PictureBox IconPwrSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   540
         Picture         =   "Main.frx":17A2E
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   22
         Top             =   2460
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconPwrSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   540
         Picture         =   "Main.frx":17F20
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   21
         Top             =   2820
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconScenSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   540
         Picture         =   "Main.frx":18412
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   20
         Top             =   1020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconScenSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   540
         Picture         =   "Main.frx":18904
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   19
         Top             =   1380
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconScenSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   540
         Picture         =   "Main.frx":18DF6
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   18
         Top             =   1740
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox IconPwrSub 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   540
         Picture         =   "Main.frx":192E8
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   17
         Top             =   2100
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   180
         Picture         =   "Main.frx":197DA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   16
         Top             =   1020
         Width           =   300
      End
   End
   Begin VB.CheckBox chkLand 
      Caption         =   "LandValues"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   13
      Top             =   2100
      Width           =   1215
   End
   Begin VB.PictureBox BufferOLapm 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7260
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox BufferOLap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7260
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox BufferGround 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox BufferSprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Buffermask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Ticker 
      Interval        =   200
      Left            =   8160
      Top             =   2400
   End
   Begin VB.Timer tmrMSG 
      Interval        =   300
      Left            =   4980
      Top             =   60
   End
   Begin VB.PictureBox picMM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   6960
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   5
      Top             =   660
      Width           =   1635
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   1380
      Min             =   1
      TabIndex        =   2
      Top             =   5820
      Value           =   1
      Width           =   5175
   End
   Begin VB.PictureBox MainPic 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   1380
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   1
      Top             =   600
      Width           =   5160
   End
   Begin VB.VScrollBar VScroll 
      Height          =   5175
      Left            =   6600
      Min             =   1
      TabIndex        =   0
      Top             =   600
      Value           =   1
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   6960
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8100
      TabIndex        =   12
      Top             =   0
      Width           =   705
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   3945
   End
   Begin VB.Label lblCOOR 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lblRate 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6240
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mLoad 
         Caption         =   "Load map"
      End
      Begin VB.Menu mBreak 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mSaveGame 
         Caption         =   "Save Game"
      End
      Begin VB.Menu mLoadgame 
         Caption         =   "Load Game"
      End
   End
   Begin VB.Menu MDebug 
      Caption         =   "Debug"
      Begin VB.Menu mRandhouses 
         Caption         =   "Random Houses"
      End
      Begin VB.Menu mSprites 
         Caption         =   "Show Sprites"
      End
      Begin VB.Menu mSetSeason 
         Caption         =   "Set Season"
         Begin VB.Menu mSeason 
            Caption         =   "Spring"
            Index           =   1
         End
         Begin VB.Menu mSeason 
            Caption         =   "Summer"
            Index           =   2
         End
         Begin VB.Menu mSeason 
            Caption         =   "Autumn"
            Index           =   3
         End
         Begin VB.Menu mSeason 
            Caption         =   "Winter"
            Index           =   4
         End
      End
      Begin VB.Menu mEco 
         Caption         =   "Economy"
         Begin VB.Menu mSetMoney 
            Caption         =   "Set money"
            Begin VB.Menu mSetMoneyval 
               Caption         =   "0"
               Index           =   0
            End
            Begin VB.Menu mSetMoneyval 
               Caption         =   "2000"
               Index           =   1
            End
            Begin VB.Menu mSetMoneyval 
               Caption         =   "20000"
               Index           =   2
            End
            Begin VB.Menu mSetMoneyval 
               Caption         =   "200000"
               Index           =   3
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCom1_Click()
    ClearMenu
    SelItem = "com1"
End Sub

Private Sub Form_Load()
    SetData
    CallDefaults
    
    MDebug.Visible = False
    
    Set Board = MainPic
    Set BufferM = Buffermask
    Set BufferS = BufferSprite
    Set BufferG = BufferGround
    Set MiniMap = picMM
    Set BufferMap = Pictures.picMapBuf
    Set BufferOL = BufferOLap
    Set BufferOLm = BufferOLapm
    
    BufferMap.AutoRedraw = True
    
    BufferM.Width = Board.Width
    BufferM.Height = Board.Height
    BufferS.Width = Board.Width
    BufferS.Height = Board.Height
    BufferG.Width = Board.Width
    BufferG.Height = Board.Height
    BufferOL.Width = Board.Width
    BufferOL.Height = Board.Height
    BufferOLm.Width = Board.Width
    BufferOLm.Height = Board.Height
        
    
    NewGame App.Path & "\maps\debug.map", "New Sandefjord", "Jonas Ask", 1
    
    SetSeason cityinfo.Season
    PaintGround
    PaintlMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End 'Få bort ALL driten
End Sub

Private Sub HScroll_Change()
    WStartX = HScroll.Value
    picMM.Cls
    BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
    picMM.Line (WStartX - 2, WStartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMM.Refresh
    PaintGround
End Sub


Private Sub IconPwrSub_Click(Index As Integer)
    MnuPwrClick Index
End Sub

Private Sub IconResSub_Click(Index As Integer)
    MnuResClick Index
End Sub

Private Sub IconScenSub_Click(Index As Integer)
    MnuScenClick Index
End Sub

Private Sub lblPause_Click()
    MainPause = Not MainPause
    If MainPause Then
        lblPause.FontBold = True
    Else
        lblPause.FontBold = False
    End If
End Sub

Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitX As Integer
Dim HitY As Integer
    'HØYRE KNAPP
    If Button = 2 Then
        HitX = GetXY(X)
        HitY = GetXY(Y)
        
        If HitX > WBredde / 2 Then
            HitX = HitX
            
            WStartX = WStartX + HitX - (WBredde / 2)
        Else
            WStartX = WStartX - ((WBredde / 2) - HitX)
        End If
        
        If HitY > WHoyde / 2 Then
            HitY = HitY
            
            WStartY = WStartY + HitY - (WBredde / 2)
        Else
            WStartY = WStartY - ((WHoyde / 2) - HitY)
        End If
        
        If WStartX <= 0 Then WStartX = 1
        If WStartX >= Bredde - WBredde Then WStartX = Bredde - WBredde
        If WStartY <= 0 Then WStartY = 1
        If WStartY >= Hoyde - WHoyde Then WStartY = Hoyde - WHoyde
        
        HScroll.Value = WStartX
        VScroll.Value = WStartY
        
    End If
    
    'VESTRE KNAPP
    If Button = 1 Then
        HitX = GetXY(X)
        HitY = GetXY(Y)
        
        If Shift = 1 Then
        MsgBox GetLandvalue(HitX + WStartX, HitY + WStartY)
        Exit Sub
        End If
        
        DragHold = True
        DragDown.X = HitX
        DragDown.Y = HitY
    End If
End Sub

Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitX As Integer, HitY As Integer, GlX As Integer, GlY As Integer
    
    lblCOOR = "Local: " & GetXY(X) & " - " & GetXY(Y) & vbNewLine _
    & "Global: " & GetXY(X) + WStartX & " - " & GetXY(Y) + WStartY
    GlX = GetXY(X) + WStartX
    GlY = GetXY(Y) + WStartY
    On Error Resume Next
    Label1.Caption = _
    "Power: " & BoardData(GlX, GlY).Power & vbNewLine _
    & "Terrain: " & BoardData(GlX, GlY).Ter & vbNewLine _
    & "TerrainType: " & BoardData(GlX, GlY).TerType & vbNewLine _
    & "Building: " & BoardData(GlX, GlY).Build & vbNewLine _
    & "Buildingtype: " & BoardData(GlX, GlY).BuildType & vbNewLine _
    & "Size: " & BoardData(GlX, GlY).Size & vbNewLine _
    & "Cild: " & BoardData(GlX, GlY).Child(1).X & " - " & BoardData(GlX, GlY).Child(1).Y & vbNewLine _
    & "Parent: " & BoardData(GlX, GlY).mParent.X & " - " & BoardData(GlX, GlY).mParent.Y & vbNewLine
    If DragHold Then

        HitX = GetXY(X)
        HitY = GetXY(Y)
        Board.Cls
        ComposeMap
        
        
        Dim DeltaX, DeltaY As Integer
        DeltaX = DragDown.X - HitX
        DeltaY = DragDown.Y - HitY
        

        If DeltaX > 0 And DeltaY > 0 Then
            Board.Line (HitX * Size, HitY * Size)-((DragDown.X + 1) * Size, (DragDown.Y + 1) * Size), vbGreen, BF
        ElseIf DeltaX < 0 And DeltaY < 0 Then
            Board.Line (DragDown.X * Size, DragDown.Y * Size)-((HitX + 1) * Size, (HitY + 1) * Size), vbGreen, BF
        ElseIf DeltaX < 0 And DeltaY > 0 Then
            Board.Line (DragDown.X * Size, HitY * Size)-((HitX + 1) * Size, (DragDown.Y + 1) * Size), vbGreen, BF
        ElseIf DeltaX > 0 And DeltaY < 0 Then
            Board.Line (HitX * Size, DragDown.Y * Size)-((DragDown.X + 1) * Size, (HitY + 1) * Size), vbGreen, BF
        
        ElseIf DeltaX = 0 And DeltaY < 0 Then
            Board.Line (DragDown.X * Size, DragDown.Y * Size)-((HitX + 1) * Size, (HitY + 1) * Size), vbGreen, BF
        ElseIf DeltaX < 0 And DeltaY = 0 Then
            Board.Line (DragDown.X * Size, DragDown.Y * Size)-((HitX + 1) * Size, (HitY + 1) * Size), vbGreen, BF
        ElseIf DeltaX = 0 And DeltaY > 0 Then
            Board.Line (DragDown.X * Size, HitY * Size)-((HitX + 1) * Size, (DragDown.Y + 1) * Size), vbGreen, BF
        ElseIf DeltaX > 0 And DeltaY = 0 Then
            Board.Line ((DragDown.X + 1) * Size, HitY * Size)-(HitX * Size, (DragDown.Y + 1) * Size), vbGreen, BF
        End If

    End If
End Sub

Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitX, HitY As Integer
    'VESTRE KNAPP
    If Button = 1 Then
        HitX = GetXY(X)
        HitY = GetXY(Y)
        
        DragHold = False
        If Shift = 1 Then GoTo NoDeal
        If DragDown.X = HitX And DragDown.Y = HitY Then
            Determin HitX + WStartX, HitY + WStartY, 1
        Else
        Dim DeltaX, DeltaY As Integer
            DragUp.X = HitX
            DragUp.Y = HitY
            DeltaX = DragDown.X - DragUp.X
            DeltaY = DragDown.Y - DragUp.Y
            
            Select Case SelItem
            Case "bridge" 'BRIDGE ROAD
                BuildBridge DragDown.X, DragDown.Y, HitX, HitY
            Case Else
                For Y = 0 To Abs(DeltaY)
                    For X = 0 To Abs(DeltaX)
                        If DeltaX = 0 And DeltaY < 0 Then
                            Determin DragDown.X - X + WStartX, DragDown.Y + Y + WStartY, 0
                        ElseIf DeltaX < 0 And DeltaY = 0 Then
                            Determin DragDown.X + X + WStartX, DragDown.Y + Y + WStartY, 0
                        ElseIf DeltaX = 0 And DeltaY > 0 Then
                            Determin DragDown.X + X + WStartX, DragDown.Y - Y + WStartY, 0
                        ElseIf DeltaX > 0 And DeltaY = 0 Then
                            Determin DragDown.X - X + WStartX, DragDown.Y + Y + WStartY, 0
                            
                        ElseIf DeltaX < 0 And DeltaY < 0 Then
                            Determin DragDown.X + X + WStartX, DragDown.Y + Y + WStartY, 0
                        ElseIf DeltaX < 0 And DeltaY > 0 Then
                            Determin DragDown.X + X + WStartX, DragDown.Y - Y + WStartY, 0
                        ElseIf DeltaX > 0 And DeltaY < 0 Then
                            Determin DragDown.X - X + WStartX, DragDown.Y + Y + WStartY, 0
                        ElseIf DeltaX > 0 And DeltaY > 0 Then
                            Determin DragDown.X - X + WStartX, DragDown.Y - Y + WStartY, 0
                        End If
                    Next X
                Next Y
            End Select
NoDeal:
            PaintGround
            PaintlMap
        End If
    End If
End Sub


Private Sub mLoad_Click()
Dim MapName As String

    MapName = InputBox("Type in the map name", "Load Map")
    If MapName = Empty Then Exit Sub
    Loadmap App.Path & "\maps\" & MapName & ".map"
    
End Sub


Private Sub mLoadgame_Click()
    frmLoad.Show , Me
End Sub

Private Sub mNewGame_Click()
    frmNew.Show
End Sub

Private Sub mRandhouses_Click()
    For Y = 0 To WHoyde
        For X = 0 To WBredde
            If BoardData(X + WStartX, Y + WStartY).Ter = 1 Then
                If BoardData(X + WStartX, Y + WStartY).Build = 0 Then
                    BoardData(X + WStartX, Y + WStartY).BuildType = RndTall(1, 3)
                    BoardData(X + WStartX, Y + WStartY).Build = 1
                    BoardData(X + WStartX, Y + WStartY).Power = 1
                End If
            End If
            
        Next X
    Next Y
    PaintGround
    PaintlMap
End Sub

Private Sub mSaveGame_Click()
    frmSave.Show , Me
End Sub

Private Sub mSeason_Click(Index As Integer)
    SetSeason Index
End Sub

Private Sub mSetMoneyval_Click(Index As Integer)
    cityinfo.Money = mSetMoneyval.Item(Index).Caption
    UpdateData
End Sub

Private Sub mSprites_Click()
    Pictures.Show
End Sub

Private Sub PicIcon_Click(Index As Integer)
    ClearMenu
    SelItem = ""
    Select Case Index
    Case 0
        For a = 0 To IconResSub.Count - 1
            IconResSub(a).Visible = True
        Next a
        IconResSub_Click (0)
    Case 1
        For a = 0 To IconScenSub.Count - 1
            IconScenSub(a).Visible = True
        Next a
        IconScenSub_Click (0)
    Case 2
        SelItem = "demo"
    Case 3
        For a = 0 To IconPwrSub.Count - 1
            IconPwrSub(a).Visible = True
        Next a
        IconPwrSub_Click (0)
    Case 4
        SelItem = "enquire"
    
    End Select
    picIcon(Index).BorderStyle = 1
End Sub

Private Sub picMM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMM.Tag = 1
    If Button = 2 And Shift = 1 Then MDebug.Visible = True: ShowMessage "Cheat enabled", vbBlue
    If Button = 1 Then
        picMM.Cls
        BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
        picMM.Line (X - (WBredde / 2) - 2, Y - (WHoyde / 2) - 2)-Step(WBredde, WHoyde), vbWhite, B
        picMM.Refresh
        WStartX = X - (WBredde / 2)
        WStartY = Y - (WHoyde / 2)
        
        If WStartX <= 0 Then WStartX = 1
        If WStartX >= Bredde - WBredde Then WStartX = Bredde - WBredde
        If WStartY <= 0 Then WStartY = 1
        If WStartY >= Hoyde - WHoyde Then WStartY = Hoyde - WHoyde
        
        HScroll.Value = WStartX
        VScroll.Value = WStartY
    End If
End Sub

Private Sub picMM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picMM.Cls
        BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
        picMM.Line (X - (WBredde / 2) - 2, Y - (WHoyde / 2) - 2)-Step(WBredde, WHoyde), vbWhite, B
        picMM.Refresh
        WStartX = X - (WBredde / 2)
        WStartY = Y - (WHoyde / 2)
        
        If WStartX <= 0 Then WStartX = 1
        If WStartX >= Bredde - WBredde Then WStartX = Bredde - WBredde
        If WStartY <= 0 Then WStartY = 1
        If WStartY >= Hoyde - WHoyde Then WStartY = Hoyde - WHoyde
        
        HScroll.Value = WStartX
        VScroll.Value = WStartY
    End If

End Sub




Private Sub tmrMSG_Timer()
    Select Case MSGTimeLeft
    Case 0
        tmrMSG.Enabled = False
        lblInfo.Caption = ""
    Case Else
        MSGTimeLeft = MSGTimeLeft - 1
        If MSGTimeLeft < 7 Then lblInfo.ForeColor = RGB(50, 50, 50)
    End Select
End Sub
Private Sub Ticker_Timer()
    OneTick
End Sub

Private Sub VScroll_Change()
    WStartY = VScroll.Value
    picMM.Cls
    BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
    picMM.Line (WStartX - 2, WStartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMM.Refresh
    PaintGround
End Sub
