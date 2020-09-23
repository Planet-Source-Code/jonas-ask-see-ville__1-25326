VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Game"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Difficulty"
      Height          =   1215
      Left            =   2640
      TabIndex        =   14
      Top             =   1980
      Width           =   2775
      Begin VB.OptionButton Opt3 
         Caption         =   "Hard"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Medium"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Easy"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "City Info"
      Height          =   1755
      Left            =   2640
      TabIndex        =   11
      Top             =   180
      Width           =   2775
      Begin VB.TextBox txtCity 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   180
         MaxLength       =   40
         TabIndex        =   0
         Top             =   540
         Width           =   2415
      End
      Begin VB.TextBox txtMayor 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   180
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Mayor name"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "City name"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map"
      Height          =   3675
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   2355
      Begin VB.FileListBox File1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2820
         Left            =   120
         Pattern         =   "*.map"
         TabIndex        =   5
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lblHoyde 
         Caption         =   "Height:"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblBredde 
         Caption         =   "Width:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3540
      TabIndex        =   6
      Top             =   3360
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4500
      TabIndex        =   8
      Top             =   3360
      Width           =   915
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
    MainPause = False
    Form1.lblPause.FontBold = False
End Sub


Private Sub cmdStart_Click()
Dim Svar As String
Dim D As Single
    If txtCity.Text = "" Then Exit Sub
    If txtMayor.Text = "" Then Exit Sub
    If File1.FileName = "" Then Exit Sub
    
    If Dirty Then
        Svar = MsgBox("The game is not saved. Continue?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Exit Sub
    End If
    
    If opt1.Value Then D = 1
    If Opt2.Value Then D = 2
    If Opt3.Value Then D = 3
    
    Me.Hide
    MainPause = False
    Form1.lblPause.FontBold = False
    NewGame File1.Path & "\" & File1.FileName, txtCity.Text, txtMayor.Text, D
    
    
End Sub

Private Sub File1_Click()
Dim FreeNum As Integer
Dim temp1 As Integer
Dim temp2 As Integer
    If File1.FileName = "" Then Exit Sub
    
    FreeNum = FreeFile
    Open File1.Path & "\" & File1.FileName For Random As FreeNum Len = 10
    Get FreeNum, 1, temp1
    Get FreeNum, 2, temp2
    
    lblHoyde.Caption = "Heigth: " & temp2
    lblBredde.Caption = "Width: " & temp1
    
    Close FreeNum
End Sub

Private Sub Form_Activate()
    MainPause = True
    Form1.lblPause.FontBold = True
    File1.Path = App.Path & "\maps\"
    File1.Refresh
    lblHoyde.Caption = "Heigth: "
    lblBredde.Caption = "Width: "
    txtCity.Text = ""
    txtMayor.Text = ""
    opt1.Value = True
    
    frmSave.Hide
    frmInfo.Hide
    frmLoad.Hide
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MainPause = False
    Form1.lblPause.Enabled = True
End Sub
