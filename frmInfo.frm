VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   420
      Width           =   3795
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   6  'Cross
         Height          =   900
         Left            =   1500
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   5
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label lblValue 
      Caption         =   "Landvalue:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2340
      Width           =   3555
   End
   Begin VB.Label lblMaint 
      Caption         =   "Yearly Mainenance:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3555
   End
   Begin VB.Label lblTer 
      Caption         =   "Terrain:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   3555
   End
   Begin VB.Label lblPower 
      Caption         =   "Power"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3555
   End
   Begin VB.Label lblName 
      Caption         =   "Name: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub

