VERSION 5.00
Begin VB.Form Pictures 
   Caption         =   "Form2"
   ClientHeight    =   11220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form2"
   ScaleHeight     =   11220
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picmcom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   7500
      Picture         =   "Pictures.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   187
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox PicCom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   7140
      Picture         =   "Pictures.frx":04F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   186
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox Picmcom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   7500
      Picture         =   "Pictures.frx":09E4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   185
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicCom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   7140
      Picture         =   "Pictures.frx":0ED6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   184
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   99
      Left            =   1200
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   183
      Top             =   6540
      Width           =   300
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   99
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   182
      Top             =   6540
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   6
      Left            =   4560
      Picture         =   "Pictures.frx":13C8
      ScaleHeight     =   585
      ScaleWidth      =   600
      TabIndex        =   181
      Top             =   2220
      Width           =   600
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   6
      Left            =   4680
      Picture         =   "Pictures.frx":26CA
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   180
      Top             =   2220
      Width           =   600
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   5
      Left            =   5520
      Picture         =   "Pictures.frx":39CC
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   179
      Top             =   2160
      Width           =   600
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   5
      Left            =   5640
      Picture         =   "Pictures.frx":4CCE
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   178
      Top             =   2160
      Width           =   600
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   5640
      Picture         =   "Pictures.frx":5FD0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   177
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   6000
      Picture         =   "Pictures.frx":64C2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   176
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   5640
      Picture         =   "Pictures.frx":69B4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   175
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   6000
      Picture         =   "Pictures.frx":6EA6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   174
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   5640
      Picture         =   "Pictures.frx":7398
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   173
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   6000
      Picture         =   "Pictures.frx":788A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   172
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   6000
      Picture         =   "Pictures.frx":7D7C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   171
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   5640
      Picture         =   "Pictures.frx":826E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   170
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox PicPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   5640
      Picture         =   "Pictures.frx":8760
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   169
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicmPark 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6000
      Picture         =   "Pictures.frx":8C52
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   168
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   6360
      Picture         =   "Pictures.frx":9144
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   167
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   6720
      Picture         =   "Pictures.frx":9636
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   166
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   6360
      Picture         =   "Pictures.frx":9B28
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   165
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   6720
      Picture         =   "Pictures.frx":A01A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   164
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   6360
      Picture         =   "Pictures.frx":A50C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   163
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   6720
      Picture         =   "Pictures.frx":A9FE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   162
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   6360
      Picture         =   "Pictures.frx":AEF0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   161
      Top             =   5520
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   6720
      Picture         =   "Pictures.frx":B3E2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   160
      Top             =   5520
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   6720
      Picture         =   "Pictures.frx":B8D4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   159
      Top             =   4080
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   6360
      Picture         =   "Pictures.frx":BDC6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   158
      Top             =   4080
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   6720
      Picture         =   "Pictures.frx":C2B8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   157
      Top             =   3720
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   6360
      Picture         =   "Pictures.frx":C7AA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   156
      Top             =   3720
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   6720
      Picture         =   "Pictures.frx":CC9C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   155
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   6360
      Picture         =   "Pictures.frx":D18E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   154
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   6720
      Picture         =   "Pictures.frx":D680
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   153
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   6360
      Picture         =   "Pictures.frx":DB72
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   152
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   6360
      Picture         =   "Pictures.frx":E064
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   151
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   6720
      Picture         =   "Pictures.frx":E556
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   150
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   6360
      Picture         =   "Pictures.frx":EA48
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   149
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   6720
      Picture         =   "Pictures.frx":EF3A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   148
      Top             =   1920
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   6360
      Picture         =   "Pictures.frx":F42C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   147
      Top             =   2280
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   6720
      Picture         =   "Pictures.frx":F91E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   146
      Top             =   2280
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   6360
      Picture         =   "Pictures.frx":FE10
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   145
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   6720
      Picture         =   "Pictures.frx":10302
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   144
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   6720
      Picture         =   "Pictures.frx":107F4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   143
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   6360
      Picture         =   "Pictures.frx":10CE6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   142
      Top             =   1200
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   6720
      Picture         =   "Pictures.frx":111D8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   141
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   6360
      Picture         =   "Pictures.frx":116CA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   140
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   6720
      Picture         =   "Pictures.frx":11BBC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   139
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   6360
      Picture         =   "Pictures.frx":120AE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   138
      Top             =   480
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6720
      Picture         =   "Pictures.frx":125A0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   137
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6360
      Picture         =   "Pictures.frx":12A92
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   136
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   60
      Picture         =   "Pictures.frx":12F84
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   135
      Top             =   7440
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "Pictures.frx":13476
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   134
      Top             =   7080
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   60
      Picture         =   "Pictures.frx":13968
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   133
      Top             =   6720
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   60
      Picture         =   "Pictures.frx":13E5A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   132
      Top             =   6360
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   60
      Picture         =   "Pictures.frx":1434C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   131
      Top             =   5880
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "Pictures.frx":1483E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   130
      Top             =   5520
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   60
      Picture         =   "Pictures.frx":14D30
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   129
      Top             =   5160
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   60
      Picture         =   "Pictures.frx":15222
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   128
      Top             =   4800
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   60
      Picture         =   "Pictures.frx":15714
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   127
      Top             =   4320
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "Pictures.frx":15C06
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   126
      Top             =   3960
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   60
      Picture         =   "Pictures.frx":160F8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   125
      Top             =   3600
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   60
      Picture         =   "Pictures.frx":165EA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   124
      Top             =   3240
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   60
      Picture         =   "Pictures.frx":16ADC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   123
      Top             =   2760
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "Pictures.frx":16FCE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   122
      Top             =   2400
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   60
      Picture         =   "Pictures.frx":174C0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   121
      Top             =   2040
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   60
      Picture         =   "Pictures.frx":179B2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   120
      Top             =   1680
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   23
      Left            =   1920
      Picture         =   "Pictures.frx":17EA4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   119
      Top             =   8340
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   23
      Left            =   1560
      Picture         =   "Pictures.frx":18396
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   118
      Top             =   8340
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   22
      Left            =   1920
      Picture         =   "Pictures.frx":18888
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   117
      Top             =   7980
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   22
      Left            =   1560
      Picture         =   "Pictures.frx":18D7A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   116
      Top             =   7980
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   21
      Left            =   1920
      Picture         =   "Pictures.frx":1926C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   115
      Top             =   7620
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   21
      Left            =   1560
      Picture         =   "Pictures.frx":1975E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   114
      Top             =   7620
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   20
      Left            =   1920
      Picture         =   "Pictures.frx":19C50
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   113
      Top             =   7260
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   20
      Left            =   1560
      Picture         =   "Pictures.frx":1A142
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   112
      Top             =   7260
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   19
      Left            =   1920
      Picture         =   "Pictures.frx":1A634
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   111
      Top             =   6900
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   19
      Left            =   1560
      Picture         =   "Pictures.frx":1AB26
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   110
      Top             =   6900
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   18
      Left            =   1920
      Picture         =   "Pictures.frx":1B018
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   109
      Top             =   6540
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   18
      Left            =   1560
      Picture         =   "Pictures.frx":1B50A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   108
      Top             =   6540
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   12
      Left            =   3600
      Picture         =   "Pictures.frx":1B9FC
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   107
      Top             =   3660
      Width           =   600
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   12
      Left            =   3600
      Picture         =   "Pictures.frx":1CCFE
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   106
      Top             =   3000
      Width           =   600
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   840
      Picture         =   "Pictures.frx":1E000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   105
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   1200
      Picture         =   "Pictures.frx":1E4F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   104
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   840
      Picture         =   "Pictures.frx":1E9E4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   103
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   1200
      Picture         =   "Pictures.frx":1EED6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   102
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   840
      Picture         =   "Pictures.frx":1F3C8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   101
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   1200
      Picture         =   "Pictures.frx":1F8BA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   100
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   840
      Picture         =   "Pictures.frx":1FDAC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   99
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   1200
      Picture         =   "Pictures.frx":2029E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   98
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   840
      Picture         =   "Pictures.frx":20790
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   97
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   1200
      Picture         =   "Pictures.frx":20C82
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   96
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   840
      Picture         =   "Pictures.frx":21174
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   95
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   1200
      Picture         =   "Pictures.frx":21666
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   94
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   840
      Picture         =   "Pictures.frx":21B58
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   93
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   1200
      Picture         =   "Pictures.frx":2204A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   92
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   11
      Left            =   3420
      Picture         =   "Pictures.frx":2253C
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   91
      Top             =   3000
      Width           =   600
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   11
      Left            =   3420
      Picture         =   "Pictures.frx":2383E
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   90
      Top             =   3660
      Width           =   600
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   10
      Left            =   3240
      Picture         =   "Pictures.frx":24B40
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   89
      Top             =   3660
      Width           =   600
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   10
      Left            =   3240
      Picture         =   "Pictures.frx":25E42
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   88
      Top             =   3000
      Width           =   600
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   99
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   87
      Top             =   6180
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   99
      Left            =   1200
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   86
      Top             =   6180
      Width           =   300
   End
   Begin VB.PictureBox picMapBuf 
      AutoRedraw      =   -1  'True
      Height          =   4515
      Left            =   4080
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   85
      Top             =   8280
      Width           =   6135
   End
   Begin VB.PictureBox PicPlant 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   99
      Left            =   3120
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   84
      Top             =   1140
      Width           =   600
   End
   Begin VB.PictureBox PicmPlant 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   99
      Left            =   3780
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   83
      Top             =   1140
      Width           =   600
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   1200
      Picture         =   "Pictures.frx":27144
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   82
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   840
      Picture         =   "Pictures.frx":27636
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   81
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   17
      Left            =   1560
      Picture         =   "Pictures.frx":27B28
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   80
      Top             =   6180
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   17
      Left            =   1920
      Picture         =   "Pictures.frx":2801A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   79
      Top             =   6180
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   16
      Left            =   1560
      Picture         =   "Pictures.frx":2850C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   78
      Top             =   5820
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   16
      Left            =   1920
      Picture         =   "Pictures.frx":289FE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   77
      Top             =   5820
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   2280
      Picture         =   "Pictures.frx":28EF0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   76
      Top             =   5460
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   2640
      Picture         =   "Pictures.frx":293E2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   75
      Top             =   5460
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   2280
      Picture         =   "Pictures.frx":298D4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   74
      Top             =   5100
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   2640
      Picture         =   "Pictures.frx":29DC6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   73
      Top             =   5100
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   2280
      Picture         =   "Pictures.frx":2A2B8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   72
      Top             =   4740
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   2640
      Picture         =   "Pictures.frx":2A7AA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   71
      Top             =   4740
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   2280
      Picture         =   "Pictures.frx":2AC9C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   70
      Top             =   4380
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   2640
      Picture         =   "Pictures.frx":2B18E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   69
      Top             =   4380
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   2280
      Picture         =   "Pictures.frx":2B680
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   68
      Top             =   4020
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   2640
      Picture         =   "Pictures.frx":2BB72
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   67
      Top             =   4020
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   2280
      Picture         =   "Pictures.frx":2C064
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   66
      Top             =   3660
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   2640
      Picture         =   "Pictures.frx":2C556
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   65
      Top             =   3660
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   2280
      Picture         =   "Pictures.frx":2CA48
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   64
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   2640
      Picture         =   "Pictures.frx":2CF3A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   63
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   2280
      Picture         =   "Pictures.frx":2D42C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   62
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   2640
      Picture         =   "Pictures.frx":2D91E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   61
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   2280
      Picture         =   "Pictures.frx":2DE10
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   60
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   2640
      Picture         =   "Pictures.frx":2E302
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   59
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   2280
      Picture         =   "Pictures.frx":2E7F4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   58
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   2640
      Picture         =   "Pictures.frx":2ECE6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   57
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   2280
      Picture         =   "Pictures.frx":2F1D8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   56
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   2640
      Picture         =   "Pictures.frx":2F6CA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   55
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   2280
      Picture         =   "Pictures.frx":2FBBC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   54
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   2640
      Picture         =   "Pictures.frx":300AE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   53
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   2280
      Picture         =   "Pictures.frx":305A0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   52
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   2640
      Picture         =   "Pictures.frx":30A92
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   51
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   2280
      Picture         =   "Pictures.frx":30F84
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   50
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   2640
      Picture         =   "Pictures.frx":31476
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   49
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   2280
      Picture         =   "Pictures.frx":31968
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   48
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   2640
      Picture         =   "Pictures.frx":31E5A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   47
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2280
      Picture         =   "Pictures.frx":3234C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   46
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicmLines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2640
      Picture         =   "Pictures.frx":3283E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   45
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicmPlant 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   3780
      Picture         =   "Pictures.frx":32D30
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   44
      Top             =   480
      Width           =   600
   End
   Begin VB.PictureBox PicPlant 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   3120
      Picture         =   "Pictures.frx":34032
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   43
      Top             =   480
      Width           =   600
   End
   Begin VB.PictureBox PicPower 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3120
      Picture         =   "Pictures.frx":35334
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   42
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicmPower 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3480
      Picture         =   "Pictures.frx":35826
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   41
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   1920
      Picture         =   "Pictures.frx":35D18
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   40
      Top             =   5460
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   1560
      Picture         =   "Pictures.frx":3620A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   39
      Top             =   5460
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   1920
      Picture         =   "Pictures.frx":366FC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   38
      Top             =   5100
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   1560
      Picture         =   "Pictures.frx":36BEE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   37
      Top             =   5100
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   1920
      Picture         =   "Pictures.frx":370E0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   36
      Top             =   4740
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   1560
      Picture         =   "Pictures.frx":375D2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   35
      Top             =   4740
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   1920
      Picture         =   "Pictures.frx":37AC4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   34
      Top             =   4380
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   1560
      Picture         =   "Pictures.frx":37FB6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   33
      Top             =   4380
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   1920
      Picture         =   "Pictures.frx":384A8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   32
      Top             =   4020
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   1560
      Picture         =   "Pictures.frx":3899A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   31
      Top             =   4020
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   1920
      Picture         =   "Pictures.frx":38E8C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   30
      Top             =   3660
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   1560
      Picture         =   "Pictures.frx":3937E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   29
      Top             =   3660
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   1920
      Picture         =   "Pictures.frx":39870
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   28
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   1560
      Picture         =   "Pictures.frx":39D62
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   27
      Top             =   3300
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   1920
      Picture         =   "Pictures.frx":3A254
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   26
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   1560
      Picture         =   "Pictures.frx":3A746
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   25
      Top             =   2940
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   1920
      Picture         =   "Pictures.frx":3AC38
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   24
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   1560
      Picture         =   "Pictures.frx":3B12A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   23
      Top             =   2580
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   1920
      Picture         =   "Pictures.frx":3B61C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   22
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   1560
      Picture         =   "Pictures.frx":3BB0E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   21
      Top             =   2220
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   1920
      Picture         =   "Pictures.frx":3C000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   1560
      Picture         =   "Pictures.frx":3C4F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   1860
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   1920
      Picture         =   "Pictures.frx":3C9E4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   18
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   1560
      Picture         =   "Pictures.frx":3CED6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   17
      Top             =   1500
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   1920
      Picture         =   "Pictures.frx":3D3C8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   1560
      Picture         =   "Pictures.frx":3D8BA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   1140
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   1920
      Picture         =   "Pictures.frx":3DDAC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   14
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   1560
      Picture         =   "Pictures.frx":3E29E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   13
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1920
      Picture         =   "Pictures.frx":3E790
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1560
      Picture         =   "Pictures.frx":3EC82
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicmRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   1920
      Picture         =   "Pictures.frx":3F174
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicRoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   1560
      Picture         =   "Pictures.frx":3F666
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   9
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   840
      Picture         =   "Pictures.frx":3FB58
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1200
      Picture         =   "Pictures.frx":4004A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   840
      Picture         =   "Pictures.frx":4053C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicmBu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   1200
      Picture         =   "Pictures.frx":40A2E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   60
      Picture         =   "Pictures.frx":40F20
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   420
      Picture         =   "Pictures.frx":41412
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "Pictures.frx":41904
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   420
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   60
      Picture         =   "Pictures.frx":41DF6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   60
      Picture         =   "Pictures.frx":422E8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1140
      Width           =   300
   End
End
Attribute VB_Name = "Pictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PicsmlPark_Click(Index As Integer)

End Sub

