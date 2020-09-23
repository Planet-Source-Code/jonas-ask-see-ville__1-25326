VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   300
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1980
      Width           =   1515
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   180
      Width           =   195
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   1590
      Left            =   2940
      Pattern         =   "*.bmp"
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicName As String

Private Sub Command1_Click()
Buffer.Cls
    For y = 0 To picMap.ScaleHeight - 1
        For x = 0 To picMap.ScaleWidth - 1
            temp = picMap.Point(x, y)
            If Not temp = 0 Then
                Buffer.PSet (x, y)
            End If
        Next x
    Next y
    SavePicture Buffer.Image, App.Path & "\m" & PicName & ".bmp"
End Sub

Private Sub File1_Click()
Dim File As String
    On Error Resume Next
    File = File1.FileName
    If File = Empty Then Exit Sub
    
    picMap.AutoSize = True
    picMap.Picture = LoadPicture(App.Path & "\" & File)
    picMap.AutoSize = False
    
    Buffer.Width = picMap.Width
    Buffer.Height = picMap.Height
    
    PicName = Mid(File, 1, Len(File) - 4)
    
    Buffer.Cls
    For y = 0 To picMap.ScaleHeight - 1
        For x = 0 To picMap.ScaleWidth - 1
            temp = picMap.Point(x, y)
            If Not temp = 0 Then
                Buffer.PSet (x, y)
            End If
        Next x
    Next y
End Sub

Private Sub Form_Load()
    File1.Path = App.Path
End Sub

