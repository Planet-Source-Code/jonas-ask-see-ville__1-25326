VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Game"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   2
      Top             =   2760
      Width           =   915
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3300
      TabIndex        =   1
      Top             =   2760
      Width           =   915
   End
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
      Left            =   180
      Pattern         =   "*.sav"
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
    If Dirty Then
        Svar = MsgBox("The game is not saved. Continue?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Exit Sub
    End If
    Me.Hide
    LoadGame File1.Path & "\" & File1.FileName
End Sub

Private Sub Form_Activate()
    frmSave.Hide
    frmNew.Hide
    frmInfo.Hide
    File1.Path = App.Path & "\saved\"
    File1.Refresh
End Sub

