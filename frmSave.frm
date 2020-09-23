VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Game"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      BackColor       =   &H00FF8080&
      CausesValidation=   0   'False
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
      Height          =   345
      Left            =   180
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2880
      Width           =   2115
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
      TabIndex        =   2
      Top             =   60
      Width           =   2115
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3300
      TabIndex        =   1
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   2880
      Width           =   915
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSave_Click()
Dim Svar
Dim a
Dim Text As String
    Text = txtName.Text
    If Text = "" Then Exit Sub
    For a = 1 To Len(Text)
        If Mid(Text, a, 1) = "." Then MsgBox "Invalid name", vbOKOnly + vbCritical, GameTitle & " - ERROR": Exit Sub
    Next
    If Dir(App.Path & "\saved\" & Text & ".sav") <> "" Then
        Svar = MsgBox("Game already exsist. Overwrite?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Exit Sub
    End If
    Me.Hide
    SaveGame App.Path & "\saved\" & Text & ".sav"
End Sub

Private Sub File1_Click()
    txtName.Text = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
End Sub

Private Sub Form_Activate()
    frmLoad.Hide
    frmNew.Hide
    frmInfo.Hide
    File1.Path = App.Path & "\saved\"
    File1.Refresh
End Sub

