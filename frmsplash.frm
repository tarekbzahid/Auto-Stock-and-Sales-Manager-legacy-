VERSION 5.00
Begin VB.Form frmsplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   2835
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   0
      Picture         =   "frmsplash.frx":F172
      ScaleHeight     =   2805
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5760
         Top             =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "by tarek bin zahid "
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub



Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCompany_Click()
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "You already have one instance of BORAC running!", vbCritical, "System Error"
Unload Me
Else
Label1.Caption = "Application Thread ID " & " " & App.ThreadID
Timer1.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
    Unload Me
    frmlogin.Show
    Timer1.Enabled = False
End Sub


