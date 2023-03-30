VERSION 5.00
Begin VB.Form frmbrwsr 
   ClientHeight    =   5130
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6540
   Icon            =   "frmbrwsr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3795
   End
   Begin VB.PictureBox tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   6540
   End
   Begin VB.PictureBox brwWebBrowser 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   5340
      TabIndex        =   0
      Top             =   1320
      Width           =   5400
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin VB.PictureBox imlIcons 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   2670
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   2325
      Width           =   1200
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "&Address:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Tag             =   "&Address:"
      Top             =   600
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "frmbrwsr.frx":F172
      Stretch         =   -1  'True
      Top             =   480
      Width           =   15225
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "&Address:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Tag             =   "&Address:"
      Top             =   0
      Width           =   3075
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "&Address:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Tag             =   "&Address:"
      Top             =   0
      Width           =   3075
   End
End
Attribute VB_Name = "frmbrwsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()

    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize

    cboAddress.Move 50, lblAddress(2).Top + lblAddress(2).Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        'brwWebBrowser.Navigate StartingAddress
    End If
Image1.Width = Me.Width
End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    'Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    'Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        'If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        'End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    'cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
   ' brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (Image1.Top + Image1.Height) - 100
End Sub

Private Sub timTimer_Timer()
   ' If brwWebBrowser.Busy = False Then
        'timTimer.Enabled = False
    '    'Me.Caption = brwWebBrowser.LocationName
   ' Else
        'Me.Caption = "Working..."
    'End If
End Sub

'Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
   ' On Error Resume Next
     
   ' timTimer.Enabled = True
     
    'Select Case Button.Key
       ' Case "Back"
          '  brwWebBrowser.GoBack
        'Case "Forward"
         '   brwWebBrowser.GoForward
       ' Case "Refresh"
         '   brwWebBrowser.Refresh
        'Case "Home"
         '   brwWebBrowser.Navigate "www.bdp-bd.com"
        'Case "Search"
          '  brwWebBrowser.GoSearch
        'Case "Stop"
           ' timTimer.Enabled = False
            'brwWebBrowser.Stop
            'Me.Caption = brwWebBrowser.LocationName
    'End Select

'End Sub

