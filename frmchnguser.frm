VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmchnguser 
   BorderStyle     =   0  'None
   Caption         =   "User Panel"
   ClientHeight    =   6720
   ClientLeft      =   3075
   ClientTop       =   435
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmchnguser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmchnguser.frx":F172
   ScaleHeight     =   6720
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      Top             =   6240
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adouser 
      Height          =   330
      Left            =   4440
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select User Options"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
      Begin VB.OptionButton optchnhuser 
         Caption         =   "Change User Info"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optdeluser 
         Caption         =   "Delete user"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optaddnwuser 
         Caption         =   "Add New User"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame frameaddnwuser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   28
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   29
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdadduser 
         Caption         =   "Add New User"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmchnguser.frx":13630
         Left            =   2400
         List            =   "frmchnguser.frx":1363A
         TabIndex        =   6
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Type New Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Retype Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Select User Access Level"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame framechnguser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmchnguser.frx":1364C
         Height          =   315
         Left            =   2280
         TabIndex        =   30
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "UID"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtchckpass 
         Appearance      =   0  'Flat
         DataSource      =   "adouser"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         DataSource      =   "adouser"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox txtuserid 
         Appearance      =   0  'Flat
         DataSource      =   "adouser"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdchng 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         DataSource      =   "adouser"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmchnguser.frx":13662
         Left            =   2280
         List            =   "frmchnguser.frx":1366C
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Change User ID"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Retype Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Select User"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Select User Access Level"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Type New Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
   End
   Begin VB.Frame framedeluser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmddeluser 
         Caption         =   "Delete User"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmchnguser.frx":1367E
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "UID"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Select User"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "USER PANEL"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "frmchnguser.frx":13694
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmchnguser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadduser_Click()

adouser.Recordset.AddNew
cmdupdate.Enabled = True
cmdadduser.Enabled = False
cmdadduser.Enabled = False
cmdupdate.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Combo4.Enabled = True
handleerrorS:

End Sub


Private Sub cmdcancel_Click()
    adouser.Recordset.Cancel
    adouser.Refresh
    cmdadduser.Enabled = True
    cmdupdate.Enabled = False
    Text3.Enabled = False
    Text3 = ""
Text4.Enabled = False
Text4 = ""
Text5.Enabled = False
Text5 = ""
Combo4.Enabled = False
End Sub

Private Sub cmdchng_Click()
On Error GoTo HandleAddDataErrors
If adouser.Recordset.Fields(2).Value = "administrator" Then
MsgBox "Sorry you cannot change administator information !", vbExclamation + vbOKOnly, "System Error"
ElseIf Combo3 = "" Or txtuserid = "" Or txtpass = "" Or txtchckpass = "" Then
MsgBox "Some fields are blank. Please fill up all the fields.", vbInformation + vbOKOnly, "System Error"
ElseIf txtpass.Text <> txtchckpass.Text Then
MsgBox "Password mismatch", vbExclamation + vbOKOnly, "System Error"
Else
adouser.Recordset.Fields(0).Value = txtuserid.Text
adouser.Recordset.Fields(1).Value = txtpass.Text
adouser.Recordset.Fields(2).Value = Combo3.Text
MsgBox "User data changed successfully", vbInformation + vbOKOnly, "System Information"
adouser.Recordset.Save
adouser.Refresh
DataCombo1.Refresh
adouser.Refresh
DataCombo1.Refresh
DataCombo1.ReFill
End If
HandleAddDataErrors:
If Err.Number = -2147217842 Then
MsgBox "Operation is cancelled because there is already a user ID with this name.Try again with a different user ID ", vbCritical + vbOKOnly, "Database Error"
 adouser.Recordset.Cancel
 End If
End Sub

Private Sub cmddeluser_Click()
If adouser.Recordset.Fields(2).Value = "administrator" Then
MsgBox "Sorry you cannot delete administator !", vbExclamation + vbOKOnly, "System Error"
Else:
ask = MsgBox("Do you want to delete this user ?", vbQuestion + vbYesNo, "System Query")
    If ask = vbYes Then
        adouser.Recordset.Delete
        adouser.Refresh
        MsgBox "User Deleted", vbInformation + vbOKOnly, "System Information"
        adouser.Refresh
        adouser.Refresh
        DataCombo1.Refresh
        DataCombo1.ReFill
        adouser.Refresh
    End If
End If

End Sub

Private Sub cmdupdate_Click()
On Error GoTo handle
If Text3 = "" Or Text4 = "" Or Text5 = "" Or Combo4 = "" Then
MsgBox "Some fields are blank. Please fill up all the fields.", vbInformation + vbOKOnly, "System Error"
ElseIf Text5 <> Text4 Then
MsgBox "Password donot match", vbExclamation + vbCritical + vbOKOnly, "Error"
Else:
adouser.Recordset.Fields(0) = Text3
adouser.Recordset.Fields(1) = Text4
adouser.Recordset.Fields(2) = Combo4.Text
adouser.Recordset.Save
adouser.Refresh
MsgBox "User added", vbInformation + vbOKOnly, "System Information"
cmdadduser.Enabled = True
cmdupdate.Enabled = False
Text3.Enabled = False
Text3 = ""
Text4.Enabled = False
Text4 = ""
Text5.Enabled = False
Text5 = ""
Combo4.Enabled = False
End If
handle:
'If Err.Description=run time error Then
'MsgBox "Operation is cancelled because there is already a user ID with this name.Try again with a different user ID ", vbCritical + vbOKOnly, "Database Error"
'adouser.Recordset.Cancel
'End If
MsgBox Err.Description, vbCritical, "System Error"
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()
adouser.Recordset.Bookmark = DataCombo1.SelectedItem

End Sub


Private Sub DataCombo2_Change()
adouser.Recordset.Bookmark = DataCombo2.SelectedItem

End Sub

Private Sub Form_Load()
Image2.Width = Me.Width
frmchnguser.Move (Screen.Width - frmchnguser.Width) / 2, (Screen.Height - frmchnguser.Height) / 2
frameaddnwuser.Visible = False
framedeluser.Visible = False
framechnguser.Visible = False
optaddnwuser.Value = False
optchnhuser.Value = False
optdeluser.Value = False
End Sub

Private Sub optaddnwuser_Click()
If optaddnwuser.Value = True Then
    frameaddnwuser.Visible = True
    framedeluser.Visible = False
    framechnguser.Visible = False
End If
End Sub

Private Sub optchnhuser_Click()
If optchnhuser.Value = True Then
    frameaddnwuser.Visible = False
    framedeluser.Visible = False
    framechnguser.Visible = True
End If
End Sub

Private Sub optdeluser_Click()
If optdeluser.Value = True Then
    frameaddnwuser.Visible = False
    framedeluser.Visible = True
    framechnguser.Visible = False
End If
End Sub

Private Sub Text6_Change()

End Sub

Private Sub Text5_Change()
If Text5.Text = Text4 Then
Text5.ForeColor = &HC000&
ElseIf Text5.Text <> Text4 Then
Text5.ForeColor = &HFF&
End If
End Sub
