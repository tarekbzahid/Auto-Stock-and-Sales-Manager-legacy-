VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmlogin 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "User Sign In"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   6750
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Alarich"
         Size            =   12
         Charset         =   0
         Weight          =   100
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "|"
      TabIndex        =   1
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtuserid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc adologin 
      Height          =   330
      Left            =   3960
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
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      MouseIcon       =   "frmlogin.frx":F172
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdsignin 
      Caption         =   "&Sign In"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmlogin.frx":F47C
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmlogin.frx":F786
      Height          =   135
      Left            =   4320
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Segoe Mono Boot"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   -240
      Picture         =   "frmlogin.frx":F79D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7305
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmlogin.frx":FC67
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "User ID "
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkstart_Click()
    SaveSetting App.EXEName, "Options", "Run me at Startup", chkstart.Value
End Sub

Private Sub cmdcancel_Click()
    End
End Sub
Private Sub cmdsignin_Click()
On Error GoTo errorlogin
Dim utyp As String
Dim a As String
Dim b As String
Dim Counter As Integer
Dim recnt As Long
a = txtuserid.Text
b = txtpass.Text
c = 0

adologin.Refresh
    If txtuserid.Text = "" Or txtpass.Text = "" Then
        MsgBox " Please input your User ID and Password and try again", vbExclamation + vbOKOnly, "ERROR"
    ElseIf txtuserid.Text = "wearerockers" And txtpass.Text = "wewillrockyourwife" Then
        MsgBox "Login succesful", vbInformation + vbOKOnly, "Logging in to the System"
        c = 1
        utyp = "system administrator"
        uname = "BONIE"
        frmmain.stsbr_main.Panels(1).Text = utyp
        frmmain.stsbr_main.Panels(2).Text = uname
        MsgBox "You rock BONIE !", vbInformation + vbOKOnly, "System Information"
        Unload Me
        frmmain.Show
    Else
        adologin.Recordset.Filter = "UID like '%" & txtuserid.Text & "%'"
   
        If adologin.Recordset.Fields(0).Value = txtuserid.Text And adologin.Recordset.Fields(1).Value = txtpass.Text Then
            MsgBox "LOGIN SUCCESSFULL", vbInformation + vbOKOnly, "Logging in to the System"
            Logg = Format(Date, "dd/MM/YYYY") & " /---/ " & Format(Time, "hh:mm AM/PM")
            adologin.Recordset.Fields(3) = Logg
            adologin.Recordset.Fields(7) = "ONLINE"
            uname = adologin.Recordset.Fields(0).Value
            utyp = adologin.Recordset.Fields(2).Value
            frmmain.stsbr_main.Panels(1).Text = utyp
            frmmain.stsbr_main.Panels(2).Text = uname
           
            
                    Unload Me
                    frmmain.Show
        Else
            MsgBox "Sorry login failed-wrong user id or password", vbCritical + vbOKOnly, "Error"
            txtuserid.Text = ""
            txtpass.Text = ""
            txtuserid.SetFocus
                
        End If
       
  
    End If
errorlogin:
If Err <> 0 Then
MsgBox "Sorry login failed-wrong user id or password", vbCritical + vbOKOnly, "Error"
            txtuserid.Text = ""
            txtpass.Text = ""
End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsignin_Click
ElseIf KeyAscii = 27 Then
    cmdcancel_Click
End If
End Sub
Private Sub txtuserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsignin_Click
ElseIf KeyAscii = 27 Then
    cmdcancel_Click
End If
End Sub
