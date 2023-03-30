VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000003&
   Caption         =   "Bike Selling and Management System"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8445
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmmain.frx":F172
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer user_check 
      Interval        =   1
      Left            =   3480
      Top             =   1320
   End
   Begin VB.Timer tmr_main 
      Interval        =   1
      Left            =   12840
      Top             =   480
   End
   Begin MSComctlLib.StatusBar stsbr_main 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5450
            MinWidth        =   3545
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4445
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
      MousePointer    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu cmdpower 
      Caption         =   "&System"
      Begin VB.Menu cmdrun 
         Caption         =   "Run on System Startup"
      End
      Begin VB.Menu hash3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdlogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu cmdexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu cmdfile 
      Caption         =   "&Data Entry"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
      Begin VB.Menu cmdhome 
         Caption         =   "HOME"
      End
      Begin VB.Menu hash1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdd_info 
         Caption         =   "Dealer Info"
      End
      Begin VB.Menu cmdord 
         Caption         =   "Order"
      End
      Begin VB.Menu hash2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdckd_info 
         Caption         =   "CKD Info"
      End
      Begin VB.Menu cmdass_info 
         Caption         =   "Bike Info"
      End
      Begin VB.Menu cmdinv_info 
         Caption         =   "Inventory Info"
      End
   End
   Begin VB.Menu cmdtrs 
      Caption         =   "&Sales"
      Begin VB.Menu cmdckd_sale 
         Caption         =   "CKD Sale"
      End
      Begin VB.Menu cmdb_sale 
         Caption         =   "Bike Sales"
      End
      Begin VB.Menu cmdinv_sale 
         Caption         =   "Inventory Sale"
      End
   End
   Begin VB.Menu cmdsystem 
      Caption         =   "&Borac"
      Begin VB.Menu cmdpcalc 
         Caption         =   "Profit Calculator"
      End
      Begin VB.Menu cmdstckclimit 
         Caption         =   "Low Inventory"
      End
      Begin VB.Menu cmdchnguser 
         Caption         =   "Add/Change User"
      End
      Begin VB.Menu Log 
         Caption         =   "User Log"
      End
      Begin VB.Menu cmddaccount 
         Caption         =   "Damage Account"
      End
   End
   Begin VB.Menu cmdabout 
      Caption         =   "&About"
      Begin VB.Menu cmdviewabout 
         Caption         =   "About "
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shutdown As Boolean

Private Sub cmdbike_Click()
    frmstock_mangmnt.Show
End Sub

Private Sub cmdass_info_Click()
frmbike_info.Show
End Sub

Private Sub cmdb_sale_Click()
frmbike_sales.Show
End Sub

Private Sub cmdbrwsr_Click()
    frmbrwsr.Show
End Sub

Private Sub cmdchnguser_Click()
    frmchnguser.Show
End Sub

Private Sub cmddeabaln_Click()
    frmdelrmng.Show
End Sub

Private Sub cmddeainfo_Click()
    frmdealer.Show
End Sub


Private Sub cmdckd_info_Click()
frmckd_info.Show
End Sub

Private Sub cmdckd_sale_Click()
frmckd_sales.Show
End Sub

Private Sub cmdd_info_Click()
frmdealerinfo.Show
End Sub

Private Sub cmddelr_bal_Click()
frmsetdealerbalnc.Show
End Sub

Private Sub cmddaccount_Click()
'frmdaccount.Show
MsgBox "UNDER CONSTRUCTION PLEASE BE PATIENT", vbInformation + vbOKOnly, "SYSTEM INFO"
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdhome_Click()
    frmhome.Show
End Sub

Private Sub cmdinv_info_Click()
frmspares_info.Show
End Sub

Private Sub cmdinv_sale_Click()
frmspares_sales.Show
End Sub

Private Sub cmdlogout_Click()
query = (MsgBox("Do you want to log out?", vbYesNo + vbQuestion, "System Log Out"))
If query = vbYes Then
    frmlogin.adologin.RecordSource = "SELECT * FROM login WHERE (((login.UID) Like '" & frmmain.stsbr_main.Panels(2).Text & "%'))"
    frmlogin.adologin.CommandType = adCmdText
    frmlogin.adologin.Refresh
    
    Logg = Format(Date, "dd/MM/YYYY") & " /---/ " & Format(Time, "hh:mm AM/PM")
            frmlogin.adologin.Recordset.Fields(4) = Logg
            frmlogin.adologin.Recordset.Fields(7) = "OFFLINE"
    
    frmlogin.adologin.Recordset.Update
    frmlogin.adologin.Refresh
    
    shutdown = True
    'MsgBox frmlogin.adologin.Recordset.Fields(0).Value
    Unload Me
    frmlogin.Show
    shutdown = False
    
frmlogin.adologin.RecordSource = "SELECT * FROM login "
frmlogin.adologin.CommandType = adCmdText
frmlogin.adologin.Refresh


ElseIf query = vbNo Then
Cancel = 1
End If
End Sub

Private Sub cmdord_Click()
    frmorder.Show
End Sub

Private Sub cmdpdlvry_Click()
    frmdelivery.Show
End Sub

Private Sub cmdstck_Click()
  
End Sub

Private Sub cmdpcalc_Click()
frmgprftcalc.Show vbModal
End Sub

Private Sub cmdrun_Click()
runques = MsgBox("Do you want to run this software on system startup ?", vbQuestion + vbYesNo, "System Query")
If runques = vbYes Then
SetRegValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run", "BORAC", "C:\dbase_bike\BORAC.exe"
ElseIf runques = vbNo Then
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "BORAC"
End If
End Sub

Private Sub cmdsrch_Click()
    frmsearch.Show
End Sub

Private Sub cmdssprs_sales_Click()
    frmspares_sales.Show
End Sub

Private Sub cmdstckclimit_Click()
   frmsetclevel.Show
End Sub

Private Sub cmdviewabout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub cmdviewlog_Click()
    frmlog.Show
End Sub

Private Sub Log_Click()
frmuserstatus.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If shutdown = False Then
query = (MsgBox("Are you sure you want to quit BORAC Sales System ?", vbYesNo + vbQuestion, "System Shut Down"))
If query = vbYes Then
 frmlogin.adologin.RecordSource = "SELECT * FROM login WHERE (((login.UID) Like '" & frmmain.stsbr_main.Panels(2).Text & "%'))"
    frmlogin.adologin.CommandType = adCmdText
    frmlogin.adologin.Refresh
    
    Logg = Format(Date, "dd/MM/YYYY") & " /---/ " & Format(Time, "hh:mm AM/PM")
            frmlogin.adologin.Recordset.Fields(4) = Logg
            frmlogin.adologin.Recordset.Fields(7) = "OFFLINE"
    
    frmlogin.adologin.Recordset.Update
    frmlogin.adologin.Refresh
    
End
End
ElseIf query = vbNo Then
Cancel = 1
End If
End If
End Sub

Private Sub tmr_main_Timer()
    
    stsbr_main.Panels(3).Text = Format(Now, "d-mmmm h:mm AM/PM")
End Sub

Private Sub user_check_Timer()

Select Case frmmain.stsbr_main.Panels(1).Text

                Case "Sales"
                    frmmain.cmdfile.Enabled = False
                    frmmain.cmdsystem.Enabled = False
                    frmmain.Show
                Case "Stock"
                    frmmain.cmdtrs.Enabled = False
                    frmmain.cmdsystem.Enabled = False
                    frmmain.Show
                Case "administrator"
                    MsgBox "Welcome Administrator !", vbInformation + vbOKOnly, "System Information"
End Select
user_check.Enabled = False
End Sub
