VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmckd_info 
   Caption         =   "CKD Info"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmckd_info.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer stck_updt 
      Left            =   4800
      Top             =   120
   End
   Begin VB.Frame framex 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   17655
      Begin VB.Frame Frame2 
         Caption         =   "Information"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   9255
         Begin VB.TextBox txtcp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "COST_PRICE"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "UNIT_PRICE"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   7680
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtdate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "LAST_ADDED_DATE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "BIKE_DETAILS"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            TabIndex        =   21
            Top             =   840
            Width           =   3855
         End
         Begin VB.TextBox txtsid 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "BID"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000000&
            Caption         =   "Cost Price"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000000&
            Caption         =   "Unit Price"
            Height          =   375
            Left            =   6000
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000000&
            Caption         =   "Details / Name"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000000&
            Caption         =   "Bike ID"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000000&
            Caption         =   "Stock Added Date"
            Height          =   375
            Left            =   6000
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "CKD TRANSACTION BETWEEN DATES"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6120
         TabIndex        =   15
         Top             =   6600
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   79233027
            CurrentDate     =   40805
         End
         Begin MSComCtl2.DTPicker fromdate 
            Height          =   375
            Left            =   2160
            TabIndex        =   18
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   79233027
            CurrentDate     =   40805
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CKD Report"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stock Balance"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9600
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtquan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "LAST_ADDED_QUANTITY"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtslvl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            DataField       =   "BIKE_LEVEL"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label llast 
            BackColor       =   &H80000000&
            Caption         =   "Last Added"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label ladd 
            BackColor       =   &H8000000D&
            Caption         =   "ADD NEW QUANTITY"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000000&
            Caption         =   "Current Stock Level"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdaddstock 
         Caption         =   "ADD Stock"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Update Entry"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "New Entry"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc adospares_info 
         Height          =   495
         Left            =   240
         Top             =   6720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
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
         RecordSource    =   "stock_ckd"
         Caption         =   "CKD Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmckd_info.frx":F172
         Height          =   3975
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "CKD INFORMATION"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2100
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmckd_info.frx":F18F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmckd_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rptallckd.Show
End Sub

Private Sub Command3_Click()
Dim prmfrom
Dim prmto
prmfrom = fromdate.Value
prmto = todate.Value
If dev_bike.rscmdCKDTRANBETWEENDATES.State = adStateOpen Then
  dev_bike.rscmdCKDTRANBETWEENDATES.Close
End If
 dev_bike.cmdCKDTRANBETWEENDATES prmfrom, prmto
rptckdtranbetwdates.Show
End Sub

Private Sub Form_Load()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmckd_info.Width - framex.Width) / 2
framex.Left = fpostn
todate = Date
fromdate = Date

End Sub

Private Sub Form_Resize()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmckd_info.Width - framex.Width) / 2
framex.Left = fpostn
End Sub
Private Sub cmbbid_Change()
txtbid.Text = cmbbid.Text
End Sub

Private Sub cmdadd_Click()
'from herecmd butn changes
    cmdadd.Enabled = False
    cmdref.Enabled = False
    cmdaddstock.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    ladd.Visible = True
    llast.Visible = False
    adospares_info.Recordset.AddNew
    txtdate = Format(Date, "dd/MM/yyyy")
    adospares_info.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    DataGrid1.Enabled = False
    txtslvl.Text = "00"
    stck_updt.Enabled = False
End Sub

Private Sub cmdaddstock_Click()
'from herecmd butn changes
On Error GoTo errorhandler

stockadd = InputBox("Type in numbers of stock you want to add to the CKD " & txtname.Text, "Add CKD Stock", "1000")
 If IsNumeric(stockadd) = True Then
    stockadd = Int(stockadd)
    If stockadd < 0 Then
        stockadd = -stockadd
    End If
        
        ask = MsgBox("Are you sure you want to add this to the stock?", vbInformation + vbYesNo, "System Information")
        If ask = vbYes Then
        txtquan = stockadd
        txtslvl = Val(txtquan) + Val(txtslvl)
        adospares_info.Recordset.Save
        adospares_info.Refresh
        adospares_info.Refresh
        MsgBox "Stock added to the database", vbInformation + vbOKOnly, "System Informtaion"
        ElseIf ask = vbNo Then
        adospares_info.Recordset.Cancel
        MsgBox "Stock NOT added to the database", vbInformation + vbOKOnly, "System Informtaion"
        End If
   
 ElseIf IsNumeric(stockadd) = False Then
 MsgBox "Please input a numeric value and try again", vbCritical + vbOKOnly, "System Error"
End If
 
 
 
 
'cmdadd.Enabled = False
adospares_info.Enabled = True
'cmdref.Enabled = False
'cmdaddstock.Enabled = False
'cmdsave.Enabled = True
'cmdcancel.Enabled = True
'ladd.Visible = True
'from here txt box locking disabled
'txtname.Locked = False
'txtquan.Locked = False
txtup.Locked = False
txtsid.Locked = False


errorhandler:
If Err.Number = -2147217842 Then
    MsgBox "Please refresh the database and try again", vbCritical + vbOKOnly, "System Error"
End If
End Sub

Private Sub cmdcancel_Click()
adospares_info.Recordset.Cancel
cmdadd.Enabled = True
cmdref.Enabled = True
cmdsave.Enabled = False
cmdaddstock.Enabled = True
adospares_info.Refresh
ladd.Visible = False
llast.Visible = True
cmdcancel.Enabled = False

    adospares_info.Enabled = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    DataGrid1.Enabled = True

MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"


End Sub

Private Sub cmdref_Click()
adospares_info.Refresh

 MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"
 
End Sub

Private Sub cmdsave_Click()
On Error GoTo erroronproduction:
If txtsid = "" Then
    MsgBox "Please input CKD ID", vbExclamation + vbOKOnly, "System Error"
    txtsid.SetFocus
ElseIf txtquan = "" Then
    MsgBox "Please input quantity to be added", vbExclamation + vbOKOnly, "System Error"
    txtquan.SetFocus
ElseIf txtup = "" Then
    MsgBox "Please input unit price", vbExclamation + vbOKOnly, "System Error"
    txtup.SetFocus
ElseIf txtcp = "" Then
    MsgBox "Please input cost price", vbExclamation + vbOKOnly, "System Error"
    txtcp.SetFocus
ElseIf Val(txtcp) > Val(txtup) Then
        MsgBox "Cost price cannot be greater than unit price!", vbExclamation + vbOKOnly, "System Error"
Else:
    txtslvl = Val(txtquan) + Val(txtslvl)
    adospares_info.Recordset.Save
    adospares_info.Refresh
    MsgBox "Database updated", vbInformation + vbOKOnly, "System Informtaion"
    cmdadd.Enabled = True
    cmdref.Enabled = True
    cmdsave.Enabled = False
    cmdaddstock.Enabled = True
    adospares_info.Refresh
    ladd.Visible = False
    llast.Visible = True
    cmdcancel.Enabled = False
    adospares_info.Enabled = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    DataGrid1.Enabled = True
    stck_updt.Enabled = True
End If
erroronproduction:
    If Err.Number <> 0 Then
        'MsgBox "Operation is cancelled because you have already used this CKD ID. Try again choosing different ID. ", vbCritical + vbOKOnly, "Database Error"
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"

End If
End Sub

Private Sub DTPicker1_Change()
txtdate.Text = DTPicker1.Value
End Sub


Private Sub Form_Unload(Cancel As Integer)
If cmdadd.Enabled = False Then
    MsgBox "SORRY YOU CANT QUIT NOW, YOU ARE IN A MIDDLE OF A PROCESS", vbExclamation
    Cancel = 1
Else
ASKQUIT = MsgBox("ARE YOU SURE YOU WANT TO QUIT " & Me.Caption & " ?", vbQuestion + vbYesNo)
    If ASKQUIT = vbYes Then
    Cancel = 0
    ElseIf ASKQUIT = vbNo Then
    Cancel = 1
    End If
End If
End Sub

Private Sub txtcp_KeyUp(KeyCode As Integer, Shift As Integer)

If IsNumeric(txtcp) = False Then
        txtcp.Text = ""
End If
End Sub

Private Sub txtquan_Change()
If IsNumeric(txtquan) = False Then
        txtquan.Text = ""
End If
End Sub

Private Sub txtup_KeyUp(KeyCode As Integer, Shift As Integer)

If IsNumeric(txtup) = False Then
        txtup.Text = ""
End If
End Sub
