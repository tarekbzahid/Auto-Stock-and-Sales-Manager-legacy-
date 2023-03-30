VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsetdealerbalnc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dealer Balance Info"
   ClientHeight    =   9045
   ClientLeft      =   6195
   ClientTop       =   1200
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsetdealerbalnc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Update Balance"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   8895
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "TOTAL_AMOUNT"
         DataSource      =   "adodelr"
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "TOTAL_PAID"
         DataSource      =   "adodelr"
         Height          =   375
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "TOTAL_DUE"
         DataSource      =   "adodelr"
         Height          =   375
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Total Amount / Tk"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Total Paid / Tk"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Total Due / Tk"
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Report"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   7680
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adodelr 
      Height          =   375
      Left            =   120
      Top             =   7680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   $"frmsetdealerbalnc.frx":F172
      Caption         =   "Dealer with Dues"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmsetdealerbalnc.frx":F204
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
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
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Dealer Payment / Tk"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "DEALER BALANCE (DUES)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   0
      Picture         =   "frmsetdealerbalnc.frx":F21A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9345
   End
End
Attribute VB_Name = "frmsetdealerbalnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
adodelr.Refresh
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"

End Sub

Private Sub Command2_Click()
rptdealer_dues.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text4.Text = "" Then
    MsgBox "Input dealer's payment", vbExclamation + vbOKOnly, "System Informtaion"
Else
    ask = MsgBox("Do you want to update dealer's payment ?", vbQuestion + vbYesNo, "System Query")
        If ask = vbYes Then
           If Val(Text4) > Val(Text3) Then
                MsgBox "Payment cannot be more than dues!", vbExclamation + vbOKOnly, "System Informtaion"
                Text4 = ""
            Else
                Text2 = Val(Text2) + Val(Text4)
                Text3 = Val(Text3) - Val(Text4)
                Dim due
                due = Text3
                adodelr.Recordset.Update
                adodelr.Recordset.Save
                adodelr.Refresh
                adodelr.Refresh
                MsgBox "Database Updated. The dealer's current due is " & due, vbExclamation + vbOKOnly, "System Informtaion"
            End If
        End If
End If
End Sub

Private Sub Text4_Change()
If IsNumeric(Text4) = False Then
        Text4.Text = ""
End If
If Val(Text4) > Val(Text3) Then
    MsgBox "Payment cannot be more than dues!", vbExclamation + vbOKOnly, "System Informtaion"
Text4 = ""
End If
End Sub

