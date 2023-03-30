VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmgprftcalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gross Profit Calculator"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmgprftcalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmgprftcalc.frx":F172
   ScaleHeight     =   7410
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "stock"
      Height          =   1815
      Left            =   12480
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      Begin MSAdodcLib.Adodc adocap 
         Height          =   330
         Left            =   120
         Top             =   240
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
         RecordSource    =   "SELECT SUM(PROFIT) FROM invoice_spares as TPROFIT"
         Caption         =   "Adodc1"
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "sales"
      Height          =   1815
      Left            =   12480
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      Begin MSAdodcLib.Adodc adockd 
         Height          =   330
         Left            =   120
         Top             =   240
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
         RecordSource    =   "SELECT SUM(PROFIT) FROM invoice_spares as TPROFIT"
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc adobike 
         Height          =   330
         Left            =   120
         Top             =   600
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
         RecordSource    =   "invoice_assembled"
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc adospares 
         Height          =   330
         Left            =   120
         Top             =   960
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
         CommandType     =   8
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
         RecordSource    =   ""
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc adoorders 
         Height          =   330
         Left            =   120
         Top             =   1320
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
         CommandType     =   8
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select date range "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11415
      Begin VB.Frame Frame6 
         Caption         =   "Current Capital of Borac (AS NOW)"
         Height          =   4095
         Left            =   5760
         TabIndex        =   26
         Top             =   1080
         Width           =   5415
         Begin VB.Label Label33 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   40
            Top             =   3600
            Width           =   2895
         End
         Begin VB.Label Label32 
            BackColor       =   &H80000003&
            Caption         =   "TOTAL CAPITAL OF BORAC/BDT"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   3600
            Width           =   2175
         End
         Begin VB.Label Label31 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   38
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label30 
            BackColor       =   &H80000003&
            Caption         =   "->LESS Total expense/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label29 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   36
            Top             =   2280
            Width           =   2895
         End
         Begin VB.Label Label28 
            BackColor       =   &H80000003&
            Caption         =   "->ADD Apprx net profit/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   34
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000003&
            Caption         =   "->Debtor/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   32
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   31
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   30
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000003&
            Caption         =   "->BIKE stock/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000003&
            Caption         =   "->Spares stock/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000003&
            Caption         =   "->CKD stock/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   240
         TabIndex        =   19
         Top             =   3480
         Width           =   5415
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   2400
            TabIndex        =   20
            ToolTipText     =   "enter the amount of approximate expenditure to get an approximate net profit"
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000003&
            Caption         =   "APPROX NET PROFIT/BDT"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000003&
            Caption         =   "TOTAL EXPENDITURE/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000003&
            Caption         =   "TOTAL GROSS PROFIT/BDT"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   22
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            Top             =   1200
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate "
         Height          =   375
         Left            =   7680
         TabIndex        =   16
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   15
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Profit calculation breakdown"
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   5415
         Begin VB.Label Label4 
            BackColor       =   &H80000003&
            Caption         =   "->CKD sales profit/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000003&
            Caption         =   "->Spares sales profit/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000003&
            Caption         =   "->BIKE sales profit/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000003&
            Caption         =   "->Order received profit/BDT"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label label10 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   6
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   3
            Top             =   840
            Width           =   2895
         End
      End
      Begin MSComCtl2.DTPicker fromdate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
         CustomFormat    =   "dd/MM/yyy"
         Format          =   78970883
         CurrentDate     =   40803
      End
      Begin MSComCtl2.DTPicker todate 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   78970883
         CurrentDate     =   41021
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "NB: CALCULATED AS, GROSS PROFIT= GRAND TOTAL OF TRANSACTION/ORDER - COST PRICE OF GOODS"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   6120
         Width           =   10935
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "End Date"
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "GROSS PROFIT"
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
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmgprftcalc.frx":1E2E4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11985
   End
End
Attribute VB_Name = "frmgprftcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

On Error GoTo errorhandler
    Dim vartotalprofit As Long
    Dim varckdprofit As Long
    Dim varbikeprofit As Long
    Dim varsparesprofit As Long
    
    
    'calculating ckd profit
    adockd.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adockd.CommandType = adCmdText
    adockd.RecordSource = "SELECT SUM(PROFIT) AS varckdprofit FROM invoice_ckd WHERE (((invoice_ckd.T_DATE) BETWEEN #" & fromdate.Value & "#  AND  #" & todate.Value & "#)) "
    adockd.Refresh
    If adockd.Recordset.Fields!varckdprofit <> 0 Then
        label10.Caption = adockd.Recordset.Fields!varckdprofit
    Else: label10.Caption = 0
    End If
    
    'calculating bike profit
    adobike.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adobike.CommandType = adCmdText
    adobike.RecordSource = "SELECT SUM(PROFIT) AS varbikeprofit FROM invoice_assembled WHERE (((invoice_assembled.T_DATE) BETWEEN #" & Format(fromdate.Value, "dd/MM/yyyy") & "#  AND #" & Format(todate.Value, "dd/MM/yyyy") & "#)) "
    adobike.Refresh
    If adobike.Recordset.Fields!varbikeprofit <> 0 Then
        Label13.Caption = adobike.Recordset.Fields!varbikeprofit
    Else: Label13.Caption = 0
    End If
    
    'calculating spares profit
    adospares.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adospares.CommandType = adCmdText
    adospares.RecordSource = "SELECT SUM(PROFIT) AS varsparesprofit FROM invoice_spares WHERE (((invoice_spares.T_DATE) BETWEEN #" & Format(fromdate.Value, "dd/MM/yyyy") & "#  AND #" & Format(todate.Value, "dd/MM/yyyy") & "#)) "
    adospares.Refresh
    If adospares.Recordset.Fields!varsparesprofit <> 0 Then
        Label12.Caption = adospares.Recordset.Fields!varsparesprofit
    Else: Label12.Caption = 0
    End If
    
    'calculating orders profit
    adoorders.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adoorders.CommandType = adCmdText
    adoorders.RecordSource = "SELECT SUM(PROFIT) AS varordersprofit FROM dealer_order WHERE (((dealer_order.ORDER_DATE) BETWEEN #" & Format(fromdate.Value, "dd/MM/yyyy") & "#  AND #" & Format(todate.Value, "dd/MM/yyyy") & "#)) "
    adoorders.Refresh
    If adoorders.Recordset.Fields!varordersprofit <> 0 Then
        Label11.Caption = adoorders.Recordset.Fields!varordersprofit
    Else: Label11.Caption = 0
    End If
    
    'calculating ckd capital
    adocap.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adocap.CommandType = adCmdText
    adocap.RecordSource = "SELECT SUM(BIKE_LEVEL*UNIT_PRICE) AS cap FROM stock_ckd"
    adocap.Refresh
    If adocap.Recordset.Fields!cap <> 0 Then
        Label22.Caption = adocap.Recordset.Fields!cap
    Else: Label22.Caption = 0
    End If
    adocap.Recordset.Close
    
    'calculating bike capital
    adocap.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adocap.CommandType = adCmdText
    adocap.RecordSource = "SELECT SUM(BIKE_LEVEL*UNIT_PRICE) AS cap FROM stock_assembled"
    adocap.Refresh
    If adocap.Recordset.Fields!cap <> 0 Then
        Label25.Caption = adocap.Recordset.Fields!cap
    Else: Label25.Caption = 0
    End If
    adocap.Recordset.Close
    
    'calculating spares capital
    adocap.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adocap.CommandType = adCmdText
    adocap.RecordSource = "SELECT SUM(STOCK_BALANCE*UNIT_PRICE) AS cap FROM stock_spares"
    adocap.Refresh
    If adocap.Recordset.Fields!cap <> 0 Then
        Label24.Caption = adocap.Recordset.Fields!cap
    Else: Label24.Caption = 0
    End If
    adocap.Recordset.Close
    
    'calculating debtor capital
    adocap.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adocap.CommandType = adCmdText
    adocap.RecordSource = "SELECT SUM(BALANCE) AS cap FROM dealer"
    adocap.Refresh
    If adocap.Recordset.Fields!cap <> 0 Then
        Label27.Caption = adocap.Recordset.Fields!cap
    Else: Label27.Caption = 0
    End If
    adocap.Recordset.Close
    
    'transferring value from left to right side
    Label29 = Label14
    Label31 = Text1
    'calculating gross profit
    Label18.Caption = Val(label10.Caption) + Val(Label13.Caption) + Val(Label12.Caption) + Val(Label11.Caption)
    'calcualting net profit
    Label14.Caption = Val(Label18.Caption) - Val(Text1.Text)
    'calcualting final capital
    Label33 = Val(Label22) + Val(Label25) + Val(Label24) + Val(Label27) + Val(Label29) - Val(Label31)
   
errorhandler:
If Err.Number <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub Command2_Click()
    label10 = "0"
    Label11 = "0"
    Label12 = "0"
    Label13 = "0"
    Label18 = "0"
    Label14 = "0"
    Text1 = ""
    Label22 = "0"
    Label25 = "0"
    Label24 = "0"
    Label27 = "0"
    Label29 = "0"
    Label31 = "0"
    Label33 = "0"
    MsgBox "Data cleared", vbInformation + vbOKOnly, "System Information"
End Sub





Private Sub Form_Load()
    fromdate.Value = Date
    todate.Value = Date
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(Text1.Text) = False Then
        Text1.Text = ""
    End If
End Sub
