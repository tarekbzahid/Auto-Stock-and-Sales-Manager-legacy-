VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmspares_sales 
   Caption         =   " Spares Sales"
   ClientHeight    =   9030
   ClientLeft      =   5565
   ClientTop       =   1725
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdelivery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmdelivery.frx":F172
   ScaleHeight     =   9030
   ScaleWidth      =   10845
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmdelivery.frx":FE3C
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10560
      Top             =   240
   End
   Begin VB.TextBox txtqty_chk 
      DataField       =   "STOCK_BALANCE"
      DataSource      =   "adospare"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoinv 
      Height          =   330
      Left            =   12600
      Top             =   0
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
      BOFAction       =   1
      EOFAction       =   1
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
      RecordSource    =   "invoice_spares"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adotemp 
      Height          =   330
      Left            =   11400
      Top             =   0
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
      RecordSource    =   "temp_inv"
      Caption         =   "adotemp_inv"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoinv_detils 
      Height          =   330
      Left            =   10200
      Top             =   0
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
      RecordSource    =   "invoice_details"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adospare 
      Height          =   330
      Left            =   13800
      Top             =   0
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
      RecordSource    =   "stock_spares"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adodelr 
      Height          =   330
      Left            =   7200
      Top             =   0
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
      RecordSource    =   "dealer"
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
   Begin TabDlg.SSTab framex 
      Height          =   8895
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   15690
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sales"
      TabPicture(0)   =   "frmdelivery.frx":FE58
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "framexx"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmdelivery.frx":FE74
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame9 
         Caption         =   "SPARES PARTS TRANSACTION BETWEEN DATES"
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
         Left            =   -74760
         TabIndex        =   66
         Top             =   480
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   67
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   6094851
            CurrentDate     =   40805
         End
         Begin MSComCtl2.DTPicker fromdate 
            Height          =   375
            Left            =   2160
            TabIndex        =   69
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   6094851
            CurrentDate     =   40805
         End
         Begin VB.Label Label8 
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   2160
            TabIndex        =   70
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame framexx 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   7335
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   18735
         Begin VB.Frame Frame1 
            Caption         =   "Sold To"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   3495
            Begin VB.OptionButton Option1 
               BackColor       =   &H80000003&
               Caption         =   "Retail"
               Height          =   375
               Left            =   120
               TabIndex        =   65
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               DataField       =   "BUSINESS_ADDRESS"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   2400
               Width           =   3255
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               DataField       =   "TOTAL_DUE"
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
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   4080
               Width           =   1815
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               DataField       =   "TOTAL_PAID"
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
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   3600
               Width           =   1815
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
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
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   3120
               Width           =   1815
            End
            Begin VB.TextBox txtuser 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1680
               Width           =   3255
            End
            Begin VB.OptionButton optothers 
               BackColor       =   &H80000003&
               Caption         =   "Others"
               Height          =   375
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   1215
            End
            Begin VB.OptionButton optdealer 
               BackColor       =   &H80000003&
               Caption         =   "Dealer"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo dcmbother 
               Bindings        =   "frmdelivery.frx":FE90
               Height          =   315
               Left            =   1320
               TabIndex        =   50
               Top             =   720
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               ListField       =   "OTHER_NAME"
               Text            =   ""
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
            Begin MSDataListLib.DataCombo dcmbdlr 
               Bindings        =   "frmdelivery.frx":FEA8
               Height          =   315
               Left            =   1320
               TabIndex        =   51
               Top             =   240
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   "DEALER_NAME"
               Text            =   ""
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
            Begin VB.Label Label4 
               BackColor       =   &H80000003&
               Caption         =   "BUSINESS NAME"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   55
               Top             =   2160
               Width           =   3255
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000003&
               Caption         =   "Balance / TK"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   54
               Top             =   4080
               Width           =   1455
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000003&
               Caption         =   "Total Paid / TK"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   53
               Top             =   3600
               Width           =   1455
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000003&
               Caption         =   "Total Tran / Tk"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   52
               Top             =   3120
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3720
            TabIndex        =   33
            Top             =   3480
            Width           =   9615
            Begin VB.TextBox txtprts_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "SID"
               DataSource      =   "adospare"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtperts_des 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "STOCK_DETAILS"
               DataSource      =   "adospare"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   720
               Width           =   4575
            End
            Begin VB.TextBox txtqty 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6600
               TabIndex        =   35
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtup 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "UNIT_PRICE"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adospare"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8040
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000C&
               Caption         =   "PARTS ID"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000C&
               Caption         =   "PARTS DESCRIPTION"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   5
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   4575
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000C&
               Caption         =   "Qty"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   6
               Left            =   6600
               TabIndex        =   39
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000C&
               Caption         =   "UNIT PRICE"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   7
               Left            =   8040
               TabIndex        =   38
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.ComboBox cmbmdeofpaymnt 
            Height          =   315
            ItemData        =   "frmdelivery.frx":FEBE
            Left            =   12600
            List            =   "frmdelivery.frx":FECB
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Height          =   615
            Left            =   12600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Text            =   "frmdelivery.frx":FEE8
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Frame Frame2 
            Caption         =   "Transaction data"
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
            Height          =   2415
            Left            =   10080
            TabIndex        =   22
            Top             =   0
            Width           =   2295
            Begin VB.Timer timer_process 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   2280
               Top             =   360
            End
            Begin VB.TextBox txtdue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               DataField       =   "DUE"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   375
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtpaid 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "PAID"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   25
               Text            =   "0"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox txtothr_chrg 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "OTHER_CHARGE"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   24
               Text            =   "0"
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtdiscnt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "DISCOUNT"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   23
               Text            =   "0"
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000D&
               Caption         =   "Due"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   30
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000D&
               Caption         =   "Paid"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   29
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000D&
               Caption         =   "Other charges"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   28
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000D&
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   840
               Width           =   1335
            End
         End
         Begin VB.TextBox txttotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "GRAND_TOTAL"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """BDT""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            DataSource      =   "adoinv"
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
            Left            =   13800
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   3720
            TabIndex        =   20
            Top             =   720
            Width           =   6255
            Begin MSDataListLib.DataList DataList1 
               Bindings        =   "frmdelivery.frx":FF2A
               Height          =   2100
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3704
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               ListField       =   "STOCK_DETAILS"
               BoundColumn     =   "STOCK_DETAILS"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.TextBox txtsub_total 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14040
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   6960
            Width           =   1695
         End
         Begin VB.CommandButton cmdrmvefrmlist 
            Caption         =   "Remove from the list"
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
            Height          =   495
            Left            =   13440
            TabIndex        =   18
            Top             =   4080
            Width           =   2175
         End
         Begin VB.CommandButton cmdadtolist 
            Caption         =   "Add to list"
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
            Height          =   495
            Left            =   13440
            TabIndex        =   17
            Top             =   3480
            Width           =   2175
         End
         Begin VB.TextBox txtinv_no 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "STID"
            DataSource      =   "adoinv"
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
            Left            =   13800
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtinv_date 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "T_DATE"
            DataSource      =   "adoinv"
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
            Left            =   13800
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            Caption         =   "Invoicing"
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   15960
            TabIndex        =   12
            Top             =   4680
            Width           =   2655
            Begin VB.CommandButton cmdinv 
               Caption         =   "Invoice"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               MouseIcon       =   "frmdelivery.frx":FF41
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   600
               Width           =   1815
            End
            Begin MSDataListLib.DataCombo dtprm 
               Bindings        =   "frmdelivery.frx":1024B
               Height          =   345
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   609
               _Version        =   393216
               Style           =   2
               ListField       =   "STID"
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Search Product Details"
            Height          =   735
            Left            =   3720
            TabIndex        =   10
            Top             =   0
            Width           =   6255
            Begin MSDataListLib.DataCombo dtcname 
               Bindings        =   "frmdelivery.frx":10260
               DataSource      =   "adospare"
               Height          =   315
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   5895
               _ExtentX        =   10398
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ListField       =   "STOCK_DETAILS"
               Text            =   ""
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
         End
         Begin VB.CommandButton cmdnw_sale 
            Caption         =   "&New Sale"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   16560
            MouseIcon       =   "frmdelivery.frx":10277
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   1815
         End
         Begin VB.CommandButton cmddo 
            Caption         =   "Do Transaction"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   16560
            MouseIcon       =   "frmdelivery.frx":10581
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdcncl 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   16560
            MouseIcon       =   "frmdelivery.frx":1088B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdref_all 
            Caption         =   "Refresh All"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   16560
            MouseIcon       =   "frmdelivery.frx":10B95
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2640
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmdelivery.frx":10E9F
            Height          =   1935
            Left            =   120
            TabIndex        =   56
            Top             =   4800
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   2
            RowHeight       =   18
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
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
         Begin VB.Label Label5 
            BackColor       =   &H80000003&
            Caption         =   " Mode of payment"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12600
            TabIndex        =   63
            Top             =   2520
            Width           =   2895
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000003&
            Caption         =   " TERMS AND CONDITION"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   12600
            TabIndex        =   62
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   " TOTAL"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   12600
            TabIndex        =   61
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            Caption         =   " SUB TOTAL"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   12720
            TabIndex        =   60
            Top             =   6960
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Transac No #"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   12600
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   " Invoice Date"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   12600
            TabIndex        =   58
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10200
            TabIndex        =   57
            Top             =   6960
            Width           =   2295
         End
      End
   End
   Begin MSAdodcLib.Adodc adoothers 
      Height          =   330
      Left            =   6000
      Top             =   0
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
      RecordSource    =   "others"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "SPARES PARTS  SALES"
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2355
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmdelivery.frx":10EB5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmspares_sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim process As String
Dim dname As String
Dim did As String



Private Sub cmdadtolist_Click()
 On Error GoTo HandleAddDataErrors
     'MsgBox "check"
 Frame2.Enabled = True
    'If Val(txtqty.Text) = 0 Then
        'txtqty.Text = ""
    'End If
    If txtqty.Text = "" Or txtqty.Text = 0 Then
        MsgBox "Please type in the product quantity", vbExclamation + vbOKOnly, "Transaction Error"
    Else: Dim pcost As Currency
        Dim a As Long
        Dim b As Currency
        Dim c As Long
        Dim d As Long
        a = Val(txtqty.Text)
        b = Val(txtup.Text)
        d = txtqty_chk.Text
        c = d - a
        pcost = a * b
        adotemp.Recordset.AddNew
        adotemp.Recordset.Fields(0) = txtprts_id
        adotemp.Recordset.Fields(1) = txtperts_des
        adotemp.Recordset.Fields(3) = txtqty
        adotemp.Recordset.Fields(2) = txtup
        adotemp.Recordset.Fields(4) = pcost
        adotemp.Recordset.Save
        adotemp.Refresh
        DataGrid1.Refresh
        adospare.Recordset.Fields(3) = c
        adospare.Recordset.Save
        adospare.Refresh
        txtprts_id = ""
        txtperts_des = ""
        txtqty = ""
        txtup = ""
        'cmdref_all_Click
        adospare.Refresh
        adoinv_detils.Refresh
    adotemp.Refresh
    DataGrid1.Columns(1).Width = 5894.929
        Timer1.Enabled = True
        timer_process.Enabled = True
        DataGrid1.Columns(1).Width = 5894.929
        Label7.Caption = adotemp.Recordset.RecordCount & " " & "Items on the list"
     'MsgBox "check"
    End If
HandleAddDataErrors:
    If Err.Number = -2147217842 Then
        MsgBox "Operation is cancelled because you are trying to enter same product twice.If you want to change the quantity try removing the product first then input the desired quantity. ", vbCritical + vbOKOnly, "Database Error"
        adospare.Refresh
        adodelr.Refresh
        adoinv_detils.Refresh
        adotemp.Refresh
        DataGrid1.Refresh
    End If
    Timer1.Enabled = True
    DataGrid1.Columns(1).Width = 5894.929
End Sub

Private Sub cmdcncl_Click()
'On Error Resume Next
Dim Counter As Integer
Counter = 0
    timer_process.Enabled = False
    Frame5.Enabled = True
    adoinv.Recordset.Cancel
    adoinv.Refresh
    cmddo.Enabled = False
    cmdcncl.Enabled = False
    cmdnw_sale.Enabled = True
    Frame2.Enabled = False
    dcmbdlr.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Frame4.Enabled = False
    cmdref_all.Enabled = True
    cmdadtolist.Enabled = False
    cmdrmvefrmlist.Enabled = False
    process = 0
    Do While adotemp.Recordset.RecordCount <> 0
     adotemp.Recordset.MoveFirst
     adospare.Recordset.MoveFirst
    Do While adospare.Recordset.Fields(0).Value <> adotemp.Recordset.Fields(0).Value
        adospare.Recordset.MoveNext
    Loop
        adospare.Recordset.Fields(3).Value = adospare.Recordset.Fields(3).Value + adotemp.Recordset.Fields(3).Value
        adospare.Recordset.Save
        adotemp.Recordset.Delete
        adotemp.Recordset.Update
        adotemp.Refresh
        adospare.Refresh
        DataGrid1.Refresh
        Timer1.Enabled = True
        'cmdref_all_Click
        adospare.Refresh
    adodelr.Refresh
    adoinv_detils.Refresh
    adotemp.Refresh
    DataGrid1.Columns(1).Width = 5894.929
        txtothr_chrg = 0
        txtdiscnt = 0
        txtpaid = 0
        txtdue = 0
        
    Loop
MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"

End Sub

Private Sub cmddo_Click()

If txtuser.Text = "" Then
    MsgBox "Please input dealer", vbExclamation + vbOKOnly, "System Error"
    txtuser.SetFocus
ElseIf adotemp.Recordset.RecordCount = 0 Then
    MsgBox "No product is in the list !", vbExclamation + vbOKOnly, "System Error"
    DataList1.SetFocus
ElseIf txtsub_total.Text = "" Or txtothr_chrg.Text = "" Or txtpaid.Text = "" Then
    MsgBox "Fill the Transaction Info", vbExclamation + vbOKOnly, "System Error"
ElseIf cmbmdeofpaymnt.Text = "" Then
    MsgBox "Please select the payment mode", vbExclamation + vbOKOnly, "System Error"
    cmbmdeofpaymnt.SetFocus
Else
MsgBox "Please be patient while the processing is done. It will take a while depending on the volume of data and your processor speed.", vbInformation + vbOKOnly, "System Information"
Dim progcount As Long
progcount = 1 + adotemp.Recordset.RecordCount + adotemp.Recordset.RecordCount
'MsgBox progcount
progcount = 100 / progcount
'MsgBox progcount
Dim tid
Dim gtotal As Long
gtotal = txttotal
    tid = txtinv_no
    adotemp.Recordset.MoveFirst
    adoinv.Recordset.Fields(2) = did
    adoinv.Recordset.Fields(3) = dname
    adoinv.Recordset.Fields(4) = Text1
    adoinv.Recordset.Fields(5) = cmbmdeofpaymnt.Text
    adoinv.Recordset.Fields(6) = txtsub_total.Text
    adoinv.Recordset.Fields(7) = txtothr_chrg.Text
    adoinv.Recordset.Fields(9) = txtpaid.Text
    adoinv.Recordset.Fields(10) = txtdue.Text
    adoinv.Recordset.Fields(11) = txtdiscnt.Text
    adoinv.Recordset.Fields(12) = frmmain.stsbr_main.Panels(2).Text
    adoinv.Recordset.Save
    adoinv.Refresh
    adotemp.Refresh
    Counter = 0
    stp = 1
    For Counter = 1 To adotemp.Recordset.RecordCount
        'MsgBox "Counter" & Counter
        'MsgBox "adotemp rec" & adotemp.Recordset.RecordCount
        'MsgBox oop
        'If Not stp = 0 Then
        adoinv_detils.Recordset.AddNew
        adoinv_detils.Recordset.Fields(0).Value = tid
        adoinv_detils.Recordset.Fields(1).Value = adotemp.Recordset.Fields(1).Value
        adoinv_detils.Recordset.Fields(2).Value = adotemp.Recordset.Fields(0).Value
        adoinv_detils.Recordset.Fields(3).Value = adotemp.Recordset.Fields(3).Value
        adoinv_detils.Recordset.Fields(4).Value = adotemp.Recordset.Fields(2).Value
        adoinv_detils.Recordset.Fields(5).Value = adotemp.Recordset.Fields(4).Value
       'ProgressBar1.Value = ProgressBar1.Value + progcount
        adoinv_detils.Recordset.Save
        'End If
        adoinv.Refresh
        'If Not adotemp.Recordset.EOF Then
        adotemp.Recordset.MoveNext
        'Else: stp = 0
        'End If
        'MsgBox "counter" & Counter
    Next Counter
        Counter = 0
        adotemp.Refresh
        For Counter = 0 To adotemp.Recordset.RecordCount
            adotemp.Refresh
            If Not adotemp.Recordset.RecordCount = 0 Then
            adotemp.Recordset.Delete
            End If
            'ProgressBar1.Value = ProgressBar1.Value + progcount
            adotemp.Refresh
        Next Counter
            adotemp.Refresh
            process = 0

    If optdealer.Value = True Then
        adodelr.Recordset.Fields(4) = adodelr.Recordset.Fields(4).Value + gtotal
        adodelr.Recordset.Fields(5) = adodelr.Recordset.Fields(5).Value + txtpaid
        adodelr.Recordset.Fields(6) = adodelr.Recordset.Fields(6).Value + txtdue
        'adodelr.Recordset.Fields(7) = adodelr.Recordset.Fields(7).Value + txtdiscnt
        adodelr.Recordset.Update
        adodelr.Refresh
    End If
    
        'If ProgressBar1.Value <> 0 Then
            'ProgressBar1.Value = 100
        'End If
            'ProgressBar1.Visible = False
            txtothr_chrg = 0
            txtdiscnt = 0
            txtpaid = 0
            txtdue = 0
            cmdnw_sale.Enabled = True
            cmddo.Enabled = False
            cmdref_all.Enabled = True
            cmdcncl.Enabled = False
            Frame5.Enabled = True
            MsgBox "Transaction completed", vbInformation + vbOKOnly, "System Information"
            txtsub_total = ""
            Label7 = ""
              'invoice generation
          If dev_bike.rscmdinv.State = adStateOpen Then
    dev_bike.rscmdinv.Close
End If
    dev_bike.cmdinv Trim(tid)
    rptinv_spares_off.Show
End If

End Sub

Private Sub cmdgat_pas_Click()
If dev_bike.rscmdinv.State = adStateOpen Then
        dev_bike.rscmdinv.Close
    End If
    dev_bike.cmdgatepass Trim(dtprm.Text)
    rptgatpass.Show
End Sub

Private Sub cmdinv_Click()
If dev_bike.rscmdinv.State = adStateOpen Then
    dev_bike.rscmdinv.Close
End If
    dev_bike.cmdinv Trim(dtprm.Text)
    rptinv_spares_off.Show
End Sub

Private Sub cmdnw_sale_Click()
Frame2.Enabled = True
Frame5.Enabled = False
cmddo.Enabled = True
cmdcncl.Enabled = True
cmdnw_sale.Enabled = False
cmdref_all.Enabled = False
dcmbdlr.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Frame4.Enabled = True
cmdadtolist.Enabled = True
cmdrmvefrmlist.Enabled = True
Dim varTID
Dim process
If adoinv.Recordset.RecordCount = 0 Then
    adoinv.Recordset.AddNew
    txtinv_date.Text = Format(Date, "dd/MM/yyyy")
    txtinv_no = "ST1005001"
    process = 1
Else: adoinv.Recordset.MoveLast
    varTID = Mid(txtinv_no.Text, 3, 9)
    adoinv.Recordset.AddNew
    txtinv_date.Text = Format(Date, "dd/MM/yyyy")
    txtinv_no.Text = "ST" + CStr(varTID + 1)
    process = 1
End If
cmdadtolist.Enabled = True
cmdrmvefrmlist.Enabled = True
End Sub

Private Sub cmdref_all_Click()
    adospare.Refresh
    adodelr.Refresh
    adoinv_detils.Refresh
    adotemp.Refresh
    DataGrid1.Columns(1).Width = 5894.929
   MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"
End Sub

Private Sub cmdrmvefrmlist_Click()
On Error GoTo RmvDataError

ask = MsgBox("Do you want to remove the current product from the list ?", vbQuestion + vbYesNo, "System query")
If ask = vbYes Then
    adospare.Recordset.MoveFirst
    Do While adospare.Recordset.Fields(0).Value <> adotemp.Recordset.Fields(0).Value
        adospare.Recordset.MoveNext
    Loop
        adospare.Recordset.Fields(3).Value = adospare.Recordset.Fields(3).Value + adotemp.Recordset.Fields(3).Value
        adospare.Recordset.Save
        adotemp.Recordset.Delete
        adotemp.Recordset.Update
        adotemp.Refresh
        adospare.Refresh
        DataGrid1.Refresh
        Timer1.Enabled = True
        'cmdref_all_Click
        adospare.Refresh
        adoinv_detils.Refresh
        adotemp.Refresh
        DataGrid1.Columns(1).Width = 5894.929
        DataGrid1.Columns(1).Width = 5894.929
        Label7.Caption = adotemp.Recordset.RecordCount & " " & "Items on the list"
End If
If adotemp.Recordset.RecordCount = o Then
Frame2.Enabled = False
End If
RmvDataError:
    If Err.Number = 3021 Then
        MsgBox "Select the product you want to remove from the list", vbInformation + vbOKOnly, "Data remove error"
    End If
    DataGrid1.Columns(1).Width = 5894.929
End Sub







Private Sub DataCombo2_Change()
    adospare.Recordset.Bookmark = DataCombo2.SelectedItem
End Sub



Private Sub Command3_Click()
Dim prmfrom
Dim prmto
prmfrom = fromdate.Value
prmto = todate.Value
If dev_bike.rscmdSPARETRANBETWEENDATES.State = adStateOpen Then
  dev_bike.rscmdSPARETRANBETWEENDATES.Close
End If
 dev_bike.cmdSPARETRANBETWEENDATES prmfrom, prmto
rptsparetranbetwdates.Show
End Sub
Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Cancel = 1
    'MsgBox DataGrid1.Columns(1).Width
End Sub

Private Sub DataList1_Click()
On Error Resume Next
    adospare.Recordset.Bookmark = DataList1.SelectedItem
    txtqty.Text = ""
    
End Sub

Private Sub DataList1_DblClick()
    adospare.Recordset.Bookmark = DataList1.SelectedItem
    
End Sub

Private Sub DBCombo1_Change()
    If Not DataCombo1 = Null Then
        adodelr.Recordset.Bookmark = DataCombo1.SelectedItem
    End If
End Sub

Private Sub dcmbdlr_Change()
    'txtdelr.Text = dcmbdlr.Text
    On Error Resume Next
    Text5 = ""
    Text6 = ""
    Text7 = ""
     adodelr.Recordset.Bookmark = dcmbdlr.SelectedItem
    did = adodelr.Recordset.Fields(0).Value
    dname = adodelr.Recordset.Fields(1).Value
    txtuser = dname
    Text4 = adodelr.Recordset.Fields(14).Value
    Text5 = adodelr.Recordset.Fields(4).Value
    Text6 = adodelr.Recordset.Fields(5).Value
    Text7 = adodelr.Recordset.Fields(6).Value
End Sub

Private Sub dcmbdlr_KeyUp(KeyCode As Integer, Shift As Integer)
    'chkother.Value = 1
    'chkdealer.Value = 0
End Sub



Private Sub dcmbother_Change()
adoothers.Recordset.Bookmark = dcmbother.SelectedItem
    did = "OTHERS"
    dname = adoothers.Recordset.Fields(0).Value
    txtuser = dname & " - OTHERS"
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
End Sub

Private Sub dtcname_Change()
adospare.Recordset.Bookmark = dtcname.SelectedItem
End Sub

Private Sub dtprm_Change()
adoinv.Recordset.Bookmark = dtprm.SelectedItem
End Sub


Private Sub Form_Load()
    DataGrid1.Columns(1).Width = 5894.929
    Image2.Width = Me.Width
    Dim fpostn As Long
    fpostn = (frmspares_sales.Width - framex.Width) / 2
    framex.Left = postn
   ' DataGrid1.Columns(2).DataFormat = vbCurrency
    'DataGrid1.Columns(4).DataFormat = vbCurrency
    
End Sub

Private Sub Form_Resize()
 Image2.Width = Me.Width
 Dim fpostn As Long
 fpostn = (frmspares_sales.Width - framex.Width) / 2
 framex.Left = fpostn
    'MsgBox fpostn
    'MsgBox frmspares_sales.Width & " " & framex.Width
End Sub

Private Sub Form_Terminate()
Dim process
If process = 1 Then
MsgBox "Sorry you cant quit now. You have process on the queue.", vbExclamation + vbOKCancel, "System Error"
Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim process
If process = 1 Then
MsgBox "Sorry you cant quit now. You have process on the queue.", vbExclamation + vbOKCancel, "System Error"
Cancel = 1
End If
End Sub

Private Sub optdealer_Click()
 If optdealer.Value = True Then
        Option1.Value = False
        txtuser.Locked = True
        dcmbdlr.Enabled = True
        dcmbother.Enabled = False
    Else: dcmbdlr.Enabled = False
    End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
txtuser.Locked = False
txtuser = ""
End If
End Sub

Private Sub Option1_LostFocus()
If Option1.Value = False Then
txtuser.Locked = True
End If
End Sub

Private Sub optothers_Click()
   If optothers.Value = True Then
        dcmbother.Enabled = True
        Option1.Value = False
        txtuser.Locked = True
        dcmbdlr.Enabled = False
    Else: dcmbother.Enabled = False
    End If
End Sub

Private Sub timer_process_Timer()
If IsNumeric(txtothr_chrg) = False Then
        txtothr_chrg.Text = 0
End If
If Val(txtsub_total) <> 0 Then
    txttotal = Val(txtsub_total) - Val(txtdiscnt) + Val(txtothr_chrg)
Else:
End If
If IsNumeric(txtpaid.Text) = False Then
        txtpaid.Text = 0
End If
If Val(txttotal.Text) <> 0 Then
    txtdue = Val(txttotal) - Val(txtpaid)
End If
timer_process.Enabled = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next 'updating subtotal by calculating the all firld value
    Dim sub_total As Currency
    If adotemp.Recordset.BOF = False Then
        adotemp.Recordset.MoveFirst
            Do While adotemp.Recordset.EOF = False
                sub_total = adotemp.Recordset.Fields(4).Value + sub_total
                adotemp.Recordset.MoveNext
            Loop
        txtsub_total = sub_total
    ElseIf adotemp.Recordset.Fields(4).Value = "" Then
        sub_total = "0"
        txtsub_total = sub_total
    End If
    Timer1.Enabled = False
    Label7.Caption = adotemp.Recordset.RecordCount & " " & "Items on the list"
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub txtdiscnt_Change()
timer_process.Enabled = True
End Sub

Private Sub txtothr_chrg_Change()
timer_process.Enabled = True
End Sub

Private Sub txtpaid_Change()
timer_process.Enabled = True
End Sub

Private Sub txtqty_Change()
    If IsNumeric(txtqty) = False Then
        txtqty.Text = 0
    ElseIf txtqty.Text = "." Then
        txtqty.Text = ""
    Else: Dim a As Long
        Dim b As Long
        a = txtqty.Text
        b = txtqty_chk.Text
            If a > b Then
                MsgBox "Sorry you only have" & " " & txtqty_chk.Text & " " & "Spares left", vbOKOnly + vbInformation, "BORAC bike selling company"
                txtqty.Text = ""
            End If
    End If
End Sub

Private Sub txtsub_total_Change()
timer_process.Enabled = True
End Sub



Private Sub txtuser_KeyPress(KeyAscii As Integer)
did = "RETAIL CUSTOMER"
dname = txtuser.Text
If IsNumeric(txtuser.Text) = True Then
txtuser.Text = ""
End If
End Sub

Private Sub txtuser_KeyUp(KeyCode As Integer, Shift As Integer)
If IsNumeric(txtuser.Text) = True Then
txtuser.Text = ""
End If
End Sub
