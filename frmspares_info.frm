VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmspares_info 
   Caption         =   "Spares Info"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmspares_info.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab framex 
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   20205
      _ExtentX        =   35639
      _ExtentY        =   16325
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
      TabCaption(0)   =   "Inventory Listings"
      TabPicture(0)   =   "frmspares_info.frx":F172
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Adodc2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DataCombo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtsearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Inventory Info "
      TabPicture(1)   =   "frmspares_info.frx":F18E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framexx"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton Command5 
         Caption         =   "FULL INVENTORY REPORT"
         Height          =   495
         Left            =   15960
         TabIndex        =   43
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   15960
         TabIndex        =   21
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SEARCH"
         Height          =   615
         Left            =   15960
         TabIndex        =   20
         Top             =   3240
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmspares_info.frx":F1AA
         Height          =   315
         Left            =   15960
         TabIndex        =   18
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "BID"
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
      Begin VB.CommandButton Command2 
         Caption         =   "LOW INVENTORY REPORT"
         Height          =   495
         Left            =   15960
         TabIndex        =   11
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "REFRESH"
         Height          =   495
         Left            =   15960
         TabIndex        =   10
         Top             =   720
         Width           =   2655
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   615
         Left            =   240
         Top             =   8280
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1085
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
         Height          =   9255
         Left            =   -75000
         TabIndex        =   2
         Top             =   360
         Width           =   19215
         Begin VB.Frame Frame1 
            Caption         =   "Spares Information"
            Height          =   2415
            Left            =   240
            TabIndex        =   22
            Top             =   120
            Width           =   15615
            Begin VB.TextBox txtcp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "COST_PRICE"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   10560
               TabIndex        =   42
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtup 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "UNIT_PRICE"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   7800
               TabIndex        =   31
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtdate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "LAST _ADDED_DATE"
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
               Left            =   1800
               TabIndex        =   30
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txtquan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "LAST_ADDED_QUANTITY"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   1800
               TabIndex        =   29
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtslvl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "STOCK_BALANCE"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   13920
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtname 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "STOCK_DETAILS"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   1800
               TabIndex        =   27
               Top             =   840
               Width           =   3975
            End
            Begin VB.TextBox txtbid 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "BID"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   7800
               TabIndex        =   26
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtsid 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "SID"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   1800
               TabIndex        =   25
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtclvl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "STOCK_CLEVEL"
               DataSource      =   "adospares_info"
               Height          =   375
               Left            =   7800
               TabIndex        =   24
               Top             =   840
               Width           =   1335
            End
            Begin MSAdodcLib.Adodc Adodc1 
               Height          =   375
               Left            =   3480
               Top             =   360
               Visible         =   0   'False
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   661
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
            Begin MSDataListLib.DataCombo cmbbid 
               Bindings        =   "frmspares_info.frx":F1C7
               Height          =   315
               Left            =   9240
               TabIndex        =   32
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               ListField       =   "BID"
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
            Begin VB.Label Label10 
               BackColor       =   &H80000000&
               Caption         =   "COST Price"
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
               Left            =   9240
               TabIndex        =   41
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label8 
               BackColor       =   &H80000000&
               Caption         =   "UNIT Price"
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
               Left            =   6120
               TabIndex        =   40
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000000&
               Caption         =   "Current Stock Level"
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
               Left            =   12240
               TabIndex        =   39
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "Bike ID"
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
               Left            =   6120
               TabIndex        =   38
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000000&
               Caption         =   "Details / Name"
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
               Left            =   120
               TabIndex        =   37
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000000&
               Caption         =   "Stock ID"
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
               Left            =   120
               TabIndex        =   36
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label12 
               BackColor       =   &H80000000&
               Caption         =   "Stock Added Date"
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
               Left            =   120
               TabIndex        =   35
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H8000000D&
               Caption         =   "Stock Critical Level"
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
               Index           =   1
               Left            =   6120
               TabIndex        =   34
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label ladd 
               BackColor       =   &H8000000D&
               Caption         =   "ADD NEW QUANTITY"
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   1320
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label llast 
               BackColor       =   &H80000000&
               Caption         =   "Last Added Quantity"
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
               Left            =   120
               TabIndex        =   23
               Top             =   1320
               Width           =   1695
            End
         End
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
            Left            =   8160
            TabIndex        =   12
            Top             =   6840
            Width           =   7695
            Begin VB.CommandButton Command3 
               Caption         =   "REPORT"
               Height          =   375
               Left            =   4080
               TabIndex        =   13
               Top             =   360
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker todate 
               Height          =   375
               Left            =   240
               TabIndex        =   14
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
               Format          =   78839811
               CurrentDate     =   40805
            End
            Begin MSComCtl2.DTPicker fromdate 
               Height          =   375
               Left            =   2160
               TabIndex        =   15
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
               Format          =   78839811
               CurrentDate     =   40805
            End
            Begin VB.Label Label7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   17
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   720
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
            Left            =   16320
            TabIndex        =   7
            Top             =   2760
            Width           =   1935
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
            Left            =   16320
            TabIndex        =   6
            Top             =   2160
            Width           =   1935
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
            Left            =   16320
            TabIndex        =   5
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "Update"
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
            Left            =   16320
            TabIndex        =   4
            Top             =   960
            Width           =   1935
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
            Left            =   16320
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc adospares_info 
            Height          =   495
            Left            =   240
            Top             =   7080
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
            RecordSource    =   "stock_spares"
            Caption         =   "Spares Information"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
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
            Bindings        =   "frmspares_info.frx":F1DC
            Height          =   3975
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   18
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
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmspares_info.frx":F1F9
         Height          =   7335
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
      Begin VB.Label Label9 
         BackColor       =   &H80000000&
         Caption         =   "VIEW BY BIKE ID"
         Height          =   255
         Left            =   15960
         TabIndex        =   19
         Top             =   2520
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "SPARES PARTS  INFORMATION"
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
      Width           =   3285
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmspares_info.frx":F20E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmspares_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbbid_Change()
txtbid.Text = cmbbid.Text
End Sub

Private Sub cmdadd_Click()
'from herecmd butn changes
cmbbid.Enabled = True
cmdadd.Enabled = False
adospares_info.Enabled = True
cmdref.Enabled = False
cmdaddstock.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
ladd.Visible = True
'from here txt box locking disabled
cmbbid.Enabled = True
Frame1.Enabled = True
DataGrid1.EditActive = False
adospares_info.Enabled = False
adospares_info.Recordset.AddNew
txtdate = Format(Date, "dd/MM/yyyy")
End Sub

Private Sub cmdaddstock_Click()
stockadd = InputBox("Type in numbers of stock you want to add to the SPARE " & txtname.Text, "Add SPARE Stock", "1000")
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
End Sub

Private Sub cmdcancel_Click()
adospares_info.Recordset.Cancel
cmdadd.Enabled = True
cmdref.Enabled = True
cmdsave.Enabled = False
cmdaddstock.Enabled = True
adospares_info.Refresh
ladd.Visible = False
cmbbid.Enabled = False
cmdcancel.Enabled = False
Frame1.Enabled = False
DataGrid1.EditActive = True
adospares_info.Enabled = True
MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"


End Sub

Private Sub cmdref_Click()
adospares_info.Refresh
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"

End Sub

Private Sub cmdsave_Click()
If txtsid = "" Then
    MsgBox "Please input Stock ID", vbExclamation + vbOKOnly, "System Error"
    txtsid.SetFocus
ElseIf txtbid = "" Then
    MsgBox "Please input BIKE ID", vbExclamation + vbOKOnly, "System Error"
    txtbid.SetFocus
ElseIf txtquan = "" Then
    MsgBox "Please input quantity to be added", vbExclamation + vbOKOnly, "System Error"
    txtquan.SetFocus
ElseIf txtclvl = "" Then
    MsgBox "Please input critical level", vbExclamation + vbOKOnly, "System Error"
    txtclvl.SetFocus
ElseIf txtup = "" Then
    MsgBox "Please input unit price", vbExclamation + vbOKOnly, "System Error"
    txtup.SetFocus
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
cmbbid.Enabled = False
cmdcancel.Enabled = False
Frame1.Enabled = False
DataGrid1.EditActive = True
adospares_info.Enabled = True
End If
erroronproduction:
    If Err.Number = -2147217842 Then
        MsgBox "Operation is cancelled because you have already used this ID. Try again choosing different ID. ", vbCritical + vbOKOnly, "Database Error"
End If
End Sub

Private Sub DTPicker1_Change()
txtdate.Text = DTPicker1.Value
End Sub

Private Sub Command1_Click()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = " SELECT *FROM stock_spares "
Adodc2.Refresh
DataGrid2.BackColor = &H80000005
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"

End Sub

Private Sub Command2_Click()
If dev_bike.rscmdlow_stck.State = adStateOpen Then
        dev_bike.rscmdlow_stck.Close
        
    End If
        rptlow_spares.Show
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

Private Sub Command4_Click()
If txtsearch = "" Then
    MsgBox "Please include a search term and try again.", vbExclamation, "Borac Sales System"
Else
    DataGrid1.BackColor = vbGreen
    If Command4.Caption = "Clear Search" Then
        Adodc2.CommandType = adCmdText
        Adodc2.RecordSource = "SELECT *FROM stock_spares c"
        Adodc2.Refresh
        Command4.Caption = "SEARCH"
        txtsearch = ""
        DataGrid1.BackColor = &H80000005
    ElseIf Command4.Caption = "SEARCH" Then
        Adodc2.CommandType = adCmdText
        Adodc2.RecordSource = "SELECT *FROM stock_spares WHERE (((stock_spares.SID)  Like '" & Me.txtsearch.Text & "%'))"
        Adodc2.Refresh
        DataGrid2.BackColor = vbGreen
        If Adodc2.Recordset.RecordCount = 0 Then
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "SELECT *FROM stock_spares WHERE (((stock_spares.STOCK_DETAILS)  Like '" & Me.txtsearch.Text & "%'))"
            Adodc2.Refresh
            DataGrid2.BackColor = vbGreen
        End If
        If Adodc2.Recordset.RecordCount = 0 Then
            MsgBox "No records matched.", vbOKOnly + vbInformation, "System Information"
        Else
            MsgBox Adodc2.Recordset.RecordCount & " " & "records found.", vbOKOnly + vbInformation, "System Information"
            DataGrid2.BackColor = vbGreen
        End If
    Command4.Caption = "Clear Search"
    DataGrid2.BackColor = &H80000005
End If
End If
End Sub

Private Sub Command5_Click()
rptallinventory.Show
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If DataCombo1.Text <> "" Then
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = " SELECT *FROM stock_spares where (((stock_spares.BID) LIKE '" & Me.DataCombo1.Text & "%'))"
    Adodc2.Refresh
    DataGrid2.BackColor = vbGreen
End If
End Sub

Private Sub Form_Load()
Me.Caption = "Spares Info"
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmspares_info.Width - framex.Width) / 2
framex.Left = fpostn
todate = Date
fromdate = Date
DataGrid1.Columns(8).Visible = False
End Sub

Private Sub Form_Resize()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmspares_info.Width - framex.Width) / 2
framex.Left = fpostn
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

Private Sub txtquan_Change()
If IsNumeric(txtquan) = False Then
        txtquan.Text = ""
End If
End Sub

Private Sub txtup_Change()
If IsNumeric(txtup) = False Then
        txtup.Text = ""
End If
End Sub
