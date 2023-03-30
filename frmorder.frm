VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmorder 
   Appearance      =   0  'Flat
   Caption         =   "ORDERS"
   ClientHeight    =   10185
   ClientLeft      =   4410
   ClientTop       =   2445
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmorder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmorder.frx":F172
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmorder.frx":FE3C
      Height          =   135
      Left            =   5760
      TabIndex        =   101
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc adodealer 
      Height          =   330
      Left            =   11640
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
      RecordSource    =   "dealer"
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
   Begin VB.TextBox txttempOID 
      DataField       =   "OID"
      DataSource      =   "adoorderlist"
      Height          =   195
      Left            =   5160
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox Text2 
      DataField       =   "DID"
      DataSource      =   "adoorderlist"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSAdodcLib.Adodc adoorderdetails 
      Height          =   330
      Left            =   10440
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
      RecordSource    =   "dealer_order_details"
      Caption         =   "Adodc2"
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
   Begin VB.Timer tmrtotal 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   120
   End
   Begin VB.TextBox txtbdet 
      DataField       =   "BIKE_DETAILS"
      DataSource      =   "adostock_ckd"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txtbid 
      DataField       =   "BID"
      DataSource      =   "adostock_ckd"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSAdodcLib.Adodc adostock_ckd 
      Height          =   330
      Left            =   8040
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
      RecordSource    =   "stock_ckd"
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
   Begin VB.TextBox txtup 
      DataField       =   "UNIT_PRICE"
      DataSource      =   "adostock_ckd"
      Height          =   375
      Left            =   6675
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSAdodcLib.Adodc adoorder 
      Height          =   330
      Left            =   9240
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
      RecordSource    =   "dealer_order"
      Caption         =   "Adodc5"
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
   Begin TabDlg.SSTab framex 
      Height          =   9375
      Left            =   -120
      TabIndex        =   1
      Top             =   720
      Width           =   20295
      _ExtentX        =   35798
      _ExtentY        =   16536
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
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
      TabCaption(0)   =   "ORDER LIST"
      TabPicture(0)   =   "frmorder.frx":FE5A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "optother"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optdealer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optretail"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "cmddelorder"
      Tab(0).Control(5)=   "DataGrid3"
      Tab(0).Control(6)=   "adostock_assembled"
      Tab(0).Control(7)=   "temporderdetails"
      Tab(0).Control(8)=   "Command4"
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(11)=   "adoorderlist"
      Tab(0).Control(12)=   "cmddone"
      Tab(0).Control(13)=   "cmdordr_his"
      Tab(0).Control(14)=   "Label29"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "NEW ORDER"
      TabPicture(1)   =   "frmorder.frx":FE76
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(14)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(15)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(17)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DataGrid1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DataList1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtitem"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdcncl"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdnw_ordr"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtqty"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdadd"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "DataCombo2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command6"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "SEARCH"
      TabPicture(2)   =   "frmorder.frx":FE92
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "ORDER DELIVERY"
      TabPicture(3)   =   "frmorder.frx":FEAE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "cmdcheck"
      Tab(3).Control(3)=   "Command7"
      Tab(3).Control(4)=   "cmdref"
      Tab(3).Control(5)=   "DataGrid5"
      Tab(3).Control(6)=   "adochckbike"
      Tab(3).Control(7)=   "adoorder_details"
      Tab(3).Control(8)=   "Adodc1"
      Tab(3).Control(9)=   "DataGrid6"
      Tab(3).ControlCount=   10
      Begin VB.Frame Frame7 
         Caption         =   "Pending Orders "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74760
         TabIndex        =   107
         Top             =   600
         Width           =   6375
         Begin VB.TextBox txtidchck 
            Appearance      =   0  'Flat
            DataField       =   "OID"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            DataField       =   "DID"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            DataField       =   "DEALER_NAME"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            DataField       =   "PARTICULARS"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   116
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            DataField       =   "ORDER_DATE"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   115
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            DataField       =   "DELIVERY_DATE"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            DataField       =   "PAYMENT_MODE"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   113
            Top             =   3240
            Width           =   3015
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            DataField       =   "BANK_DETAIL"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   3720
            Width           =   3015
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            DataField       =   "GRAND TOTAL"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   4200
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            DataField       =   "PAID"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   4680
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            DataField       =   "DUE"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   109
            Top             =   5160
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "frmorder.frx":FECA
            Height          =   315
            Left            =   3840
            TabIndex        =   108
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "OID"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label40 
            BackColor       =   &H80000003&
            Caption         =   "Order ID"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label39 
            BackColor       =   &H80000003&
            Caption         =   "Dealer ID"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   129
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label38 
            BackColor       =   &H80000003&
            Caption         =   "Dealer Name"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   128
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label37 
            BackColor       =   &H80000003&
            Caption         =   "Particulars"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   127
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label36 
            BackColor       =   &H80000003&
            Caption         =   "Order Date"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   126
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label35 
            BackColor       =   &H80000003&
            Caption         =   "Delivery Date"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   125
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label34 
            BackColor       =   &H80000003&
            Caption         =   "Due "
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   124
            Top             =   5160
            Width           =   1815
         End
         Begin VB.Label Label33 
            BackColor       =   &H80000003&
            Caption         =   "Paid"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   123
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label32 
            BackColor       =   &H80000003&
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   122
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackColor       =   &H80000003&
            Caption         =   "Bank  (if neccessary)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   121
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label30 
            BackColor       =   &H80000003&
            Caption         =   "Payment Mode "
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Top             =   3240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Order Particulars"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -68280
         TabIndex        =   105
         Top             =   600
         Width           =   10935
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "frmorder.frx":FEDF
            Height          =   3855
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
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
      End
      Begin VB.CommandButton cmdcheck 
         Caption         =   "Check stock availability"
         Height          =   375
         Left            =   -59520
         TabIndex        =   104
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Deliver order"
         Height          =   495
         Left            =   -60720
         TabIndex        =   103
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   -58920
         TabIndex        =   102
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "INFORMATION"
         Height          =   3375
         Left            =   13080
         TabIndex        =   88
         Top             =   480
         Width           =   6495
         Begin VB.TextBox txtaddcosts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """BDT""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   94
            Text            =   "00"
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox txtacost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """BDT""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   93
            Text            =   "500"
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtdue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   92
            Text            =   "00"
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtadvance 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   91
            Text            =   "00"
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txttotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   2760
            Width           =   2535
         End
         Begin VB.TextBox txtsubtotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Additional Costs"
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
            Left            =   240
            TabIndex        =   100
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Assembled Cost Per BIKE"
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
            Left            =   240
            TabIndex        =   99
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Due"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   12
            Left            =   240
            TabIndex        =   98
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Paid in Advance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   97
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   " TOTAL"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   96
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   " SUB TOTAL"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   16
            Left            =   240
            TabIndex        =   95
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "SHOW"
         Height          =   375
         Left            =   8280
         TabIndex        =   87
         Top             =   7200
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmorder.frx":FEFE
         Height          =   315
         Left            =   10320
         TabIndex        =   86
         Top             =   6840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ListField       =   "OID"
         Text            =   "DataCombo2"
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&ADD TO CART"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   84
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   83
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Height          =   4335
         Left            =   13080
         TabIndex        =   61
         Top             =   3960
         Width           =   6495
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "frmorder.frx":FF15
            Left            =   1800
            List            =   "frmorder.frx":FF1F
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmorder.frx":FF30
            Left            =   4440
            List            =   "frmorder.frx":FF3D
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1680
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtdlvry_date 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "DELIVERY_DATE"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txtdlvry_stts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "STATUS"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtprtclrs 
            Appearance      =   0  'Flat
            DataField       =   "PARTICULARS"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            LinkTimeout     =   0
            Locked          =   -1  'True
            MaxLength       =   200
            MousePointer    =   3  'I-Beam
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   3120
            Width           =   4575
         End
         Begin VB.TextBox txtordr_date 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "ORDER_DATE"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtdelr_id 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "DID"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtordr_id 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtdelr_name 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "DEALER_NAME"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1200
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "frmorder.frx":FF5F
            Height          =   315
            Left            =   4440
            TabIndex        =   62
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DEALER_NAME"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmorder.frx":FF77
            Height          =   315
            Left            =   4440
            TabIndex        =   71
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DID"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   4440
            TabIndex        =   72
            Top             =   2640
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   79298563
            CurrentDate     =   40949
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Payment Mode"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   13
            Left            =   240
            TabIndex        =   81
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Delivery Status"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   79
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Particulars / Information"
            ForeColor       =   &H80000007&
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   78
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Order placing date"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   77
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Delivery date"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   76
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Name"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   75
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Dealer ID"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   74
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Order ID"
            ForeColor       =   &H80000007&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   73
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdnw_ordr 
         Appearance      =   0  'Flat
         Caption         =   "PLACE ORDER"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdcncl 
         Caption         =   "CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   6720
         Width           =   1455
      End
      Begin VB.TextBox txtitem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "BID"
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "UNIT_PRICE"
         Height          =   375
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   52
         Top             =   5760
         Width           =   1695
      End
      Begin VB.OptionButton optother 
         BackColor       =   &H80000000&
         Caption         =   "PREPARED"
         Height          =   375
         Left            =   -70320
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optdealer 
         BackColor       =   &H80000000&
         Caption         =   "PENDING"
         Height          =   375
         Left            =   -71640
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optretail 
         BackColor       =   &H80000000&
         Caption         =   "DELIVERED"
         Height          =   375
         Left            =   -68880
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "INFORMATION"
         Height          =   7335
         Left            =   -60120
         TabIndex        =   23
         Top             =   480
         Width           =   5055
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "PREPAIRED_BY"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   47
            Top             =   5640
            Width           =   3255
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000000&
            Caption         =   "Prepared by"
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
            Left            =   240
            TabIndex        =   46
            Top             =   5640
            Width           =   1455
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "PAID"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   45
            Top             =   4680
            Width           =   3255
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "DUE"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   44
            Top             =   5160
            Width           =   3255
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000000&
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
            Left            =   240
            TabIndex        =   43
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000000&
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
            Left            =   240
            TabIndex        =   42
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "PAYMENT_MODE"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   41
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "GRAND TOTAL"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   40
            Top             =   4200
            Width           =   3255
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000000&
            Caption         =   "Grand Total"
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
            Left            =   240
            TabIndex        =   39
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000000&
            Caption         =   "Payment mode"
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
            Left            =   240
            TabIndex        =   38
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "DELIVERY_DATE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   37
            Top             =   3240
            Width           =   3255
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000000&
            Caption         =   "Delivery Date"
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
            Left            =   240
            TabIndex        =   36
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "ORDER_DATE"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   " dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   35
            Top             =   2760
            Width           =   3255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "STATUS"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   34
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000000&
            Caption         =   "Order Date"
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
            Left            =   240
            TabIndex        =   33
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "DID"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   32
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000000&
            Caption         =   "Customer ID"
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
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000000&
            Caption         =   "Order Status"
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
            Left            =   240
            TabIndex        =   30
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000000&
            Caption         =   "Order ID"
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
            Left            =   240
            TabIndex        =   29
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "OID"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   28
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "BUSINESS_NAME"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000000&
            Caption         =   "Company Name"
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
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "DEALER_NAME"
            DataSource      =   "adoorderlist"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   25
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000000&
            Caption         =   "Bill To"
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
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmddelorder 
         Caption         =   "Undo Order"
         Height          =   495
         Left            =   -70560
         TabIndex        =   22
         Top             =   7920
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmorder.frx":FF8F
         Height          =   6855
         Left            =   -74640
         TabIndex        =   21
         Top             =   960
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
      Begin MSAdodcLib.Adodc adostock_assembled 
         Height          =   330
         Left            =   -74520
         Top             =   7920
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
         RecordSource    =   "stock_assembled"
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc temporderdetails 
         Height          =   330
         Left            =   -74520
         Top             =   8280
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   ""
         Caption         =   "Adodc2"
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
      Begin VB.CommandButton Command4 
         Caption         =   "Order Details"
         Height          =   495
         Left            =   -65520
         TabIndex        =   19
         Top             =   7920
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "ORDER QUERY CRITERIA"
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
         TabIndex        =   13
         Top             =   1800
         Width           =   7695
         Begin VB.TextBox txtsearch 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            TabIndex        =   16
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "frmorder.frx":FFAA
            Left            =   240
            List            =   "frmorder.frx":FFB4
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Mark as Pending"
         Height          =   495
         Left            =   -57600
         TabIndex        =   8
         Top             =   7920
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   -62880
         TabIndex        =   7
         Top             =   7920
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc adoorderlist 
         Height          =   375
         Left            =   -74520
         Top             =   8640
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ORDERS"
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
      Begin VB.CommandButton cmddone 
         Caption         =   "Mark as Prepared"
         Height          =   495
         Left            =   -60120
         TabIndex        =   3
         Top             =   7920
         Width           =   2415
      End
      Begin VB.CommandButton cmdordr_his 
         Caption         =   "Dealer Order History"
         Height          =   495
         Left            =   -68040
         TabIndex        =   2
         Top             =   7920
         Width           =   2415
      End
      Begin VB.Frame Frame9 
         Caption         =   "ORDER QUERY BETWEEN DATES"
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
         TabIndex        =   9
         Top             =   600
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "d/MM/yyyy"
            Format          =   79298563
            CurrentDate     =   40805
         End
         Begin MSComCtl2.DTPicker fromdate 
            Height          =   375
            Left            =   2160
            TabIndex        =   12
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "d/MM/yyyy"
            Format          =   79298563
            CurrentDate     =   40805
         End
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frmorder.frx":FFC8
         Height          =   2205
         Left            =   240
         TabIndex        =   54
         Top             =   5880
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3889
         _Version        =   393216
         Appearance      =   0
         ListField       =   "BIKE_DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmorder.frx":FFE3
         Height          =   4335
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   7646
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
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmorder.frx":FFFF
         Height          =   255
         Left            =   -69360
         TabIndex        =   131
         Top             =   7320
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc adochckbike 
         Height          =   330
         Left            =   -70680
         Top             =   7320
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
         RecordSource    =   "stock_assembled"
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
      Begin MSAdodcLib.Adodc adoorder_details 
         Height          =   330
         Left            =   -71880
         Top             =   7320
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
         CommandType     =   1
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -74640
         Top             =   7320
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
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
         RecordSource    =   "SELECT dealer_order.* FROM dealer_order WHERE [dealer_order.STATUS]=""PREPARED"""
         Caption         =   " Orders"
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
      Begin MSDataGridLib.DataGrid DataGrid6 
         Bindings        =   "frmorder.frx":10019
         Height          =   255
         Left            =   -68760
         TabIndex        =   132
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   5
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
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "  ORDER RECEIPT"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   975
         Index           =   17
         Left            =   8160
         TabIndex        =   85
         Top             =   6720
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ADD QUANTITY"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   82
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ITEM ID"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   60
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNIT PRICE"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   59
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ITEM DESCRIPTION"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   58
         Top             =   5520
         Width           =   3855
      End
      Begin VB.Label Label29 
         BackColor       =   &H80000000&
         Caption         =   "LOAD ORDER STATUS:"
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
         Left            =   -74640
         TabIndex        =   51
         Top             =   480
         Width           =   14295
      End
   End
   Begin VB.Label Label6 
      Caption         =   "mm/dd/yy"
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERS"
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
      Index           =   8
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Image img_title 
      Height          =   615
      Left            =   0
      Picture         =   "frmorder.frx":10030
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset

Private Sub dtdid_Change()
txtdelr_id = dtdid.Text
End Sub

Private Sub cmdadd_Click()
On Error GoTo errorhandler

If cmdadd.Caption = "ADD TO CART" Then
'adding to the list
     If Val(txtqty.Text) = 0 Then
        MsgBox "Enter a valid quantity", vbExclamation
     Else
        Dim TCP As Double
        Dim COST As Double
            TCP = Val(txtqty) * Val(adostock_ckd.Recordset.Fields(6).Value)
            COST = Val(txtqty) * Val(adostock_ckd.Recordset.Fields(5).Value)
            rsTemp.AddNew
            rsTemp.Fields(0) = adostock_ckd.Recordset.Fields!bid
            rsTemp.Fields(1) = adostock_ckd.Recordset.Fields!BIKE_DETAILS
            rsTemp.Fields(2) = txtqty.Text
            rsTemp.Fields(3) = adostock_ckd.Recordset.Fields!unit_price
            rsTemp.Fields(4) = COST
            rsTemp.Fields(5) = TCP
            rsTemp.Update
            tmrtotal.Enabled = True
    End If

ElseIf cmdadd.Caption = "REMOVE FROM CART" Then
    'removing from the list
    rsTemp.Delete
    tmrtotal.Enabled = True
End If


errorhandler:
If Err <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
     
End Sub

Private Sub cmdcheck_Click()
On Error GoTo errorhandler
Dim Counter As Integer
Counter = 0
Dim stts As Integer
stts = 0
    adoorder_details.Recordset.MoveFirst
    adochckbike.Recordset.MoveFirst
    Do
    Do While adoorder_details.Recordset.Fields(1).Value <> adochckbike.Recordset.Fields(0).Value
        adochckbike.Recordset.MoveNext
        'MsgBox "first search" & " " & adoorder_details.Recordset.Fields(1).Value
    Loop
        Counter = Counter + 1
        'MsgBox "Counter" & Counter
        'MsgBox adochckbike.Recordset.Fields(2).Value & adochckbike.Recordset.Fields(0).Value
        If adochckbike.Recordset.Fields(2).Value < adoorder_details.Recordset.Fields(3).Value Then
        Else: stts = stts + 1
        End If
        'MsgBox "stts" & stts
        If adoorder_details.Recordset.EOF = False Then
        adoorder_details.Recordset.MoveNext
        adochckbike.Recordset.MoveFirst
        End If
    Loop While Counter <> adoorder_details.Recordset.RecordCount
        If stts = adoorder_details.Recordset.RecordCount Then
            MsgBox "You can deliver this order, stock is available.", vbOKOnly + vbInformation, "System Information"
        Else
            MsgBox "Stock is low on one or more items.", vbOKOnly + vbCritical, "System Information"
        End If
errorhandler:
Select Case Err.Number
    Case 3021
        MsgBox "Process is cancelled due to a fatal error. Possible reasons are incomplete records or missing bike ID. ", vbCritical + vbOKOnly, "Database Error"
End Select
End Sub

Private Sub cmdcncl_Click()
    
    adoorder.Recordset.Cancel
    
    'deinitialising objects
    cmdnw_ordr.Enabled = True
    txtadvance.Locked = True
    txtprtclrs.Locked = True
    DataCombo1.Visible = False
    Combo1.Visible = False
    DataCombo3.Visible = False
    DTPicker2.Visible = False
    cmdcncl.Enabled = False
    Frame1.Enabled = False
    cmdadd.Enabled = False
    cmdnw_ordr.BackColor = &H8000000F
    cmdcncl.BackColor = &H8000000F
    
    'setting values to zeros
    txtdue = ""
    txttotal = ""
    txtadvance = 0
    txtacost = 500
    txtaddcosts = 0
    txtdue = 0
    
    tmrtotal.Enabled = False
    'closing temp recordset
    rsTemp.Close
    Set rsTemp = Nothing
    
    cmdnw_ordr.Caption = "PLACE ORDER"
    MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"

End Sub


Private Sub cmddelorder_Click()
On Error GoTo errorhandler
ask = MsgBox("Are you sure you want to undo this order " & txttempOID & " ?", vbCritical + vbYesNo, "System Query")
 If ask = vbYes Then
    adodealer.RecordSource = "SELECT *FROM dealer WHERE (((dealer.DID) Like '" & Me.Text2.Text & "%'))"
    adodealer.CommandType = adCmdText
    adodealer.Refresh
    adodealer.Recordset.Fields(4) = adodealer.Recordset.Fields(4).Value - adoorderlist.Recordset.Fields(9).Value
    adodealer.Recordset.Fields(5) = adodealer.Recordset.Fields(5).Value - adoorderlist.Recordset.Fields(10).Value
    adodealer.Recordset.Fields(6) = adodealer.Recordset.Fields(6).Value - adoorderlist.Recordset.Fields(11).Value
    adodealer.Recordset.Update
    adodealer.Refresh
    
    adoorderdetails.RecordSource = "SELECT *FROM dealer_order_details WHERE (((dealer_order_details.OID) Like '" & Me.txttempOID.Text & "%'))"
    adoorderdetails.CommandType = adCmdText
    adoorderdetails.Refresh
    
    
    Do Until adoorderdetails.Recordset.RecordCount = 0
        adoorderdetails.Recordset.MoveFirst
        adoorderdetails.Recordset.Delete
        adoorderdetails.Recordset.Update
    Loop

    adoorderdetails.RecordSource = "SELECT *FROM dealer_order_details"
    adoorderdetails.CommandType = adCmdText
    adoorderdetails.Refresh

    adoorderlist.Recordset.Delete
    adoorderlist.Refresh
    adoorderlist.Refresh
    adoorderlist.Refresh
    
    
MsgBox "Order rolled back", vbInformation + vbOKOnly, "System Information"

End If
errorhandler:
If Err <> 0 Then
 MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub cmddone_Click()
On Error GoTo errhandler
Dim Counter
Dim break
break = 0
If Label3.Caption = "PENDING" Then
        ask = MsgBox("Do you want to mark this order" & " " & adoorderlist.Recordset.Fields(0).Value & " " & " as PREPARED ?", vbQuestion + vbYesNo, "System Query")
        If ask = vbYes Then
       'updating stock balance of ckd
            temporderdetails.Refresh
            Do
                adostock_ckd.Refresh
                
                
                Do While temporderdetails.Recordset.Fields(1).Value <> adostock_ckd.Recordset.Fields(0).Value
                    adostock_ckd.Recordset.MoveNext
                Loop
                Counter = Counter + 1
                If adostock_ckd.Recordset.Fields(2).Value >= temporderdetails.Recordset.Fields(3).Value Then
                adostock_ckd.Recordset.Fields(2) = Val(adostock_ckd.Recordset.Fields(2).Value) - Val(temporderdetails.Recordset.Fields(3).Value)
                adostock_ckd.Recordset.Save
                adostock_ckd.Recordset.Update
                adostock_ckd.Refresh
                adoorderlist.Recordset.Fields(4) = "PREPARED"
                adoorderlist.Recordset.Update
                
                    If temporderdetails.Recordset.EOF = False Then
                        temporderdetails.Recordset.MoveNext
                    End If
                
                
                Else: break = 1
                End If
        Loop Until Counter = temporderdetails.Recordset.RecordCount
        
        'updating stock balance of assembled bikes
        Counter = 0
        temporderdetails.Refresh
        
        
        Do
                adostock_assembled.Refresh
                
                
                Do While temporderdetails.Recordset.Fields(1).Value <> adostock_assembled.Recordset.Fields(0).Value
                    adostock_assembled.Recordset.MoveNext
                Loop
                Counter = Counter + 1
                adostock_assembled.Recordset.Fields(2) = Val(adostock_assembled.Recordset.Fields(2).Value) + Val(temporderdetails.Recordset.Fields(3).Value)
                adostock_assembled.Recordset.Save
                adostock_assembled.Recordset.Update
                adostock_assembled.Refresh
                    If temporderdetails.Recordset.EOF = False Then
                        temporderdetails.Recordset.MoveNext
                    End If
        Loop Until Counter = temporderdetails.Recordset.RecordCount
            If break = 0 Then
            MsgBox "Order status has been changed to PREPARED", vbInformation + vbOKOnly, "System Information"
            ElseIf break = 1 Then
            MsgBox "Sorry you dont have sufficient CKDs. Please update CKD stock and try again.", vbCritical + vbOKOnly, "System Error"
            End If
        
        End If
Else
        MsgBox "Sorry you can't change this status.", vbCritical + vbOKOnly, "System Error"

End If
errhandler:
If Err.Number = 3021 Then
MsgBox "Critical Error. Please check for any missing records.", vbCritical + vbOKOnly, "System Error"
End If
End Sub

Private Sub cmdnw_ordr_Click()
On Error GoTo errorhandler
    'completing transaction
If cmdnw_ordr.Caption = "COMPLETE ORDER" Then
    If rsTemp.RecordCount = 0 Then
        MsgBox "You didnt select any product", vbExclamation, "System Error"
    ElseIf txtprtclrs = "" Then
        MsgBox "Please input the particulars", vbExclamation + vbOKOnly, "System Error"
        txtprtclrs.SetFocus
    ElseIf txtdelr_id = "" Then
        MsgBox "Please select Dealer ID", vbExclamation + vbOKOnly, "System Error"
        txtdelr_id.SetFocus
    ElseIf txtdlvry_date = "" Then
        MsgBox "Please select delivery date", vbExclamation + vbOKOnly, "System Error"
        txtdlvry_date.SetFocus
    ElseIf txtadvance = "" Then
        MsgBox "Please input payment", vbExclamation + vbOKOnly, "System Error"
        txtadvance.SetFocus
    ElseIf Combo2.Text = "" Then
        MsgBox "Please select payment mode", vbExclamation + vbOKOnly, "System Error"
        Combo2.SetFocus
    Else

        Dim TCP As Double
        Dim PROFIT As Double

    'updating delaer db
    adodealer.Recordset.Fields(4) = adodealer.Recordset.Fields(4).Value + Val(txttotal)
    adodealer.Recordset.Fields(5) = adodealer.Recordset.Fields(5).Value + Val(txtadvance)
    adodealer.Recordset.Fields(6) = adodealer.Recordset.Fields(6).Value + Val(txtdue)
    adodealer.Recordset.Update
    adodealer.Refresh
    
    'copying of update temp to order details db
    rsTemp.MoveFirst
    For Counter = 1 To rsTemp.RecordCount
        adoorderdetails.Recordset.AddNew
        adoorderdetails.Recordset.Fields(0) = txtordr_id.Text
        adoorderdetails.Recordset.Fields(1) = rsTemp.Fields(0).Value
        adoorderdetails.Recordset.Fields(2) = rsTemp.Fields(1).Value
        adoorderdetails.Recordset.Fields(3) = rsTemp.Fields(2).Value
        adoorderdetails.Recordset.Fields(4) = rsTemp.Fields(3).Value
        adoorderdetails.Recordset.Fields(5) = rsTemp.Fields(4).Value
        adoorderdetails.Recordset.Update
        adoorderdetails.Refresh
        If rsTemp.EOF = False Then
        rsTemp.MoveNext
        End If
    Next
    
    
    'adding up TCP
     rsTemp.MoveFirst
     Do While rsTemp.EOF = False
        TCP = rsTemp.Fields(5).Value + TCP
        rsTemp.MoveNext
     Loop
    PROFIT = Val(txttotal) - TCP
    
    'adding assembly cost
     rsTemp.MoveFirst
     Do While rsTemp.EOF = False
        ASS_COST = rsTemp.Fields(2).Value + ASS_COST
        rsTemp.MoveNext
     Loop
     ASS_COST = ASS_COST * Val(txtacost)
    
    'order update
    adoorder.Recordset.AddNew
    adoorder.Recordset.Fields(0) = txtordr_id.Text
    adoorder.Recordset.Fields(1) = txtdelr_id.Text
    adoorder.Recordset.Fields(2) = txtdelr_name.Text
    adoorder.Recordset.Fields(3) = txtprtclrs.Text
    adoorder.Recordset.Fields(4) = txtdlvry_stts.Text
    adoorder.Recordset.Fields(5) = txtordr_date.Text
    adoorder.Recordset.Fields(6) = txtdlvry_date.Text
    adoorder.Recordset.Fields(7) = Combo2.Text
    'adoorder.Recordset.Fields(8) = ""
    adoorder.Recordset.Fields(9) = txttotal.Text
    adoorder.Recordset.Fields(10) = txtadvance.Text
    adoorder.Recordset.Fields(11) = txtdue.Text
    adoorder.Recordset.Fields(12) = frmmain.stsbr_main.Panels(2).Text
    adoorder.Recordset.Fields(13) = adodealer.Recordset.Fields(14).Value
    adoorder.Recordset.Fields(14) = PROFIT
    adoorder.Recordset.Fields(15) = txtaddcosts.Text
    adoorder.Recordset.Fields(16) = ASS_COST
    adoorder.Recordset.Update
    adoorder.Refresh

    'deinitialising objects
    cmdnw_ordr.Enabled = True
    txtadvance.Locked = True
    txtprtclrs.Locked = True
    DataCombo1.Visible = False
    Combo1.Visible = False
    DataCombo3.Visible = False
    DTPicker2.Visible = False
    cmdcncl.Enabled = False
    Frame1.Enabled = False
    cmdadd.Enabled = False
  
    'setting values to zeros
    txtdue = ""
    txttotal = ""
    txtadvance = 0
    txtacost = 500
    txtaddcosts = 0
    txtdue = 0
        
        
    'closing temp recordset
    rsTemp.Close
    Set rsTemp = Nothing
    
    
    
        cmdnw_ordr.BackColor = &H8000000F
        cmdcncl.BackColor = &H8000000F
          
        MsgBox "Order Placed!", vbInformation + vbOKOnly, "System Information"
        cmdnw_ordr.Caption = "PLACE ORDER"
        cmdcncl.Enabled = False
          
         
    If dev_bike.rscmdinvorders.State = adStateOpen Then
        dev_bike.rscmdinvorders.Close
    End If
    dev_bike.cmdinvorders Trim(txtordr_id.Text)
    rptinv_order.Show
          
          
    End If


ElseIf cmdnw_ordr.Caption = "PLACE ORDER" Then


    'connecting database
    adodealer.Refresh
    adoorder.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adoorder.RecordSource = "SELECT *FROM dealer_order"
    adoorder.CommandType = adCmdText
    adoorder.Refresh


    'initialising objects
    txtprtclrs.Locked = False
    DataCombo1.Visible = True
    DataCombo3.Visible = True
    DTPicker2.Visible = True
    cmdcncl.Enabled = True
    txtadvance.Locked = False
    cmdadd.Enabled = True
    DTPicker2 = Date
    cmdadd.Enabled = True
    Frame1.Enabled = True
    
    
    
        'creating temp table
        Set rsTemp = New ADODB.Recordset
        rsTemp.ActiveConnection = Nothing
        rsTemp.CursorLocation = adUseClient
        rsTemp.LockType = adLockBatchOptimistic
    
        rsTemp.Fields.Append "PARTS ID", adVarChar, 20
        rsTemp.Fields.Append "PARTS DESCRIPTION", adVarChar, 50
        rsTemp.Fields.Append "QTY", adVarNumeric, 10
        rsTemp.Fields.Append "UNIT PRICE", adVarNumeric, 10
        rsTemp.Fields.Append "TOTAL COST", adVarNumeric, 20
        rsTemp.Fields.Append "TCP", adVarNumeric, 20
        rsTemp.Open
        Set DataGrid1.DataSource = rsTemp
        DataGrid1.Columns(5).Visible = False
        DataGrid1.Refresh


    'creating an order ID

    Dim varOID
    If adoorder.Recordset.RecordCount = 0 Then
        txtordr_id = "OID1005001"
        txtordr_date = Format(Date, "dd/MM/yyyy")
        txtdlvry_stts.Text = "PENDING"
    Else: adoorder.Recordset.Sort = "OID ASC"
        adoorder.Recordset.MoveLast
        varOID = Mid(adoorder.Recordset.Fields(0).Value, 4, 10)
        txtordr_id.Text = "OID" & CStr(varOID + 1)
        txtordr_date = Format(Date, "dd/MM/yyyy")
        txtdlvry_stts.Text = "PENDING"
    End If

    cmdnw_ordr.Caption = "COMPLETE ORDER"
     cmdnw_ordr.BackColor = vbGreen
    cmdcncl.BackColor = vbRed

End If

   

errorhandler:

If Err <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub cmdordr_his_Click()
If dev_bike.rscmddealerorderhistorybyDID.State = adStateOpen Then
    dev_bike.rscmddealerorderhistorybyDID.Close
End If
    dev_bike.cmddealerorderhistorybyDID Trim(Text2.Text)
    rptdelrordrhis.Show
End Sub


Private Sub cmdrmv_Click()
On Error GoTo RmvDataError
   
RmvDataError:
    If Err.Number = 3021 Then
        MsgBox "Select the product you want to remove from the list", vbInformation + vbOKOnly, "Data remove error"
    End If
End Sub


Private Sub cmdref_Click()
Adodc1.Refresh
adoorder_details.Refresh
adoorder_details.Refresh
adochckbike.Refresh
DataGrid4.Refresh
DataCombo1.ReFill
DataCombo1.Refresh


 MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Information"
End Sub

Private Sub Combo1_Click()
txtdlvry_stts.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
If adoorderlist.RecordSource <> "" Then
    adoorderlist.Refresh
    adoorderlist.Refresh
    DataGrid3.Columns(1).Visible = False
    DataGrid3.Columns(2).Visible = False
    DataGrid3.Columns(4).Visible = False
    DataGrid3.Columns(12).Visible = False
    DataGrid3.Columns(13).Visible = False
    DataGrid3.Columns(14).Visible = False
    DataGrid3.Columns(5).Visible = False
    DataGrid3.Columns(6).Visible = False
    DataGrid3.Columns(8).Visible = False
    MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Information"
End If

End Sub

Private Sub Combo2_Click()
If Combo2.Text = "BANK" Then
        Label1(14).Visible = True
        txtbank.Visible = True
   ElseIf Combo2.Text <> bank Then
        'Label1(14).Visible = False
        'txtbank.Visible = False
    End If
End Sub

Private Sub Command2_Click()
On Error GoTo errhandler
Dim dounter
Dim break
break = 0
If Label3.Caption = "PREPARED" Then
        ask = MsgBox("Do you want to mark this order" & " " & adoorderlist.Recordset.Fields(0).Value & " " & " as PENDING ?", vbQuestion + vbYesNo, "System Query")
        If ask = vbYes Then
       'updating stock balance of assembled bikes
            temporderdetails.Refresh
            Do
                adostock_assembled.Refresh
                
                
                Do While temporderdetails.Recordset.Fields(1).Value <> adostock_assembled.Recordset.Fields(0).Value
                    adostock_assembled.Recordset.MoveNext
                Loop
                dounter = dounter + 1
                If adostock_assembled.Recordset.Fields(2).Value >= temporderdetails.Recordset.Fields(3).Value Then
                adostock_assembled.Recordset.Fields(2) = Val(adostock_assembled.Recordset.Fields(2).Value) - Val(temporderdetails.Recordset.Fields(3).Value)
                adostock_assembled.Recordset.Save
                adostock_assembled.Recordset.Update
                adostock_assembled.Refresh
                adoorderlist.Recordset.Fields(4) = "PENDING"
                adoorderlist.Recordset.Update
                
                    If temporderdetails.Recordset.EOF = False Then
                        temporderdetails.Recordset.MoveNext
                    End If
                
                
                Else: break = 1
                End If
        Loop Until dounter = temporderdetails.Recordset.RecordCount
        
        'updating stock balance of ckd bikes
        dounter = 0
        temporderdetails.Refresh
        
        
        Do
                adostock_ckd.Refresh
                
                Do While temporderdetails.Recordset.Fields(1).Value <> adostock_ckd.Recordset.Fields(0).Value
                     adostock_ckd.Recordset.MoveNext
                Loop
                dounter = dounter + 1
                adostock_ckd.Recordset.Fields(2) = Val(adostock_ckd.Recordset.Fields(2).Value) + Val(temporderdetails.Recordset.Fields(3).Value)
                adostock_ckd.Recordset.Save
                adostock_ckd.Recordset.Update
                adostock_ckd.Refresh
                    If temporderdetails.Recordset.EOF = False Then
                        temporderdetails.Recordset.MoveNext
                    End If
        Loop Until dounter = temporderdetails.Recordset.RecordCount
            If break = 0 Then
            MsgBox "Order status has been changed to PENDING", vbInformation + vbOKOnly, "System Information"
            ElseIf break = 1 Then
            MsgBox "Sorry you dont have sufficient BIKEs. Please update BIKE stock and try again.", vbCritical + vbOKOnly, "System Error"
            End If
        
        End If
Else
        MsgBox "Sorry you can't change this status.", vbCritical + vbOKOnly, "System Error"

End If
errhandler:
If Err.Number = 3021 Then
MsgBox "Critical Error. Please check for any missing records.", vbCritical + vbOKOnly, "System Error"
End If
End Sub




Private Sub Command3_Click()
Dim prmfrom
Dim prmto
prmfrom = fromdate.Value
prmto = todate.Value
If dev_bike.rscmdorderbtndates.State = adStateOpen Then
  dev_bike.rscmdorderbtndates.Close
 End If
 dev_bike.cmdorderbtndates prmfrom, prmto
rptorderbydates.Show
End Sub

Private Sub Command5_Click()
If Combo3.Text = "BY DID" And txtsearch.Text <> "" Then
    If dev_bike.rscmdORDERbyDID.State = adStateOpen Then
        dev_bike.rscmdORDERbyDID.Close
    End If
        dev_bike.cmdORDERbyDID Trim(txtsearch.Text)
        rptorderbyDID.Show
ElseIf Combo3.Text = "BY OID" And txtsearch.Text <> "" Then
    If dev_bike.rscmdORDERbyOID.State = adStateOpen Then
        dev_bike.rscmdORDERbyOID.Close
    End If
        dev_bike.cmdORDERbyOID Trim(txtsearch.Text)
        rptorderbyOID.Show

Else
  MsgBox "Operation is cancelled because you didnt input correct parameters ", vbCritical + vbOKOnly, "Input Error"
End If
End Sub

Private Sub Command4_Click()
    Dim var_cap As String
If dev_bike.rscmdORDER_DETAILS.State = adStateOpen Then
    dev_bike.rscmdORDER_DETAILS.Close
End If
    dev_bike.cmdORDER_DETAILS Trim(Label12.Caption)
    var_cap = "ORDER ID " + Label12.Caption
    rptorder_details.Sections("PageHeader").Controls.Item("label1").Caption = var_cap
    rptorder_details.Show
End Sub

Private Sub Command6_Click()
If dev_bike.rscmdinvorders.State = adStateOpen Then
    dev_bike.rscmdinvorders.Close
End If
    dev_bike.cmdinvorders Trim(DataCombo2.Text)
    rptinv_order.Show
End Sub

Private Sub Command7_Click()
On Error GoTo errorhandler
If Not Adodc1.Recordset.RecordCount = 0 Then
Dim Counter As Integer
Counter = 0
Dim stts As Integer
stts = 0
    adoorder_details.Recordset.MoveFirst
    adochckbike.Recordset.MoveFirst
    Do
    Do While adoorder_details.Recordset.Fields(1).Value <> adochckbike.Recordset.Fields(0).Value
        adochckbike.Recordset.MoveNext
    Loop
        Counter = Counter + 1
        If adochckbike.Recordset.Fields(2).Value < adoorder_details.Recordset.Fields(3).Value Then
        Else: stts = stts + 1
        End If
        If adoorder_details.Recordset.EOF = False Then
        adoorder_details.Recordset.MoveNext
        adochckbike.Recordset.MoveFirst
        End If
    Loop While Counter <> adoorder_details.Recordset.RecordCount
        If stts = adoorder_details.Recordset.RecordCount Then
            deliver = MsgBox("You you sure you want to deliver this stock ?", vbYesNo + vbQuestion, "System query")
                If deliver = vbYes Then
                
                    Adodc1.Recordset.Fields(4) = "DELIVERED"
                    Adodc1.Recordset.Fields(6) = Format(Date, "dd/MM/yyyy")
                    Adodc1.Recordset.Update
                    Adodc1.Refresh
                    adoorder_details.Refresh
                    adochckbike.Refresh
                    Adodc1.Refresh
                    adoorder_details.Refresh
                    adochckbike.Refresh
                    DataCombo1.ReFill
                    MsgBox "Order delivered.", vbOKOnly + vbInformation, "System Information"
                     
                End If
        Else
            MsgBox "Stock is low on one or more items.", vbOKOnly + vbCritical, "System Information"
        End If
Else: MsgBox "Sorry no order is prepared to be delivered.", vbInformation + vbOKOnly, "Borac Sales System"
End If
errorhandler:
Select Case Err.Number
    Case 3021
        MsgBox "Process is cancelled due to a fatal error. Possible reasons are incomplete records or missing bike ID. ", vbCritical + vbOKOnly, "Database Error"
End Select
End Sub

Private Sub DataCombo1_Change()
    adodealer.Recordset.Bookmark = DataCombo1.SelectedItem
    txtdelr_id.Text = DataCombo1.Text
    'txtdelr_name = ""

End Sub

Private Sub DataCombo1_Click(Area As Integer)
    txtdelr_name = adodealer.Recordset.Fields(1).Value
End Sub




Private Sub DataCombo3_Change()
    adodealer.Recordset.Bookmark = DataCombo3.SelectedItem
    txtdelr_name.Text = DataCombo3.Text
End Sub

Private Sub DataCombo3_Click(Area As Integer)
    txtdelr_id = adodealer.Recordset.Fields(0).Value
End Sub

Private Sub DataCombo4_Change()
Adodc1.Recordset.Bookmark = DataCombo4.SelectedItem
End Sub

Private Sub DataGrid2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    DataGrid2.Columns(0).Width = 1200
    DataGrid2.Columns(1).Width = 2580
    DataGrid2.Columns(2).Width = 1065
End Sub

Private Sub DataGrid1_GotFocus()
cmdadd.Caption = "REMOVE FROM CART"
End Sub

Private Sub DataList1_Click()
On Error GoTo errorhandler
    adostock_ckd.Recordset.Bookmark = DataList1.SelectedItem
    txtitem = adostock_ckd.Recordset.Fields!bid
    Text1 = adostock_ckd.Recordset.Fields!unit_price
errorhandler:
If Err <> 0 Then
 MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub DataList1_GotFocus()
cmdadd.Caption = "ADD TO CART"
End Sub

Private Sub DTPicker1_Change()
    txtordr_date.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
    txtdlvry_date.Text = Format(DTPicker2.Value, "dd/MM/yyyy")
End Sub

Private Sub DTPicker2_Click()
    txtdlvry_date.Text = Format(DTPicker2.Value, "dd/MM/yyyy")
End Sub

Private Sub Form_Load()
img_title.Width = Me.Width
fpostn = (frmorder.Width - framex.Width) / 2
framex.Left = fpostn
'DataGrid1.Columns(4).Visible = False
'adoorderlist.Recordset.MoveLast
'DataGrid2.Columns(0).Width = 1200
'DataGrid2.Columns(1).Width = 2580
'DataGrid2.Columns(2).Width = 1065
todate = Date
fromdate = Date

    
End Sub

Private Sub Form_Resize()
img_title.Width = Me.Width
fpostn = (frmorder.Width - framex.Width) / 2
framex.Left = fpostn
End Sub

Private Sub Timer1_Timer()



 'updating subtotal by calculating the all firld value
    Dim sub_total As Currency
    If adotemp.Recordset.BOF = False Then
        adotemp.Recordset.MoveFirst
            Do While adotemp.Recordset.EOF = False
                sub_total = adotemp.Recordset.Fields(4).Value + sub_total
                adotemp.Recordset.MoveNext
            Loop
       End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
If cmdnw_ordr.Caption = "COMPLETE ORDER" Then
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

Private Sub Label3_Change()
If Label3.Caption = "PREPARED" Then
    Label3.BackColor = &H8000000D
ElseIf Label3.Caption = "DELIVERED" Then
    Label3.BackColor = &HFF00&
ElseIf Label3.Caption = "PENDING" Then
    Label3.BackColor = &HFF&
ElseIf Label3.Caption = "" Then
    Label3.BackColor = vbWhite
End If
End Sub

Private Sub optdealer_Click()
If optdealer.Value = True Then
    adoorderlist.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adoorderlist.CommandType = adCmdText
    adoorderlist.RecordSource = "SELECT *FROM dealer_order WHERE (((dealer_order.STATUS) LIKE '" & Me.optdealer.Caption & "%'))"
    adoorderlist.Refresh
    '
    DataGrid3.Columns(1).Visible = False
    DataGrid3.Columns(2).Visible = False
    DataGrid3.Columns(4).Visible = False
    DataGrid3.Columns(12).Visible = False
    DataGrid3.Columns(13).Visible = False
    DataGrid3.Columns(14).Visible = False
    DataGrid3.Columns(5).Visible = False
    DataGrid3.Columns(6).Visible = False
    DataGrid3.Columns(8).Visible = False
    DataGrid3.Refresh
End If
End Sub

Private Sub optother_Click()
If optother.Value = True Then
    adoorderlist.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adoorderlist.CommandType = adCmdText
    adoorderlist.RecordSource = "SELECT *FROM dealer_order WHERE (((dealer_order.STATUS) LIKE '" & Me.optother.Caption & "%'))"
    adoorderlist.Refresh
    
    DataGrid3.Columns(1).Visible = False
    DataGrid3.Columns(2).Visible = False
    DataGrid3.Columns(4).Visible = False
    DataGrid3.Columns(12).Visible = False
    DataGrid3.Columns(13).Visible = False
    DataGrid3.Columns(14).Visible = False
    DataGrid3.Columns(5).Visible = False
    DataGrid3.Columns(6).Visible = False
    DataGrid3.Columns(8).Visible = False
    DataGrid3.Refresh
End If
End Sub

Private Sub optretail_Click()
If optretail.Value = True Then
    adoorderlist.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
    adoorderlist.CommandType = adCmdText
    adoorderlist.RecordSource = "SELECT *FROM dealer_order WHERE (((dealer_order.STATUS) LIKE '" & Me.optretail.Caption & "%'))"
    adoorderlist.Refresh
    
    DataGrid3.Columns(1).Visible = False
    DataGrid3.Columns(2).Visible = False
    DataGrid3.Columns(4).Visible = False
    DataGrid3.Columns(12).Visible = False
    DataGrid3.Columns(13).Visible = False
    DataGrid3.Columns(14).Visible = False
    DataGrid3.Columns(5).Visible = False
    DataGrid3.Columns(6).Visible = False
    DataGrid3.Columns(8).Visible = False
    DataGrid3.Refresh
End If
End Sub



Private Sub ordr_sttschck_Timer()

End Sub

Private Sub tmrtotal_Timer()
'On Error Resume Next 'updating subtotal by calculating the all field value
    Dim g_total As Currency
    Dim bike_count As Integer
    If rsTemp.BOF = False Then
        rsTemp.MoveFirst
            Do While rsTemp.EOF = False
                g_total = rsTemp.Fields(4).Value + g_total
                bike_count = rsTemp.Fields(2).Value + bike_count
                rsTemp.MoveNext
            Loop
    End If
    
        txtsubtotal = g_total
        txttotal = g_total + Val(txtaddcosts) + Val(bike_count) * Val(txtacost)
        txtdue = Val(txttotal) - Val(txtadvance)
    tmrtotal.Enabled = False
End Sub

Private Sub txtacost_Change()
    If IsNumeric(txtacost) = False Then
        txtacost = 500
    ElseIf txtacost = 0 Or txtacost = "" Then
        txtacost = 500
    End If
      
End Sub

Private Sub txtacost_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(txtacost) = False Then
        txtacost = 500
        'tmrtotal.Enabled = True
    ElseIf txtacost = 0 Then
        txtacost = 500
        'tmrtotal.Enabled = True
    ElseIf txtacost = "" Then
        txtacost = 500
        'tmrtotal.Enabled = True
    End If
    tmrtotal.Enabled = True
End Sub

Private Sub txtaddcosts_Change()
If IsNumeric(txtaddcosts) = False Then
        txtaddcosts = 0
Else: tmrtotal.Enabled = True
End If
End Sub

Private Sub txtadvance_Change()
    If IsNumeric(txtadvance) = False Then
        txtadvance.Text = ""
    End If
    
End Sub

Private Sub txtadvance_KeyUp(KeyCode As Integer, Shift As Integer)
tmrtotal.Enabled = True
End Sub

Private Sub txtdlvry_stts_Change()
If txtdlvry_stts.Text = "PENDING" Then
    txtdlvry_stts.ForeColor = &HFF&
ElseIf txtdlvry_stts.Text = "DELIVERED" Then
    txtdlvry_stts.ForeColor = &HC000&
ElseIf txtdlvry_stts.Text = "PREPARED" Then
    txtdlvry_stts.ForeColor = &HFF0000
End If
End Sub

Private Sub txtidchck_Change()
If Not Adodc1.Recordset.RecordCount = 0 Then
adoorder_details.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
adoorder_details.RecordSource = "SELECT dealer_order_details.* FROM dealer_order_details WHERE (((dealer_order_details.OID) Like '" & Me.txtidchck.Text & "%'))"
adoorder_details.CommandType = adCmdText
adoorder_details.Refresh
End If
End Sub

Private Sub txtqty_Change()
On Error GoTo errorhandler
    Dim a As Integer
    If IsNumeric(txtqty) = False Then
        txtqty = ""
    End If
    a = Val(txtqty.Text)
    txtqty = a
errorhandler:
Select Case Err.Number
    Case 6
        MsgBox "Your input is out of range.", vbCritical + vbOKOnly, "Data input error"
        txtqty = ""
End Select
End Sub

Private Sub txttempOID_Change()
'connecting temporary adodc to table dealer order details
            temporderdetails.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
            temporderdetails.CommandType = adCmdText
            temporderdetails.RecordSource = "SELECT * FROM dealer_order_details WHERE (((dealer_order_details.OID) Like '" & Me.txttempOID.Text & "%'))"
            temporderdetails.Refresh
           
            
            
            
            
End Sub

Private Sub txttotal_Change()
txtdue = Val(txttotal) - Val(txtadvance)
End Sub

