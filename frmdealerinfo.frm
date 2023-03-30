VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdealerinfo 
   Caption         =   "Dealer Info"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15255
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdealerinfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   15255
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "frmdealerinfo.frx":F172
      Height          =   255
      Left            =   7560
      TabIndex        =   117
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
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
   Begin TabDlg.SSTab framex 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   19965
      _ExtentX        =   35216
      _ExtentY        =   15690
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Dealer List"
      TabPicture(0)   =   "frmdealerinfo.frx":F191
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2(2)"
      Tab(0).Control(1)=   "Image3"
      Tab(0).Control(2)=   "Label33"
      Tab(0).Control(3)=   "DataGrid2"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "Command3"
      Tab(0).Control(6)=   "txtloc"
      Tab(0).Control(7)=   "Command6"
      Tab(0).Control(8)=   "Text5"
      Tab(0).Control(9)=   "Text19"
      Tab(0).Control(10)=   "Command12"
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(12)=   "Command8"
      Tab(0).Control(13)=   "Command9"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Add New Dealer"
      TabPicture(1)   =   "frmdealerinfo.frx":F1AD
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "framexx"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dealer Payback"
      TabPicture(2)   =   "frmdealerinfo.frx":F1C9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Adodc1"
      Tab(2).Control(1)=   "adolist"
      Tab(2).Control(2)=   "adodealerpayment"
      Tab(2).Control(3)=   "adopayment"
      Tab(2).Control(4)=   "Frame2"
      Tab(2).Control(5)=   "DataGrid3"
      Tab(2).Control(6)=   "Frame3"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Others"
      TabPicture(3)   =   "frmdealerinfo.frx":F1E5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdview_o_list"
      Tab(3).Control(1)=   "cmdrefresh"
      Tab(3).Control(2)=   "cmdnew"
      Tab(3).Control(3)=   "cmdupdate"
      Tab(3).Control(4)=   "cmdeditentry"
      Tab(3).Control(5)=   "cmdcancelentry"
      Tab(3).Control(6)=   "adoothers"
      Tab(3).Control(7)=   "DataCombo5"
      Tab(3).Control(8)=   "Frame5"
      Tab(3).Control(9)=   "DataGrid4"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Search"
      TabPicture(4)   =   "frmdealerinfo.frx":F201
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "Command10"
      Tab(4).Control(2)=   "Command11"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton Command9 
         Caption         =   "Spares Tran. History Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71880
         TabIndex        =   137
         Top             =   8040
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "BIKE Tran. History Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69120
         TabIndex        =   136
         Top             =   8040
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CKD Tran. History Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66720
         TabIndex        =   135
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton Command12 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   -60600
         TabIndex        =   134
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -65400
         TabIndex        =   133
         Top             =   6600
         Width           =   4695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Dealer's Dues REPORT"
         Height          =   375
         Left            =   -74760
         TabIndex        =   129
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "All dealer's REPORT"
         Height          =   375
         Left            =   -74760
         TabIndex        =   128
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         DataField       =   "DID"
         DataSource      =   "adodealerlist"
         Height          =   375
         Left            =   -69240
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Frame6 
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
         TabIndex        =   121
         Top             =   480
         Width           =   7695
         Begin VB.CommandButton Command7 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   124
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "frmdealerinfo.frx":F21D
            Left            =   240
            List            =   "frmdealerinfo.frx":F227
            TabIndex        =   123
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtsearch 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2040
            TabIndex        =   122
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71880
         TabIndex        =   118
         Top             =   7440
         Width           =   1695
      End
      Begin VB.CommandButton cmdview_o_list 
         Caption         =   "&View others list"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66360
         TabIndex        =   113
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -68040
         TabIndex        =   112
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "&New Entry"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   111
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73080
         TabIndex        =   110
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdeditentry 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -71400
         TabIndex        =   109
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmdcancelentry 
         Caption         =   "&Cancel Entry"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69720
         TabIndex        =   108
         Top             =   6240
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc adoothers 
         Height          =   375
         Left            =   -74640
         Top             =   5760
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         RecordSource    =   "others"
         Caption         =   "Others"
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
         Left            =   -65160
         Top             =   660
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
      Begin MSAdodcLib.Adodc adolist 
         Height          =   330
         Left            =   -66360
         Top             =   660
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
         RecordSource    =   "delaer_payback"
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
      Begin MSAdodcLib.Adodc adodealerpayment 
         Height          =   330
         Left            =   -68760
         Top             =   660
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
         RecordSource    =   "delaer_payback"
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc adopayment 
         Height          =   330
         Left            =   -67560
         Top             =   660
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
         RecordSource    =   "SELECT *FROM dealer WHERE dealer.BALANCE>0"
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
      Begin VB.Frame Frame2 
         Caption         =   "Choose dealer for the late payment"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   11175
         Begin VB.TextBox Text16 
            DataField       =   "DID"
            DataSource      =   "adopayment"
            Height          =   375
            Left            =   5040
            TabIndex        =   126
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdref 
            Caption         =   "&Refresh"
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
            Left            =   3840
            TabIndex        =   69
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&New Payment"
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
            Left            =   240
            TabIndex        =   41
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtbname 
            Appearance      =   0  'Flat
            DataField       =   "BUSINESS_NAME"
            DataSource      =   "adopayment"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7080
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1560
            Width           =   3735
         End
         Begin VB.TextBox txtarea 
            Appearance      =   0  'Flat
            DataField       =   "AREA"
            DataSource      =   "adopayment"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7080
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   840
            Width           =   3735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Show Payment History"
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
            Left            =   1680
            TabIndex        =   36
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtdue 
            Appearance      =   0  'Flat
            DataField       =   "BALANCE"
            DataSource      =   "adopayment"
            Height          =   375
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   360
            Width           =   3735
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "frmdealerinfo.frx":F23C
            DataField       =   "DID"
            Height          =   315
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DID"
            BoundColumn     =   ""
            Text            =   "DID"
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "frmdealerinfo.frx":F255
            DataField       =   "DEALER_NAME"
            Height          =   315
            Left            =   2760
            TabIndex        =   33
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DEALER_NAME"
            BoundColumn     =   ""
            Text            =   "NAME"
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
         Begin VB.Label Label32 
            BackColor       =   &H80000000&
            Caption         =   "BUSINESS NAME"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6120
            TabIndex        =   120
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000000&
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6120
            TabIndex        =   119
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000000&
            Caption         =   "Dealer ID:                                        Name: "
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   5655
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            DataField       =   "DEALER_NAME"
            DataSource      =   "adopayment"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   1200
            Width           =   5655
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000000&
            Caption         =   "DUES"
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
            Left            =   6120
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtloc 
         DataField       =   "IMG_LOC"
         DataSource      =   "adodealerlist"
         Height          =   405
         Left            =   -69000
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Order History Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   29
         Top             =   8040
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Dealer Information Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   28
         Top             =   7440
         Width           =   2895
      End
      Begin VB.Frame framexx 
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
         Height          =   7935
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   18495
         Begin VB.Frame Frame4 
            Caption         =   "Personal Info"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   120
            TabIndex        =   76
            Top             =   0
            Width           =   11415
            Begin VB.TextBox Text17 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "TIN"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   131
               Top             =   2760
               Width           =   3135
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "AREA"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   2280
               Width           =   3135
            End
            Begin VB.TextBox txtphn_3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               CausesValidation=   0   'False
               DataField       =   "PHONE_NUMBER3"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "00000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adodelr"
               Height          =   405
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   80
               Top             =   1800
               Width           =   3135
            End
            Begin VB.TextBox txtphn_2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               CausesValidation=   0   'False
               DataField       =   "PHONE_NUMBER2"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "00000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   79
               Top             =   1320
               Width           =   3135
            End
            Begin VB.TextBox txtphn_1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               CausesValidation=   0   'False
               DataField       =   "PHONE_NUMBER1"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "00000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox txt_name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "DEALER_NAME"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   360
               Width           =   3135
            End
            Begin VB.TextBox txt_eml 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "E-MAIL"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   1320
               Width           =   3615
            End
            Begin VB.TextBox txt_dob 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "JOINED ON"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   78
               Top             =   1800
               Width           =   1935
            End
            Begin VB.TextBox txtdlr_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               DataField       =   "DID"
               DataSource      =   "adodelr"
               ForeColor       =   &H80000015&
               Height          =   375
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtr_addrs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "RESIDANCE_ADDRESS"
               DataSource      =   "adodelr"
               Height          =   495
               Left            =   6720
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   77
               Top             =   2280
               Width           =   4455
            End
            Begin VB.TextBox txtvot_id 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "VOTER_ID"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   840
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   8760
               TabIndex        =   81
               Top             =   1800
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   78708739
               CurrentDate     =   40725
               MinDate         =   39083
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Bindings        =   "frmdealerinfo.frx":F26E
               Height          =   315
               Left            =   8760
               TabIndex        =   82
               Top             =   360
               Visible         =   0   'False
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   "DID"
               Text            =   "DataCombo2"
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
               BackColor       =   &H80000000&
               Caption         =   "TIN"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   130
               Top             =   2760
               Width           =   1575
            End
            Begin VB.Label Label10 
               BackColor       =   &H80000000&
               Caption         =   "Area"
               Height          =   375
               Left            =   120
               TabIndex        =   92
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000000&
               Caption         =   "Phone Number (3)"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   91
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000000&
               Caption         =   "Phone Number (2)"
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label8 
               BackColor       =   &H80000000&
               Caption         =   "Phone Number (1)"
               Height          =   375
               Left            =   120
               TabIndex        =   89
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label12 
               BackColor       =   &H80000000&
               Caption         =   "Name"
               Height          =   375
               Left            =   120
               TabIndex        =   88
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "E mail"
               Height          =   375
               Index           =   0
               Left            =   5160
               TabIndex        =   87
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000000&
               Caption         =   "Joining Date"
               Height          =   375
               Left            =   5160
               TabIndex        =   86
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label9 
               BackColor       =   &H80000000&
               Caption         =   "Dealer ID"
               Height          =   375
               Left            =   5160
               TabIndex        =   85
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000000&
               Caption         =   "Residance Address"
               Height          =   495
               Index           =   1
               Left            =   5160
               TabIndex        =   84
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000000&
               Caption         =   "Voter ID"
               Height          =   375
               Index           =   1
               Left            =   5160
               TabIndex        =   83
               Top             =   840
               Width           =   1575
            End
         End
         Begin VB.Frame frameinfo 
            Caption         =   "Bussiness Info"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   3480
            Width           =   11415
            Begin VB.TextBox txtsaf_dep 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "DEPOSIT_AMOUNT"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1320
               Width           =   2895
            End
            Begin VB.TextBox txtaddrs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "BUSINESS_ADDRESS"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox txtbus_name 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "BUSINESS_NAME"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   1920
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Top             =   360
               Width           =   3135
            End
            Begin VB.TextBox txtt_tran 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "TOTAL_AMOUNT"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtt_paid 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "TOTAL_PAID"
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtt_due 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               CausesValidation=   0   'False
               DataField       =   "BALANCE"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               DataSource      =   "adodelr"
               Height          =   375
               Left            =   6840
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               Caption         =   "-ve indicate advance payment"
               Height          =   255
               Left            =   8880
               TabIndex        =   75
               Top             =   1440
               Width           =   2415
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000000&
               Caption         =   "Safety Deposit / BDT"
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
               TabIndex        =   74
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000000&
               Caption         =   "Business Address"
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
               Index           =   0
               Left            =   120
               TabIndex        =   73
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000000&
               Caption         =   "Business Name"
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
               Index           =   1
               Left            =   120
               TabIndex        =   72
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "Total Trans / BDT"
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
               Index           =   2
               Left            =   5280
               TabIndex        =   27
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "Total Paid / BDT"
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
               Index           =   3
               Left            =   5280
               TabIndex        =   26
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "Balance / BDT"
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
               Index           =   4
               Left            =   5280
               TabIndex        =   25
               Top             =   1320
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdcncl 
            Caption         =   "&Cancel Entry"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   15480
            TabIndex        =   11
            Top             =   2160
            Width           =   1815
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   15480
            TabIndex        =   10
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton cmdupdt 
            Caption         =   "&Update"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   15480
            TabIndex        =   9
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&New Entry"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   15480
            TabIndex        =   8
            Top             =   120
            Width           =   1815
         End
         Begin VB.CommandButton cmdbws 
            Caption         =   "..."
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
            Height          =   375
            Left            =   11640
            TabIndex        =   7
            Top             =   4800
            Width           =   615
         End
         Begin VB.CommandButton cmdclr 
            Caption         =   "Clear"
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
            Height          =   375
            Left            =   12360
            TabIndex        =   6
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox txtimg_addrs 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DataField       =   "IMG_LOC"
            DataSource      =   "adodelr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12960
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   5760
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdrefsh 
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   15480
            TabIndex        =   4
            Top             =   2760
            Width           =   1815
         End
         Begin MSAdodcLib.Adodc adodelr 
            Height          =   450
            Left            =   120
            Top             =   7440
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   794
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
            Caption         =   "Dealer Information"
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
         Begin MSComDlg.CommonDialog diagpic 
            Left            =   15240
            Top             =   4560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".jpg"
            InitDir         =   "C:\dbase_bike\dealer_pics"
         End
         Begin VB.Image img_delr 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   4215
            Left            =   11640
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3615
         End
      End
      Begin MSAdodcLib.Adodc adodealerlist 
         Height          =   495
         Left            =   8280
         Top             =   -600
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
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
         RecordSource    =   "SELECT DID,DEALER_NAME,AREA,DEPOSIT_AMOUNT,TOTAL_AMOUNT,TOTAL_PAID,BALANCE,IMG_LOC FROM dealer"
         Caption         =   "Dealer Information"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmdealerinfo.frx":F284
         Height          =   5415
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   4
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
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmdealerinfo.frx":F2A0
         Height          =   315
         Left            =   -69840
         TabIndex        =   114
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "OTHER_NAME"
         Text            =   ""
      End
      Begin VB.Frame Frame5 
         Caption         =   "Personal Info"
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
         Height          =   2535
         Left            =   -74760
         TabIndex        =   93
         Top             =   480
         Width           =   12495
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            DataField       =   "PROFESSION"
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   8640
            TabIndex        =   100
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            DataField       =   "JOINED ON"
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   8640
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            DataField       =   "E-MAIL"
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   8640
            TabIndex        =   98
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            DataField       =   "OTHER_NAME"
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   1680
            TabIndex        =   97
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            DataField       =   "PHONE_1"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   1680
            TabIndex        =   96
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            DataField       =   "PHONE_2"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "00000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoothers"
            Height          =   375
            Left            =   1680
            TabIndex        =   95
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            DataField       =   "PHONE_3"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "00000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoothers"
            Height          =   405
            Left            =   1680
            TabIndex        =   94
            Top             =   1800
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   10680
            TabIndex        =   115
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   78708739
            CurrentDate     =   40725
            MinDate         =   39083
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000003&
            Caption         =   "Profession"
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   107
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label31 
            BackColor       =   &H80000003&
            Caption         =   "Joined On"
            Height          =   375
            Left            =   7080
            TabIndex        =   106
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000003&
            Caption         =   "E mail"
            Height          =   375
            Index           =   1
            Left            =   7080
            TabIndex        =   105
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label30 
            BackColor       =   &H80000003&
            Caption         =   "Name"
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label29 
            BackColor       =   &H80000003&
            Caption         =   "Phone Number (1)"
            Height          =   375
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label28 
            BackColor       =   &H80000003&
            Caption         =   "Phone Number (2)"
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000003&
            Caption         =   "Phone Number (3)"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   1800
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmdealerinfo.frx":F2B8
         Height          =   2535
         Left            =   -74760
         TabIndex        =   116
         Top             =   3120
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   4471
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmdealerinfo.frx":F2D0
         Height          =   3735
         Left            =   -74760
         TabIndex        =   138
         Top             =   3360
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   6588
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
      Begin VB.Frame Frame3 
         Caption         =   "Edit/Change Payment Information"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74760
         TabIndex        =   42
         Top             =   3360
         Visible         =   0   'False
         Width           =   18735
         Begin VB.CommandButton cmdrefpinfo 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4200
            TabIndex        =   71
            Top             =   2640
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dcpayment 
            Bindings        =   "frmdealerinfo.frx":F2E6
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "P######"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   68
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "PAY_ID"
            Text            =   "DataCombo5"
         End
         Begin VB.CommandButton cmdcancelpayment 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2880
            TabIndex        =   67
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CommandButton cmdupdatepayment 
            Caption         =   "Update"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1560
            TabIndex        =   66
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            DataField       =   "ACCOUNT_NUMBER"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            DataField       =   "BANK"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   840
            Width           =   5175
         End
         Begin VB.TextBox txtpayment 
            Appearance      =   0  'Flat
            DataField       =   "AMOUNT_PAID"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdeditpay 
            Caption         =   "Edit Payment "
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   8.25
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   51
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            DataField       =   "PAY_ID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "P#####"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adolist"
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtdid 
            Appearance      =   0  'Flat
            DataField       =   "DID"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtdname 
            Appearance      =   0  'Flat
            DataField       =   "DEALER_NAME"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtbusname 
            Appearance      =   0  'Flat
            DataField       =   "BUSINESS_NAME"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            DataField       =   "CURRENT_DUE"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            DataField       =   "PREVIOUS_DUE"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            DataField       =   "DATE"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   2760
            Width           =   2175
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            DataField       =   "PREPARED_BY"
            DataSource      =   "adolist"
            Height          =   375
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   3240
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo dcdealer 
            Bindings        =   "frmdealerinfo.frx":F2FC
            Height          =   315
            Left            =   3240
            TabIndex        =   70
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DID"
            Text            =   "DataCombo5"
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000003&
            Caption         =   "BANK ACC. INFO"
            Height          =   375
            Left            =   6120
            TabIndex        =   65
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000003&
            Caption         =   "BANK"
            Height          =   375
            Left            =   6120
            TabIndex        =   64
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000003&
            Caption         =   "PAYMENT"
            Height          =   375
            Left            =   6120
            TabIndex        =   63
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000003&
            Caption         =   "PAYMENT ID"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000003&
            Caption         =   "DEALER ID"
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000003&
            Caption         =   "DEALER NAME"
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label21 
            BackColor       =   &H80000003&
            Caption         =   "BUSINESS NAME"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000003&
            Caption         =   "CURRENT DUE"
            Height          =   375
            Left            =   6120
            TabIndex        =   58
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000003&
            Caption         =   "PREVIOUS DUE"
            Height          =   375
            Left            =   6120
            TabIndex        =   57
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000003&
            Caption         =   "PREPARED BY"
            Height          =   375
            Left            =   6120
            TabIndex        =   56
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000003&
            Caption         =   "DATE"
            Height          =   375
            Left            =   6120
            TabIndex        =   55
            Top             =   2760
            Width           =   1455
         End
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000000&
         Caption         =   "Search by # DEALER NAME, DEALER ID, NATIONAL ID, TIN"
         Height          =   255
         Left            =   -65400
         TabIndex        =   132
         Top             =   6240
         Width           =   4695
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4575
         Left            =   -59160
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         DataField       =   "DEALER_NAME"
         DataSource      =   "adodealerlist"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   2
         Top             =   6240
         Width           =   6855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "DEALER  INFORMATION"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   127
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmdealerinfo.frx":F312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmdealerinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
On Error Resume Next

'initialising controls
    Text1.Locked = False
    txtt_due = ""
    DTPicker1.Enabled = True
    cmdadd.Enabled = False
    cmdupdt.Enabled = True
    cmdedit.Enabled = False
    cmdcncl.Enabled = True
    cmdrefsh.Enabled = False
    txt_name.Locked = False
    txtphn_1.Locked = False
    txtphn_2.Locked = False
    txtphn_3.Locked = False
    txtbus_name.Locked = False
    txtaddrs.Locked = False
    txtr_addrs.Locked = False
    txtsaf_dep.Locked = False
    txt_eml.Locked = False
    txtvot_id.Locked = False
    cmdbws.Enabled = True
    cmdclr.Enabled = True
    adodelr.Enabled = False
    txtt_tran.Locked = False
    txtt_paid.Locked = False
    txtt_due.Locked = False
    Text17.Locked = False

Dim varDID
If adodelr.Recordset.RecordCount = 0 Then
    adodelr.Recordset.AddNew
    txtdlr_id = "DID1005001"
Else: adodelr.Recordset.Sort = "DID ASC"
    adodelr.Recordset.MoveLast
    
    varDID = Mid(txtdlr_id.Text, 4, 10)
    adodelr.Recordset.AddNew
    txtdlr_id.Text = "DID" + CStr(varDID + 1)
End If

End Sub

Private Sub cmdcancelentry_Click()
    adoothers.Recordset.Cancel
    adoothers.Refresh
    cmdnew.Enabled = True
    cmdupdate.Enabled = False
    cmdcancelentry.Enabled = False
    cmdeditentry.Enabled = True
    DataCombo5.Enabled = True
    cmdrefresh.Enabled = True
    Frame5.Enabled = False
    MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"
    
End Sub

Private Sub cmdcancelpayment_Click()
cmdupdatepayment.Enabled = False
cmdeditpay.Enabled = True
dcpayment.Enabled = True
dcdealer.Visible = False
cmdrefpinfo.Enabled = True
cmdcancelpayment.Enabled = True
adodelr.Recordset.Cancel
adodelr.Refresh
MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"
End Sub

Private Sub cmdcncl_Click()
onprocess = False
'DataGrid1.Enabled = True
cmdbws.Enabled = False
Text1.Locked = True
cmdclr.Enabled = False
DTPicker1.Enabled = False
adodelr.Recordset.Cancel
adodelr.Refresh
cmdadd.Enabled = True
cmdupdt.Enabled = False
cmdedit.Enabled = True
cmdcncl.Enabled = False
cmdrefsh.Enabled = True
txt_name.Locked = True
txtphn_1.Locked = True
txtphn_2.Locked = True
txtphn_3.Locked = True
txtt_due = ""
txtbus_name.Locked = True
txtaddrs.Locked = True
txtr_addrs.Locked = True
txtsaf_dep.Locked = True
txt_eml.Locked = True
txtvot_id.Locked = True
cmdbws.Enabled = False
cmdclr.Enabled = False
adodelr.Enabled = True
txtt_tran.Locked = True
txtt_paid.Locked = True
Text17.Locked = True
txtt_due.Locked = True
MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"

End Sub

Private Sub cmdedit_Click()
'starting editing

    DTPicker1.Enabled = True
    Text1.Locked = False
    cmdadd.Enabled = False
    cmdupdt.Enabled = True
    cmdedit.Enabled = False
    cmdcncl.Enabled = True
    cmdrefsh.Enabled = False
    txt_name.Locked = False
    txtphn_1.Locked = False
    txtphn_2.Locked = False
    txtphn_3.Locked = False
    txtbus_name.Locked = False
    txtaddrs.Locked = False
    txtr_addrs.Locked = False
    txtsaf_dep.Locked = False
    txt_eml.Locked = False
    txtvot_id.Locked = False
    cmdbws.Enabled = True
    cmdclr.Enabled = True
    adodelr.Enabled = False
    'txtt_tran.Locked = False
    'txtt_paid.Locked = False
    'txtt_due.Locked = False
    Text17.Locked = False
End Sub

Private Sub cmdeditentry_Click()
Frame5.Enabled = True
DataCombo5.Enabled = False
cmdeditentry.Enabled = False
MsgBox "You can edit", vbInformation + vbOKOnly, "System Informtaion"
DataCombo5.Enabled = False
cmdnew.Enabled = False
cmdupdate.Enabled = True
cmdrefresh.Enabled = False
cmdcancelentry.Enabled = True
Frame5.Enabled = True
End Sub

Private Sub cmdeditpay_Click()
    dcpayment.Enabled = False
    dcdealer.Visible = True
    cmdcancelpayment.Enabled = True
    cmdupdatepayment.Enabled = True
    cmdeditpay.Enabled = False
    cmdrefpinfo.Enabled = False
End Sub

Private Sub cmdnew_Click()
adoothers.Recordset.AddNew
MsgBox "You can add another person now", vbInformation + vbOKOnly, "System Informtaion"
Frame5.Enabled = True
DataCombo5.Enabled = False
cmdupdate.Enabled = True
cmdnew.Enabled = False
cmdcancelentry.Enabled = True
cmdeditentry.Enabled = False
cmdrefresh.Enabled = False
Frame5.Enabled = True
End Sub


Private Sub cmdref_Click()
adopayment.Refresh
adodealerpayment.Refresh
DataCombo3.ReFill
DataCombo3.Refresh
DataCombo4.ReFill
DataCombo4.Refresh
adolist.Refresh
adolist.Recordset.Sort = "PAY_ID DESC"
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"


End Sub

Private Sub cmdrefpinfo_Click()
adolist.Refresh
adodelr.Refresh
DataGrid3.Refresh
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Information"
End Sub

Private Sub cmdrefresh_Click()
adoothers.Refresh
DataCombo5.Refresh
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"
End Sub

Private Sub cmdrefsh_Click()
adodelr.Refresh
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"

End Sub

Private Sub cmdupdate_Click()
On Error GoTo errorhandler
If Text13 = "" Then
MsgBox "Please fill up the fields and try again", vbCritical + vbOKOnly, "Database Error"
Else
cmdnew.Enabled = True
cmdupdate.Enabled = False
cmdcancelentry.Enabled = False
cmdeditentry.Enabled = True
DataCombo5.Enabled = True
cmdrefresh.Enabled = True
Frame5.Enabled = False
adoothers.Recordset.Update
MsgBox "Database updated", vbInformation + vbOKOnly, "System Information"
End If

errorhandler:
If Err <> 0 Then
 MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub cmdupdatepayment_Click()
On Error Resume Next
cmdrefpinfo.Enabled = True
cmdcancelpayment.Enabled = False
dcpayment.Visible = True
dcpayment.Enabled = True
cmdeditpay.Enabled = True
dcdealer.Visible = False
cmdrefpinfo.Enabled = True
cmdcancelpayment.Enabled = True
cmdupdatepayment.Enabled = False
adolist.Recordset.Update
adolist.Recordset.Save
adolist.Refresh
MsgBox "Database updated", vbInformation + vbOKOnly, "System Information"
cmdrefpinfo_Click
End Sub

Private Sub cmdupdt_Click()
  
If txt_name = "" Or txtphn_1 = "" Or Text1 = "" Or Text17 = "" Or txtbus_name = "" Or txtsaf_dep = "" Then
MsgBox "Please complete all the fields and try again", vbExclamation, "System Error"

Else
'uninitialising controls
    Text1.Locked = True
    cmdadd.Enabled = True
    cmdupdt.Enabled = False
    cmdedit.Enabled = True
    cmdcncl.Enabled = False
    cmdrefsh.Enabled = True
    cmdbws.Enabled = False
    cmdclr.Enabled = False
    DTPicker1.Enabled = False
    txt_name.Locked = True
    txtphn_1.Locked = True
    txtphn_2.Locked = True
    txtphn_3.Locked = True
    txtbus_name.Locked = True
    txtaddrs.Locked = True
    txtr_addrs.Locked = True
    txtsaf_dep.Locked = True
    txt_eml.Locked = True
    txtt_due = ""
    txtvot_id.Locked = True
    adodelr.Enabled = True
    txtt_tran.Locked = True
    txtt_paid.Locked = True
    txtt_due.Locked = True
    Text17.Locked = True
    
    'updating database
    adodelr.Recordset.Fields(6) = Val(txtt_due)
    adodelr.Recordset.Update
    
    

MsgBox "Database updated", vbInformation + vbOKOnly, "System Informtaion"
End If
End Sub

Private Sub DataList1_Click()
adodelr.Recordset.Bookmark = DataList1.SelectedItem
End Sub

Private Sub cmdview_o_list_Click()
If dev_bike.rscmdotherslist.State = adStateOpen Then
  dev_bike.rscmdotherslist.Close
 End If
rptotherlist.Show
End Sub

Private Sub Command1_Click()
If dev_bike.rscmdDEALERbyDID.State = adStateOpen Then
    dev_bike.rscmdDEALERbyDID.Close
End If
    dev_bike.cmdDEALERbyDID Trim(Text5.Text)

rptDelrinfo.Show

End Sub

Private Sub Command10_Click()
If dev_bike.rscmdall_dealer.State = adStateOpen Then
    dev_bike.rscmdall_dealer.Close
    End If
    'dev_bike.cmdall_dealer Trim(txtsearch.Text)
    rptalldealer.Show
End Sub

Private Sub Command11_Click()
If dev_bike.rscmddues.State = adStateOpen Then
    dev_bike.rscmddues.Close
End If
    'dev_bike.cmddues Trim(txtsearch.Text)
    rptdealer_dues.Show
End Sub

Private Sub Command12_Click()
On Error GoTo errorhandler
If Text19 = "" Then
    MsgBox "Please include a search term and try again.", vbExclamation, "Borac Sales System"
Else
    DataGrid2.BackColor = vbGreen
If Command12.Caption = "Clear Search" Then
   adodealerlist.CommandType = adCmdText
   adodealerlist.RecordSource = "SELECT DID,DEALER_NAME,AREA,DEPOSIT_AMOUNT,TOTAL_AMOUNT,TOTAL_PAID,BALANCE,IMG_LOC FROM dealer"
   adodealerlist.Refresh
   DataGrid2.Columns(7).Visible = False
   Command12.Caption = "SEARCH"
   Text19 = ""
   DataGrid2.BackColor = &H80000005
ElseIf Command12.Caption = "SEARCH" Then
    adodealerlist.CommandType = adCmdText
    adodealerlist.RecordSource = "SELECT *FROM dealer WHERE (((dealer.DID)  Like '" & Me.Text19.Text & "%'))"
    adodealerlist.Refresh
    If adodealerlist.Recordset.RecordCount = 0 Then
        adodealerlist.CommandType = adCmdText
        adodealerlist.RecordSource = "SELECT *FROM dealer WHERE (((dealer.DEALER_NAME)  Like '" & Me.Text19.Text & "%'))"
        adodealerlist.Refresh
    ElseIf adodealerlist.Recordset.RecordCount = 0 Then
        adodealerlist.CommandType = adCmdText
        adodealerlist.RecordSource = "SELECT *FROM dealer WHERE (((dealer.TIN)  Like '" & Me.Text19.Text & "%'))"
        adodealerlist.Refresh
    ElseIf adodealerlist.Recordset.RecordCount = 0 Then
        adodealerlist.CommandType = adCmdText
        adodealerlist.RecordSource = "SELECT *FROM dealer WHERE (((dealer.NATIONAL_ID)  Like '" & Me.Text19.Text & "%'))"
        adodealerlist.Refresh
    End If
    If adodealerlist.Recordset.RecordCount = 0 Then
        MsgBox "No records matched.", vbOKOnly + vbInformation, "System Information"
    Else
        MsgBox adodealerlist.Recordset.RecordCount & " " & "records found.", vbOKOnly + vbInformation, "System Information"
    End If
    
    Command12.Caption = "Clear Search"
    
End If

End If
errorhandler:
If Err.Number <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub Command2_Click()
If dev_bike.rscmdCKDtranbyDID.State = adStateOpen Then
    dev_bike.rscmdCKDtranbyDID.Close
End If
    
    dev_bike.cmdCKDtranbyDID Trim(adodealerlist.Recordset.Fields(0).Value)
    rptckdtranhistory.Show
End Sub

Private Sub Command3_Click()
If dev_bike.rscmddealerorderhistorybyDID.State = adStateOpen Then
    dev_bike.rscmddealerorderhistorybyDID.Close
End If
    dev_bike.cmddealerorderhistorybyDID Trim(Text5.Text)
    rptdelrordrhis.Show
End Sub

Private Sub Command4_Click()
       ask = MsgBox("Do you want to add payment to this dealer - " & " " & Label14.Caption & "" & "?", vbQuestion + vbYesNo, "System query")
            If ask = vbYes Then
                User = Label14.Caption
                txtpayment.Enabled = True
                payment = InputBox("Input Payment")
                    If IsNumeric(payment) = False Then
                        MsgBox "Please input numeric value and try again ", vbExclamation + vbOKOnly, "System Error"
                    Else
                        bank = InputBox("Input Bank name")
                        accno = InputBox("Input Account no")
                        Dim pdue
                        Dim cdue
                        pdue = txtdue
                        cdue = txtdue - payment
                        adodealerpayment.Recordset.AddNew
                        adodealerpayment.Recordset.Fields(1) = adopayment.Recordset.Fields(0).Value
                        adodealerpayment.Recordset.Fields(2) = adopayment.Recordset.Fields(1).Value
                        adodealerpayment.Recordset.Fields(3) = adopayment.Recordset.Fields(14).Value
                        adodealerpayment.Recordset.Fields(4) = bank
                        adodealerpayment.Recordset.Fields(5) = accno
                        adodealerpayment.Recordset.Fields(6) = payment
                        adodealerpayment.Recordset.Fields(7) = pdue
                        adodealerpayment.Recordset.Fields(8) = cdue
                        adodealerpayment.Recordset.Fields(9) = Format(Date, "dd/MM/yyyy")
                        adodealerpayment.Recordset.Fields(10) = frmmain.stsbr_main.Panels(2).Text
                        adopayment.Recordset.Fields(5) = adopayment.Recordset.Fields(5).Value + payment
                        adopayment.Recordset.Fields(6) = adopayment.Recordset.Fields(4).Value - adopayment.Recordset.Fields(5).Value
                        adodealerpayment.Recordset.Save
                        adopayment.Recordset.Fields(6) = cdue
                        adopayment.Recordset.Update
                        MsgBox "Dealer - " & " " & User & " has current due " & cdue, vbinf + vbOKOnly, "System Information"
                        adodealerpayment.Refresh
                        adopayment.Refresh
                        adopayment.Refresh
                        
                    End If
             End If
     
End Sub

Private Sub Command5_Click()
    If dev_bike.rscmdDEALERpaybck.State = adStateOpen Then
        dev_bike.rscmdDEALERpaybck.Close
    End If
        dev_bike.cmdDEALERpaybck Trim(Text16.Text)
        rptdealerpayback.Show

End Sub

Private Sub Command6_Click()
adodealerlist.Refresh
DataGrid2.Columns(7).Visible = False
MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"

End Sub

Private Sub Command7_Click()
If Combo3.Text = "BY DID" And txtsearch.Text <> "" Then
    If dev_bike.rscmdDEALERbyDID.State = adStateOpen Then
    dev_bike.rscmdDEALERbyDID.Close
    End If
    dev_bike.cmdDEALERbyDID Trim(txtsearch.Text)
    rptDelrinfo.Show

ElseIf Combo3.Text = "BY NAME" And txtsearch.Text <> "" Then
    If dev_bike.rscmdDEALERbyNAME.State = adStateOpen Then
    dev_bike.rscmdDEALERbyNAME.Close
    End If
    dev_bike.cmdDEALERbyNAME Trim(txtsearch.Text)
    rptDelrinfobyname.Show
Else
  MsgBox "Operation is cancelled because you didnt input correct parameters ", vbCritical + vbOKOnly, "Input Error"
End If

End Sub

Private Sub DataCombo1_Change()
adodealerlist.Recordset.Bookmark = DataCombo1.SelectedItem
End Sub

Private Sub Command8_Click()
If dev_bike.rscmdBIKEtranbyDID.State = adStateOpen Then
    dev_bike.rscmdBIKEtranbyDID.Close
End If
    
    dev_bike.cmdBIKEtranbyDID Trim(adodealerlist.Recordset.Fields(0).Value)
    rptbiketranhistory.Show
End Sub

Private Sub Command9_Click()
If dev_bike.rscmdSPAREtranbyDID.State = adStateOpen Then
    dev_bike.rscmdSPAREtranbyDID.Close
End If
    
    dev_bike.cmdSPAREtranbyDID Trim(adodealerlist.Recordset.Fields(0).Value)
    rptsparestranhistory.Show
End Sub

Private Sub DataCombo3_Change()
On Error Resume Next
adopayment.Recordset.Bookmark = DataCombo3.SelectedItem

End Sub

Private Sub DataCombo3_DblClick(Area As Integer)
On Error Resume Next
adopayment.Recordset.Bookmark = DataCombo3.SelectedItem
End Sub

Private Sub DataCombo4_Change()
adopayment.Recordset.Bookmark = DataCombo4.SelectedItem
End Sub

Private Sub DataCombo4_DblClick(Area As Integer)
On Error Resume Next
adopayment.Recordset.Bookmark = DataCombo4.SelectedItem
End Sub

Private Sub DataCombo5_Change()
On Error Resume Next
adoothers.Recordset.Bookmark = DataCombo5.SelectedItem
End Sub

Private Sub dcdealer_Change()
On Error Resume Next
adodelr.Recordset.Bookmark = dcdealer.SelectedItem
txtdid = adodelr.Recordset.Fields(0).Value
txtdname = adodelr.Recordset.Fields(1).Value
txtbusname = adodelr.Recordset.Fields(14).Value
End Sub

Private Sub dcdealer_DblClick(Area As Integer)
On Error Resume Next
adodelr.Recordset.Bookmark = dcdealer.SelectedItem
txtdid = adodelr.Recordset.Fields(0).Value
txtdname = adodelr.Recordset.Fields(1).Value
txtbusname = adodelr.Recordset.Fields(14).Value
End Sub

Private Sub dcpayment_Change()
adolist.Recordset.Bookmark = dcpayment.SelectedItem
End Sub

Private Sub DTPicker1_Change()
txt_dob.Text = Format(DTPicker1.Value, "dd / MM / yyyy")
End Sub

Private Sub DTPicker2_Change()
Text15 = DTPicker2.Value
End Sub

Private Sub Form_Load()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmdealerinfo.Width - framex.Width) / 2
framex.Left = fpostn
DataGrid2.Columns(7).Visible = False
cmdupdt.Enabled = False
cmdcncl.Enabled = False
adodealerlist.Visible = True
adodealerlist.Left = 0
DTPicker1.Value = Date
adolist.Recordset.Sort = "PAY_ID DESC"
End Sub

Private Sub Form_Resize()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmdealerinfo.Width - framex.Width) / 2
framex.Left = fpostn
End Sub

Private Sub cmdbws_Click()
    With diagpic
        .CancelError = False
        .Filter = "Picture"
        .ShowOpen
     End With
    txtimg_addrs.Text = diagpic.FileName
        If txtimg_addrs.Text = "" Then
            MsgBox " No picture file is selected"
        End If
                
End Sub

Private Sub cmdclr_Click()
    img_delr.Picture = LoadPicture()
    txtimg_addrs.Text = ""
End Sub

Private Sub Frame2_Click()
    'ext4.Text = "Search by dealer's name"
    'Text4.FontItalic = True
End Sub

Private Sub Text12_Change()
If IsNumeric(Text12) = False Then
        Text12.Text = ""
End If
End Sub

Private Sub Text4_Click()
    If Text4.Text = "Search by dealer's name" Then
        Text4.Text = ""
    End If
    Text4.FontItalic = False
End Sub

Private Sub Timer1_Timer()
    If Text4.Text = "" Then
       Text4.Text = "Search by dealer's name"
    End If
End Sub

Private Sub Text4_LostFocus()
    'Text4.Text = "Search by dealer's name"
    'Text4.FontItalic = True
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Label2_Change(Index As Integer)
On Error GoTo picerror
    If adodealerlist.Recordset.EOF = False And adodealerlist.Recordset.BOF = False Then
    Dim LOC As Variant
    'LOC = adodealerlist.Recordset.Fields(16).Value
    'MsgBox adodealerlist.Recordset.Fields(16).Value
    'MsgBox txtloc.Text
        If txtloc = "" Then
            Image3.Picture = LoadPicture("C:\dbase_bike\pics\noimage.jpg")
        Else
            Image3.Picture = LoadPicture(txtloc.Text)
        End If
    End If
picerror:
    If Err.Number = 53 Then
        Image3.Picture = LoadPicture("C:\dbase_bike\pics\noimage.jpg")
    End If
End Sub

Private Sub Text6_Change()
If IsNumeric(Text6) = False Then
        Text6.Text = ""
End If
End Sub

Private Sub Text7_Change()
If IsNumeric(Text7) = False Then
        Text7.Text = ""
End If
End Sub

Private Sub txt_name_KeyUp(KeyCode As Integer, Shift As Integer)
If IsNumeric(txt_name) = True Then
        txt_name.Text = ""
End If
End Sub

Private Sub txtimg_addrs_Change()
On Error GoTo errorhandler
    img_delr.Picture = LoadPicture()
    img_delr.Picture = LoadPicture(txtimg_addrs.Text)
errorhandler:
If Err.Number = 481 Then
MsgBox "Sorry you have chosen an invalid picture format. Try common files types like JPEG.", vbExclamation + vbOKOnly, "Invalid Picture Type"
End If
End Sub

Private Sub txtphn_1_Change()
If IsNumeric(txtphn_1) = False Then
        txtphn_1.Text = ""
End If
End Sub

Private Sub txtphn_2_Change()
If IsNumeric(txtphn_2) = False Then
        txtphn_2.Text = ""
End If
End Sub

Private Sub txtphn_3_Change()
If IsNumeric(txtphn_3) = False Then
       txtphn_3.Text = ""
End If
End Sub

Private Sub txtsaf_dep_Change()
If IsNumeric(txtsaf_dep) = False Then
        txtsaf_dep.Text = ""
End If
End Sub

Private Sub txtt_due_KeyUp(KeyCode As Integer, Shift As Integer)
'If IsNumeric(txtt_due) = False Then
       'txtt_due.Text = ""
'End If
End Sub

Private Sub txtt_paid_Change()
'If IsNumeric(txtt_paid) = False Then
       'txtt_paid.Text = 0

'End If
End Sub

Private Sub txtt_paid_KeyUp(KeyCode As Integer, Shift As Integer)
'If IsNumeric(txtt_paid) = False Then
       'txtt_paid.Text = ""
'Else
'txtt_due = Val(txtt_tran) - Val(txtt_paid)
'End If
End Sub

Private Sub txtt_tran_Change()
'If IsNumeric(txtt_tran) = False Then
       'txtt_tran.Text = 0

'End If
End Sub

Private Sub txtt_tran_KeyUp(KeyCode As Integer, Shift As Integer)
'If IsNumeric(txtt_tran) = False Then
       'txtt_tran.Text = ""
'Else
'txtt_due = Val(txtt_tran) - Val(txtt_paid)
'End If
End Sub
