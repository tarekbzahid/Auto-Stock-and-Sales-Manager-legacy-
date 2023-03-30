VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbike_info 
   Caption         =   "BIKE Info"
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
   Icon            =   "frmbike_info.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame framex 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   17535
      Begin VB.Frame Frame9 
         Caption         =   "BIKE TRANSACTION BETWEEN DATES"
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
         Left            =   6000
         TabIndex        =   25
         Top             =   7200
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "REPORT"
            Height          =   375
            Left            =   4080
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   375
            Left            =   240
            TabIndex        =   27
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
            CustomFormat    =   "d/MM/yyyy"
            Format          =   78839811
            CurrentDate     =   40805
         End
         Begin MSComCtl2.DTPicker fromdate 
            Height          =   375
            Left            =   2160
            TabIndex        =   28
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
            CustomFormat    =   "d/MM/yyyy"
            Format          =   78839811
            CurrentDate     =   40805
         End
         Begin VB.Label Label18 
            Height          =   255
            Left            =   2160
            TabIndex        =   30
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label19 
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame stockbal 
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
         Left            =   10440
         TabIndex        =   19
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txtquan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "LAST_ADDED_QUANTITY"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1680
            TabIndex        =   21
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtslvl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            DataField       =   "BIKE_LEVEL"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            Caption         =   "Add Quantity"
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
            TabIndex        =   24
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label llast 
            BackColor       =   &H80000000&
            Caption         =   "Last Added "
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
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000000&
            Caption         =   "Current Stock Level"
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
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Assembled Bike Info"
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
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10215
         Begin VB.TextBox txtcp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "COST_PRICE"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   5040
            TabIndex        =   39
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Frame ckdinfo 
            Caption         =   "CKD Info"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   6720
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   3375
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               DataField       =   "COST_PRICE"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   1800
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               DataField       =   "BIKE_LEVEL"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               DataField       =   "UNIT_PRICE"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               DataField       =   "BIKE_DETAILS"
               DataSource      =   "Adodc1"
               Height          =   375
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label11 
               BackColor       =   &H80000000&
               Caption         =   "CKD Cost Price"
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
               TabIndex        =   41
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000000&
               Caption         =   "CKD Level"
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
               TabIndex        =   37
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000000&
               Caption         =   "CKD Unit Price"
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
               TabIndex        =   36
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label9 
               BackColor       =   &H80000000&
               Caption         =   "Name"
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
               TabIndex        =   35
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.TextBox txtbid 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "BID"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "BIKE_DETAILS"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtup 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "UNIT_PRICE"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            TabIndex        =   10
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtdate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "LAST_ADDED_DATE"
            DataSource      =   "adospares_info"
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmbike_info.frx":F172
            Height          =   315
            Left            =   3240
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "BID"
            Text            =   "CKD"
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
            Caption         =   "Bike Cost Price "
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
            Left            =   3360
            TabIndex        =   38
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000000&
            Caption         =   "Bike ID"
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
            TabIndex        =   17
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000000&
            Caption         =   "Details / Name"
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
            TabIndex        =   16
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000000&
            Caption         =   "Bike Unit Price "
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
            TabIndex        =   15
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000000&
            Caption         =   "Stock Added Date"
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
            TabIndex        =   14
            Top             =   1320
            Width           =   1695
         End
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
         TabIndex        =   7
         Top             =   360
         Width           =   1935
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
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdaddstock 
         Caption         =   "ADD CKD >> BIKE"
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
         TabIndex        =   5
         Top             =   1680
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
         Left            =   14520
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
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
         TabIndex        =   3
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BIKE Report"
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
         TabIndex        =   2
         Top             =   3480
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   10560
         Top             =   2640
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
      Begin MSAdodcLib.Adodc adospares_info 
         Height          =   495
         Left            =   120
         Top             =   7200
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
         RecordSource    =   "stock_assembled"
         Caption         =   "BIKE Information"
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
         Bindings        =   "frmbike_info.frx":F187
         Height          =   4095
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   7223
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "BIKE INFORMATION"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
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
      Picture         =   "frmbike_info.frx":F1A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmbike_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
adospares_info.Refresh
    If dev_bike.rscmdinvbike.State = adStateOpen Then
        dev_bike.rscmdinvbike.Close
        
    End If
rptallbike.Show
End Sub

Private Sub Command3_Click()
Dim prmfrom
Dim prmto
prmfrom = fromdate.Value
prmto = todate.Value
If dev_bike.rscmdBIKETRANBETWEENDATES.State = adStateOpen Then
  dev_bike.rscmdBIKETRANBETWEENDATES.Close
 End If
 dev_bike.cmdBIKETRANBETWEENDATES prmfrom, prmto
rptbiketranbetwbdates.Show
End Sub

Private Sub DataCombo1_Change()
Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
txtbid = DataCombo1.Text
txtname = Text3
ckdinfo.Visible = True
End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next
Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
txtbid = DataCombo1.Text
Option1.Value = False

Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
txtbid = DataCombo1.Text
txtname = Text3
ckdinfo.Visible = True

End Sub

Private Sub DataCombo2_Change()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "SELECT *FROM stock_ckd WHERE (((stock_ckd.BID) Like '" & DataCombo2.Text & "%'))"
Adodc1.Refresh
MsgBox Adodc1.Recordset.RecordCount
End Sub



Private Sub DataGrid1_Scroll(Cancel As Integer)
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = " SELECT *FROM stock_ckd WHERE (((stock_ckd.BID)  Like '" & txtbid.Text & "%'))"
Adodc1.Refresh
MsgBox Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmbike_info.Width - framex.Width) / 2
framex.Left = fpostn
todate = Date
fromdate = Date

End Sub

Private Sub Form_Resize()
Image2.Width = Me.Width
Dim fpostn As Long
fpostn = (frmbike_info.Width - framex.Width) / 2
framex.Left = fpostn
End Sub

Private Sub cmdadd_Click()
'from herecmd butn changes

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = " SELECT *FROM stock_ckd"
Adodc1.Refresh


Adodc1.Refresh
Adodc1.Refresh
DataCombo1.ReFill
cmdadd.Enabled = False
adospares_info.Enabled = True
cmdref.Enabled = False
cmdaddstock.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
stockbal.Enabled = True
Frame2.Enabled = True
adospares_info.Recordset.AddNew
txtdate = Format(Date, "dd/MM/yyyy")
txtslvl = "00"
adospares_info.Enabled = False
Label7.Visible = True
DataCombo1.Visible = True
DataGrid1.Enabled = False
End Sub

Private Sub cmdaddstock_Click()
On Error GoTo errorhandler


'from herecmd butn changes
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = " SELECT *FROM stock_ckd WHERE (((stock_ckd.BID)  Like '" & txtbid.Text & "%'))"
Adodc1.Refresh

stockadd = InputBox("Type in numbers of stock you want to add to the BIKE " & txtname.Text & ". You have " & Adodc1.Recordset.Fields(2).Value & " CKDs left", "Add BIKE Stock", "1000")
 
 If IsNumeric(stockadd) = False Then
    MsgBox "Please input a numeric value and try again", vbCritical + vbOKOnly, "System Error"
 ElseIf Adodc1.Recordset.Fields(2).Value < Val(stockadd) Then
    MsgBox "You cannot assemble more CKDs than you already have!", vbCritical + vbOKOnly, "System Error"
 Else
    stockadd = Int(stockadd)
    If stockadd < 0 Then
        stockadd = -stockadd
    End If
    ask = MsgBox("Are you sure you want to add this to the stock?", vbInformation + vbYesNo, "System Information")
        If ask = vbYes Then
            
            txtquan = stockadd
            txtslvl = Val(txtquan) + Val(txtslvl)
            adospares_info.Recordset.Fields(4) = Format(Date, "dd/MM/yyyy")
            adospares_info.Recordset.Update
            adospares_info.Refresh
            adospares_info.Refresh
            
            Adodc1.Recordset.Fields(2) = Adodc1.Recordset.Fields(2).Value - Val(stockadd)
            Adodc1.Recordset.Update
            Adodc1.Refresh
            
            
            MsgBox "Stock added to the database", vbInformation + vbOKOnly, "System Informtaion"
        ElseIf ask = vbNo Then
            adospares_info.Recordset.Cancel
            MsgBox "Stock NOT added to the database", vbInformation + vbOKOnly, "System Informtaion"
        End If
   



End If
errorhandler:
If Err.Number <> 0 Then

If Err.Number = -2147217864 Then
    MsgBox "Please refresh the database and try again", vbCritical + vbOKOnly, "System Error"
Else
MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End If
End Sub

Private Sub cmdcancel_Click()
adospares_info.Recordset.Cancel
cmdadd.Enabled = True
cmdref.Enabled = True
cmdsave.Enabled = False
cmdaddstock.Enabled = True
adospares_info.Refresh
'ladd.Visible = False
cmdcancel.Enabled = False

DataCombo1.Visible = False


ckdinfo.Visible = False
adospares_info.Enabled = True
DataGrid1.Enabled = True
Label7.Visible = False
Frame2.Enabled = False
stockbal.Enabled = False

MsgBox "Process Cancelled", vbInformation + vbOKOnly, "System Information"


End Sub

Private Sub cmdref_Click()
adospares_info.Refresh
Adodc1.Refresh
Adodc1.Refresh

MsgBox "Database refreshed", vbInformation + vbOKOnly, "System Informtaion"
End Sub

Private Sub cmdsave_Click()
On Error GoTo erroronproduction
If txtbid = "" Then
    MsgBox "Please input Bike ID", vbExclamation + vbOKOnly, "System Error"
    txtbid.SetFocus
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
    MsgBox "Cost price cannot be greater than unit price !", vbExclamation + vbOKOnly, "System Error"
ElseIf Val(txtquan) > Adodc1.Recordset.Fields(2).Value Then
    MsgBox "Sorry you cant assemble more than what you already have !", vbExclamation + vbOKOnly, "System Error"
   
  
Else:
    txtslvl = Val(txtquan) + Val(txtslvl)
    Adodc1.Recordset.Fields(2) = Adodc1.Recordset.Fields(2).Value - Val(txtquan)
    Adodc1.Recordset.Update
    Adodc1.Refresh
    Adodc1.Refresh
   
    adospares_info.Recordset.Save
    adospares_info.Refresh
    
 
    
    
Label7.Visible = False
Frame2.Enabled = False
cmdadd.Enabled = True
cmdref.Enabled = True
cmdsave.Enabled = False
cmdaddstock.Enabled = True
adospares_info.Refresh

ckdinfo.Visible = False
cmdcancel.Enabled = False

Label6.Visible = False
Text2.Visible = False
DataCombo1.Visible = False
Label4.Visible = False
Text1.Visible = False
DataGrid1.Enabled = True
DataGrid1.Enabled = True
Label7.Visible = False
DataCombo1.Visible = False

ckdinfo.Visible = False
DataCombo1.Visible = False
cmdcancel.Enabled = False
DataGrid1.Enabled = True
adospares_info.Enabled = True
stockbal.Enabled = False
ckdinfo.Visible = False
stockbal.Enabled = False

MsgBox "Database updated", vbInformation + vbOKOnly, "System Informtaion"

End If
erroronproduction:
    If Err.Number = -2147217842 Then
        MsgBox "Operation is cancelled because you already assembled this bike before. To assemble more CKDs please click on ADD CKD >> BIKE. ", vbCritical + vbOKOnly, "Database Error"
        cmdcancel_Click
    ElseIf Err.Number <> 0 Then
        MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
        cmdcancel_Click
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

Private Sub Option1_Click()
txtbid.Text = "<OTHER>"
End Sub

Private Sub txtquan_KeyUp(KeyCode As Integer, Shift As Integer)

If IsNumeric(txtquan) = False Then
        txtquan.Text = ""
ElseIf Val(txtquan) > Val(Text1) Then
    MsgBox "Sorry you cant assemble more than what you already have !", vbExclamation + vbOKOnly, "System Error"
    txtquan.Text = ""
End If

End Sub

Private Sub txtup_KeyUp(KeyCode As Integer, Shift As Integer)
If IsNumeric(txtup) = False Then
        txtup.Text = ""
'ElseIf Val(txtup) < Val(Text2) Then
'MsgBox "Do you really want to price your bike lower than CKD ?", vbExclamation + vbOKOnly, "System Information"
End If
End Sub
