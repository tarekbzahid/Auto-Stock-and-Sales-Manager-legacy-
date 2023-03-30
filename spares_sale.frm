VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmspares_sales 
   Caption         =   "SPARES SALE"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "spares_sale.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   18585
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optretail 
      BackColor       =   &H80000000&
      Caption         =   "RETAIL"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdinvoice 
      Caption         =   "INVOICE"
      Height          =   735
      Left            =   12360
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "PRODUCT"
      Enabled         =   0   'False
      Height          =   3735
      Left            =   240
      TabIndex        =   35
      Top             =   5880
      Width           =   11895
      Begin VB.TextBox txtperts_des 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "BIKE_DETAILS"
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtprts_id 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "BID"
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
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         Height          =   615
         Left            =   9720
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "UNIT_PRICE"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
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
         Left            =   6600
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   8160
         TabIndex        =   40
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdadd_remove 
         Caption         =   "ADD TO CART"
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
         Height          =   495
         Left            =   9720
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtsearch 
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
         Height          =   735
         Left            =   9720
         MultiLine       =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "spares_sale.frx":F172
         Left            =   9720
         List            =   "spares_sale.frx":F17C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1935
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "spares_sale.frx":F18E
         Height          =   2400
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4233
         _Version        =   393216
         Appearance      =   0
         ListField       =   "STOCK_DETAILS"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   50
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   49
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   8160
         TabIndex        =   48
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6600
         TabIndex        =   47
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000000&
         Caption         =   "MODE OF PAYMENT"
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
         Left            =   9720
         TabIndex        =   46
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFORMATION"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   15000
      TabIndex        =   7
      Top             =   720
      Width           =   4695
      Begin VB.Timer timer_process 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4560
         Top             =   480
      End
      Begin VB.TextBox txtsub_total 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtttotal 
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
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtinv_no 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataSource      =   "adoinv"
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox txtinv_date 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataSource      =   "adoinv"
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox txtdiscnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Text            =   "0"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtothr_chrg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtpaid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Text            =   "0"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtdue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtdebt 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7800
         Width           =   2175
      End
      Begin VB.TextBox txtpay 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2175
      End
      Begin VB.TextBox txttottran 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6840
         Width           =   2175
      End
      Begin VB.TextBox txtbname 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   6240
         Width           =   4215
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5520
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "TOTAL / BDT"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "SUB TOTAL / BDT"
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
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "TRANSACTION NO #"
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
         Left            =   240
         TabIndex        =   31
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "INVOICE DATE"
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
         Left            =   240
         TabIndex        =   30
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Discount / BDT"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Other charges / BDT"
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
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Paid / BDT"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Due / BDT"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "TOT. DEBT / BDT"
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
         Left            =   240
         TabIndex        =   25
         Top             =   7800
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "TOT. PAYMENT / BDT"
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
         Left            =   240
         TabIndex        =   24
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "TOT.TRANSACTION / BDT"
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
         Left            =   240
         TabIndex        =   23
         Top             =   6840
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   6000
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "CUSTOMER NAME"
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
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   5280
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "START TRANSACTION"
      Height          =   855
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCEL "
      Enabled         =   0   'False
      Height          =   855
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6840
      Width           =   2415
   End
   Begin VB.OptionButton optdealer 
      BackColor       =   &H80000000&
      Caption         =   "DEALER"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.OptionButton optother 
      BackColor       =   &H80000000&
      Caption         =   "OTHER"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtterms 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "ALL GOODS RECEIVED IN GOOD CONDITION"
      Top             =   9360
      Width           =   7335
   End
   Begin MSDataListLib.DataCombo dtprm 
      Bindings        =   "spares_sale.frx":F1A5
      Height          =   315
      Left            =   12360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "STID"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   "Customer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc adostock 
      Height          =   330
      Left            =   6840
      Top             =   120
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   240
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   840
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
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
   Begin MSAdodcLib.Adodc adoinvoice 
      Height          =   330
      Left            =   8040
      Top             =   120
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
      RecordSource    =   "invoice_spares"
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
   Begin MSAdodcLib.Adodc adoinvoice_details 
      Height          =   330
      Left            =   9240
      Top             =   120
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "assembled_details"
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
   Begin MSAdodcLib.Adodc adocustomer 
      Height          =   330
      Left            =   10440
      Top             =   120
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SPARES SALES"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   55
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "spares_sale.frx":F1BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   54
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000000&
      Caption         =   "SELL TO:"
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
      Left            =   240
      TabIndex        =   53
      Top             =   5400
      Width           =   7695
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000000&
      Caption         =   "SPECIAL NOTE"
      Height          =   255
      Left            =   12360
      TabIndex        =   52
      Top             =   9120
      Width           =   2535
   End
End
Attribute VB_Name = "frmspares_sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset

Private Sub cmdadd_remove_Click()
On Error GoTo errorhandler
    If cmdadd_remove.Caption = "ADD TO CART" Then
    
        If txtqty = 0 Then
            MsgBox "Type-in a quantity other than zero !", vbOKOnly + vbInformation, "System Error"
        ElseIf txtqty = "" And txtprts_id = "" Then
             MsgBox "First select product and type-in quantity then try adding. ", vbExclamation, "System Error"
        ElseIf adostock.Recordset.Fields(3).Value = 0 Then
            MsgBox "No stock left ! ", vbExclamation, "System Error"
        Else
                
                Dim TCP As Long
                Dim TOT_COST As Long
                Dim NSTOCK As Integer
                TCP = Val(adostock.Recordset.Fields(8).Value) * Val(txtqty)
                TOT_COST = Val(adostock.Recordset.Fields(7).Value) * Val(txtqty)
                rsTemp.AddNew
                    rsTemp.Fields("PARTS ID") = txtprts_id.Text
                    rsTemp.Fields("PARTS DESCRIPTION") = txtperts_des.Text
                    rsTemp.Fields("QTY") = Val(txtqty.Text)
                    rsTemp.Fields("UNIT PRICE") = txtup.Text
                    rsTemp.Fields("TOTAL COST") = TOT_COST
                    rsTemp.Fields("TCP") = TCP
                    rsTemp.Update
                DataGrid1.Refresh
            
                NSTOCK = Val(adostock.Recordset.Fields(3).Value) - Val(txtqty)
                adostock.Recordset.Fields(3) = NSTOCK
                adostock.Recordset.Save
                adostock.Refresh
                adostock.Refresh
                DataList1.Refresh
                DataList1.ReFill
                cmdadd_remove.Enabled = False
                timer_process.Enabled = True
                txtqty = 0
        End If
       
  ElseIf cmdadd_remove.Caption = "REMOVE FROM THE CART" Then
   
        ask = MsgBox("Do you want to remove the current product from the list ?", vbQuestion + vbYesNo, "System query")
        If ask = vbYes Then
            adostock.Recordset.MoveFirst
            Do While adostock.Recordset.Fields(0).Value <> rsTemp.Fields(0).Value
                adostock.Recordset.MoveNext
            Loop
            adostock.Recordset.Fields(3).Value = adostock.Recordset.Fields(3).Value + rsTemp.Fields(3).Value
            adostock.Recordset.Save
            rsTemp.Delete
            rsTemp.Update
            adostock.Refresh
            adostock.Refresh
            DataGrid1.Refresh
            cmdadd_remove.Enabled = False
            timer_process.Enabled = True
        End If

  End If
errorhandler:
If Err <> 0 Then
Select Case Err.Number
    Case -2147217842
        MsgBox "Operation is cancelled because you are trying to enter same product twice.If you want to change the quantity try removing the product first then input the desired quantity. ", vbCritical, "Database Error"
        MsgBox "First select product and type-in quantity then try adding. ", vbExclamation, "System Error"
    Case 94
        MsgBox "Incomplete record in the database.", vbCritical, "Database Error"
    Case 3021
         MsgBox "Select the product you want to remove from the list", vbInformation + vbOKOnly, "Data remove error"
    Case Else
     MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"

End Select
End If
End Sub

Private Sub cmdcancel_Click()
'On Error GoTo errorhandler
    'cancelling transaction
    
    Do While rsTemp.RecordCount <> 0
     rsTemp.MoveFirst
     adostock.Recordset.MoveFirst
    Do While adostock.Recordset.Fields(0).Value <> rsTemp.Fields(0).Value
        adostock.Recordset.MoveNext
    Loop
        adostock.Recordset.Fields(3).Value = adostock.Recordset.Fields(3).Value + rsTemp.Fields(3).Value
        adostock.Recordset.Save
        rsTemp.Delete
        rsTemp.Update
        adostock.Refresh
        adostock.Refresh
        DataGrid1.Refresh
        adostock.Refresh
        
    Loop

    Label2(0).Caption = ""
    cmdstart.Caption = "START TRANSACTION"
    cmdstart.BackColor = &H8000000F
    cmdcancel.BackColor = &H8000000F
    rsTemp.Close
    Set rsTemp = Nothing
    

    'deinitialsing objects
    Frame1.Enabled = False
    Frame2.Enabled = False
    optdealer.Enabled = False
    optother.Enabled = False
    optretail.Enabled = False
    DataCombo1.Enabled = False
    txtterms.Enabled = False
    cmdcancel.Enabled = False
    Label8.Caption = 0
    
    'repalceing with 0s
    txtsub_total = 0
    txtothr_chrg = 0
    txtdiscnt = 0
    txtpaid = 0
    txtdue = 0
    txtttotal = 0
    txtinv_no = ""
    txtinv_date = ""
    txtuser = ""
    txtbname = ""
    txttottran = ""
    txtpay = ""
    txtdebt = ""




errorhandler:
If Err <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If

End Sub

Private Sub cmdinvoice_Click()
If dev_bike.rscmdinv.State = adStateOpen Then
    dev_bike.rscmdinv.Close
End If
    dev_bike.cmdinv Trim(dtprm.Text)
    rptinv_spares_off.Show

End Sub

Private Sub cmdsearch_Click()
On Error GoTo errorhandler
If txtsearch = "" Then
    MsgBox "Please include a search term and try again.", vbExclamation, "Borac Sales System"
Else
    DataList1.BackColor = vbGreen
If cmdsearch.Caption = "Clear Search" Then
   adostock.CommandType = adCmdText
   adostock.RecordSource = "SELECT *FROM stock_spares c"
   adostock.Refresh
   cmdsearch.Caption = "SEARCH"
   txtsearch = ""
   DataList1.BackColor = &H80000005
ElseIf cmdsearch.Caption = "SEARCH" Then
  adostock.CommandType = adCmdText
  adostock.RecordSource = "SELECT *FROM stock_spares WHERE (((stock_spares.BID)  Like '" & Me.txtsearch.Text & "%'))"
  adostock.Refresh
    If adostock.Recordset.RecordCount = 0 Then
        adostock.CommandType = adCmdText
        adostock.RecordSource = "SELECT *FROM stock_spares WHERE (((stock_spares.BIKE_DETAILS)  Like '" & Me.txtsearch.Text & "%'))"
        adostock.Refresh
    End If
    If adostock.Recordset.RecordCount = 0 Then
        MsgBox "No records matched.", vbOKOnly + vbInformation, "System Information"
    Else
        MsgBox adostock.Recordset.RecordCount & " " & "records found.", vbOKOnly + vbInformation, "System Information"
    End If
    
    cmdsearch.Caption = "Clear Search"
    
End If

End If
errorhandler:
If Err.Number <> 0 Then
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End If
End Sub

Private Sub cmdstart_Click()
    'creating temp database
    If cmdstart.Caption = "START TRANSACTION" Then
        Set rsTemp = New ADODB.Recordset
        rsTemp.ActiveConnection = Nothing
        rsTemp.CursorLocation = adUseClient
        rsTemp.LockType = adLockBatchOptimistic
    
        'create fields
        rsTemp.Fields.Append "PARTS ID", adVarChar, 50
        rsTemp.Fields.Append "PARTS DESCRIPTION", adVarChar, 50
        rsTemp.Fields.Append "UNIT PRICE", adVarNumeric, 10
        rsTemp.Fields.Append "QTY", adVarNumeric, 10
        rsTemp.Fields.Append "TOTAL COST", adVarNumeric, 20
        rsTemp.Fields.Append "TCP", adVarNumeric, 20
        rsTemp.Open

        Set DataGrid1.DataSource = rsTemp

        'hiding tcp from datagrid
        DataGrid1.Columns(5).Visible = False
        cmdstart.Caption = "COMPLETE TRANSACTION"
        cmdstart.BackColor = vbGreen
        cmdcancel.BackColor = vbRed

         
 'setting up a new transaction ID without adding new
    
    If adoinvoice.Recordset.RecordCount = 0 Then
        txtinv_date.Text = Format(Date, "dd/MM/yyyy")
        txtinv_no = "ST1005001"
    Else: adoinvoice.Recordset.Sort = "STID ASC"
    adoinvoice.Recordset.MoveLast
        varTID = Mid(adoinvoice.Recordset.Fields(0).Value, 3, 9)
        txtinv_date.Text = Format(Date, "dd/MM/yyyy")
        txtinv_no.Text = "ST" + CStr(varTID + 1)
    End If
    
  'refreshing database
    adostock.Refresh
    adostock.Refresh
    

 'initialsing objects
    Frame1.Enabled = True
    Frame2.Enabled = True
    optdealer.Enabled = True
    optother.Enabled = True
    optretail.Enabled = True
    DataCombo1.Enabled = True
    txtterms.Enabled = True
    cmdcancel.Enabled = True
   
    
    ElseIf cmdstart.Caption = "COMPLETE TRANSACTION" Then
    If txtuser.Text = "" Then
        MsgBox "Please input customer", vbExclamation + vbOKOnly, "System Error"
        txtuser.SetFocus
    ElseIf rsTemp.RecordCount = 0 Then
        MsgBox "No product is in the list !", vbExclamation + vbOKOnly, "System Error"
        DataList1.SetFocus
    ElseIf txtsub_total.Text = "" Or txtothr_chrg.Text = "" Or txtpaid.Text = "" Then
        MsgBox "Fill the Transaction Info", vbExclamation + vbOKOnly, "System Error"
    ElseIf Combo1.Text = "" Then
        MsgBox "Please select the payment mode", vbExclamation + vbOKOnly, "System Error"
        Combo1.SetFocus
    Else
        
    'declaring variables
    
    Dim TCP As Double
    Dim DID As String
    Dim PROFIT As Double
    Dim TI As String
    
     
    'filling up the dealer table
                     
    If optdealer.Value = True Then
        adocustomer.Recordset.Fields(4) = adocustomer.Recordset.Fields(4).Value + txtttotal.Text
        adocustomer.Recordset.Fields(5) = adocustomer.Recordset.Fields(5).Value + txtpaid.Text
        adocustomer.Recordset.Fields(6) = adocustomer.Recordset.Fields(6).Value + txtdue.Text
        adocustomer.Recordset.Update
        adocustomer.Refresh
        adocustomer.Refresh
    End If
    
    'setting up invoice details table
    With adoinvoice_details
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
        .CommandType = adCmdText
        .RecordSource = "SELECT *FROM invoice_details"
        adoinvoice_details.Refresh
    End With
    
    'getting DID
    If optdealer.Value = True Then
        DID = adocustomer.Recordset.Fields(0).Value
    ElseIf optother.Value = True Then
        DID = "<other>"
    End If
    
    'adding up TCP
     rsTemp.MoveFirst
     Do While rsTemp.EOF = False
        TCP = rsTemp.Fields(5).Value + TCP
        rsTemp.MoveNext
     Loop
        'MsgBox TCP
        
    'getting PROFIT
    PROFIT = Val(txtttotal) - TCP
    
    
    'setting up a new transaction ID with adding new
    adoinvoice.Refresh
    adoinvoice.Refresh
    adocustomer.Refresh
    adocustomer.Refresh
   
    If adoinvoice.Recordset.RecordCount = 0 Then
        txtinv_date.Text = Format(Date, "dd/MM/yyyy")
        txtinv_no = "ST1005001"
        adoinvoice.Recordset.AddNew
        
    Else: adoinvoice.Recordset.Sort = "STID ASC"
    adoinvoice.Recordset.MoveLast
        varTID = Mid(adoinvoice.Recordset.Fields(0).Value, 3, 9)
        txtinv_date.Text = Format(Date, "dd/MM/yyyy")
        txtinv_no.Text = "ST" + CStr(varTID + 1)
        adoinvoice.Recordset.AddNew
    End If
    TI = txtinv_no.Text
    
    'filling the invoice table
        adoinvoice.Recordset.Fields(0) = txtinv_no.Text
        adoinvoice.Recordset.Fields(1) = txtinv_date.Text
        adoinvoice.Recordset.Fields(2) = DID
        adoinvoice.Recordset.Fields(3) = txtuser.Text
        adoinvoice.Recordset.Fields(4) = txtterms.Text
        adoinvoice.Recordset.Fields(5) = Combo1.Text
        adoinvoice.Recordset.Fields(6) = txtsub_total.Text
        adoinvoice.Recordset.Fields(7) = txtothr_chrg.Text
        adoinvoice.Recordset.Fields(8) = txtttotal.Text
        adoinvoice.Recordset.Fields(9) = txtpaid.Text
        adoinvoice.Recordset.Fields(10) = txtdue.Text
        adoinvoice.Recordset.Fields(11) = txtdiscnt.Text
        adoinvoice.Recordset.Fields(12) = frmmain.stsbr_main.Panels(2).Text
        adoinvoice.Recordset.Fields(13) = PROFIT
        adoinvoice.Recordset.Save
        adoinvoice.Refresh
        adoinvoice.Refresh


     'filling up the invoice details table
     rsTemp.MoveFirst
    For Counter = 1 To rsTemp.RecordCount
        adoinvoice_details.Recordset.AddNew
        adoinvoice_details.Recordset.Fields(0).Value = txtinv_no.Text
        adoinvoice_details.Recordset.Fields(1).Value = rsTemp.Fields(1).Value
        adoinvoice_details.Recordset.Fields(2).Value = rsTemp.Fields(0).Value
        adoinvoice_details.Recordset.Fields(3).Value = rsTemp.Fields(3).Value
        adoinvoice_details.Recordset.Fields(4).Value = rsTemp.Fields(2).Value
        adoinvoice_details.Recordset.Fields(5).Value = rsTemp.Fields(4).Value
        adoinvoice_details.Recordset.Update
        adoinvoice_details.Refresh
        rsTemp.MoveNext
    Next Counter
    
   
    
    cmdstart.BackColor = &H8000000F
    cmdcancel.BackColor = &H8000000F
    
    'deinitialsing objects
    Frame1.Enabled = False
    Frame2.Enabled = False
    optdealer.Enabled = False
    optother.Enabled = False
    optretail.Enabled = False
    DataCombo1.Enabled = False
    txtterms.Enabled = False
    cmdcancel.Enabled = False
    Label2(0).Caption = ""
    Label8.Caption = 0
    
    'repalceing with 0s
    txtsub_total = 0
    txtothr_chrg = 0
    txtdiscnt = 0
    txtpaid = 0
    txtdue = 0
    txtttotal = 0
    txtinv_no = ""
    txtinv_date = ""
    txtuser = ""
    txtbname = ""
    txttottran = ""
    txtpay = ""
    txtdebt = ""
     
    MsgBox "TRANSACTION SUCCESSFULLY COMPLETED !", vbInformation, "BORAC SALES SYSTEM"
     
    rsTemp.Close
    Set rsTemp = Nothing
    cmdstart.Caption = "START TRANSACTION"
    
        If dev_bike.rscmdinv.State = adStateOpen Then
    dev_bike.rscmdinv.Close
End If
    dev_bike.cmdinv Trim(TI)
    rptinv_spares_off.Show
    
    
    
    End If
    
End If
End Sub

Private Sub DataCombo1_Change()
On Error GoTo errorhandler
    txtuser = ""
    txtbname = ""
    txttottran = ""
    txtpay = ""
    txtdebt = ""
    adocustomer.Recordset.Bookmark = DataCombo1.SelectedItem
    If optdealer.Value = True Then
        txtuser.Text = adocustomer.Recordset.Fields(1).Value
        txtbname = adocustomer.Recordset.Fields(14).Value
        txttottran = adocustomer.Recordset.Fields(4).Value
        txtpay = adocustomer.Recordset.Fields(5).Value
        txtdebt = adocustomer.Recordset.Fields(6).Value
    ElseIf optother.Value = True Then
        txtuser.Text = adocustomer.Recordset.Fields(0).Value
        txtbname = "<OTHER>"
        txttottran = "<OTHER>"
        txtpay = "<OTHER>"
        txtdebt = "<OTHER>"
    End If
  
errorhandler:
If Err.Number <> 0 Then
Select Case Err.Number
Case 94
        MsgBox "Incomplete record in the database.", vbCritical, "Database Error"
Case Else
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End Select
End If
End Sub

Private Sub DataGrid1_GotFocus()
    cmdadd_remove.Caption = "REMOVE FROM THE CART"
    cmdadd_remove.Enabled = True
End Sub


Private Sub DataList1_Click()
On Error GoTo errorhandler
    adostock.Recordset.Bookmark = DataList1.SelectedItem
    txtprts_id = adostock.Recordset.Fields(0).Value
    txtperts_des = adostock.Recordset.Fields(2).Value
    txtup = adostock.Recordset.Fields(7).Value
    Label8.Caption = adostock.Recordset.Fields(3).Value
     
errorhandler:
If Err.Number <> 0 Then
Select Case Err.Number
Case 94
        MsgBox "Incomplete record in the database.", vbCritical, "Database Error"
Case Else
    MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"
End Select
End If
End Sub


Private Sub DataList1_GotFocus()
 cmdadd_remove.Caption = "ADD TO CART"
 cmdadd_remove.Enabled = True

End Sub

Public Sub Form_Load()
Image2.Width = 150000
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdstart.Caption = "COMPLETE TRANSACTION" Then
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

Private Sub optdealer_Click()
    If optdealer.Value = True Then
        Set DataCombo1.RowSource = Nothing
        DataCombo1.Refresh
        With adocustomer
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
        .CommandType = adCmdText
        .RecordSource = "SELECT *FROM dealer"
        adocustomer.Refresh
        Set DataCombo1.RowSource = adocustomer
        DataCombo1.ListField = "DEALER_NAME"
        DataCombo1.ReFill
        DataCombo1.Refresh
       
        End With
    End If
End Sub

Private Sub optother_Click()
    If optother.Value = True Then
        Set DataCombo1.RowSource = Nothing
        DataCombo1.Refresh
        With adocustomer
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\dbase_bike\dbase_bike.mdb;Persist Security Info=False"
        .CommandType = adCmdText
        .RecordSource = "SELECT *FROM others"
        adocustomer.Refresh
        Set DataCombo1.RowSource = adocustomer
        DataCombo1.ListField = "OTHER_NAME"
        DataCombo1.ReFill
        DataCombo1.Refresh
        
        End With
    End If
End Sub

Private Sub optretail_Click()
retail_cust = InputBox("Enter name of retail customer", "RETAIL CUSTOMER")
txtuser.Text = retail_cust
End Sub

Private Sub timer_process_Timer()
'calculating subtotal
Dim sub_total As Currency
If rsTemp.RecordCount <> 0 Then
    If rsTemp.BOF = False Then
        rsTemp.MoveFirst
            Do While rsTemp.EOF = False
                sub_total = rsTemp.Fields(4).Value + sub_total
                rsTemp.MoveNext
            Loop
        txtsub_total = sub_total
    End If
ElseIf rsTemp.RecordCount = 0 Then
        sub_total = "0"
        txtsub_total = sub_total
End If

'refreshing database
   

'calculating transaction values
    txtttotal.Text = Val(txtsub_total.Text) + Val(txtothr_chrg.Text) - Val(txtdiscnt.Text)
    txtdue.Text = Val(txtttotal.Text) - Val(txtpaid.Text)
    
'calculating recordcount of temp
    Label2(0).Caption = rsTemp.RecordCount & " " & " ITEM/ITEMS ARE IN THE LIST"
    
timer_process.Enabled = False
End Sub

Private Sub txtdiscnt_KeyUp(KeyCode As Integer, Shift As Integer)
timer_process.Enabled = True
End Sub

Private Sub txtothr_chrg_KeyUp(KeyCode As Integer, Shift As Integer)
timer_process.Enabled = True
End Sub

Private Sub txtpaid_KeyUp(KeyCode As Integer, Shift As Integer)
timer_process.Enabled = True
End Sub

Private Sub txtqty_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errorhandler
    
     If IsNumeric(txtqty) = False Then
        txtqty.Text = 0
    ElseIf txtqty.Text = "-" Then
        txtqty.Text = "0"
    ElseIf txtqty.Text = "." Then
        txtqty.Text = "+"
    Else:
        txtqty.Text = CInt(txtqty)
        Dim a As Long
        Dim b As Long
        a = Val(txtqty.Text)
        b = Val(adostock.Recordset.Fields(3).Value)
            If Val(adostock.Recordset.Fields(3).Value) = 0 Then
                MsgBox "Sorry no stock is left!", vbOKOnly + vbInformation, "BORAC SALES SYSTEM"
                txtqty.Text = 0
            ElseIf a > b Then
                MsgBox "Sorry you only have" & " " & adostock.Recordset.Fields(3).Value & " " & "Spares left", vbOKOnly + vbInformation, "BORAC SALES SYSTEM"
                txtqty.Text = ""
            End If
    End If
errorhandler:
If Err <> 0 Then
Select Case Err.Number
     Case 94
        MsgBox "Incomplete record in the database.", vbCritical, "Database Error"
     Case Else
     MsgBox "Error number: " & " " & Err.Number & vbCrLf & "Error description: " & Err.Description, vbCritical, "ERROR"

End Select
End If
End Sub


