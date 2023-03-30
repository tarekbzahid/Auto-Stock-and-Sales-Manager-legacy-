VERSION 5.00
Begin VB.Form frmTESTINGONLY 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Text            =   "500"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "0"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "total"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "other charge"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "discount"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESTINGONLY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Adodc1.Recordset.MoveFirst
        'Do While Adodc1.Recordset.EOF = False
        For c = 0 To Adodc1.Recordset.RecordCount + 1
            Adodc2.Recordset.AddNew
            Adodc2.Recordset.Fields(4).Value = Adodc1.Recordset.Fields(7).Value
            Adodc2.Recordset.Save
            Adodc2.Refresh
            Adodc1.Recordset.MoveNext
            c = c + 1
        Next
            'Adodc1.Recordset.MoveNext
        'Loop








'Adodc1.Recordset.Open
'Adodc2.Recordset.Open

'Adodc1.Recordset.MoveFirst
'Do Until Adodc1.Recordset.EOF
   'Adodc2.Recordset.AddNew
   'For i = 0 To Adodc1.Fields.
     ' Adodc2.Fields(i) = Adodc1.Fields(i)
   'Next i
   'Adodc2.Update
   'Adodc1.MoveNext
'Loop
End Sub

Private Sub Command2_Click()
Dim p As Long
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
p = DataGrid1.Columns(7) + p
Adodc1.Recordset.MoveNext
Loop
Label1.Caption = p
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = Adodc1.Recordset.RecordCount
MsgBox a
End Sub

Private Sub Command4_Click()
    Dim a As Currency
    Dim b As Currency
    Dim c As Currency
    a = 45.67
    b = 56.87
    c = a * b
    Text2.Text = c
    DataGrid1.Columns(0).Value = c
End Sub



Private Sub DataGrid1_AfterUpdate()
Dim p As Integer
Adodc1.Recordset.MoveFirst
 Do While Adodc1.Recordset.EOF = False
 p = DataGrid1.Columns(7) + p
 Adodc1.Recordset.MoveNext
Loop
Label1.Caption = p
End Sub

Private Sub Text1_Change()
Text3 = Text4 - Text1 + Text2
End Sub

Private Sub Text2_Change()
Text3 = Text4 - Text1 + Text2
End Sub
