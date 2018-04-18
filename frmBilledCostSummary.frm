VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBilledCostSummary 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Billed Invoice Summary"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   12675
   Begin VB.ComboBox cbo_spread 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtInvoiceNumber 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   870
      Width           =   1815
   End
   Begin VB.TextBox txtSubConName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTP_tdate 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   49283073
      CurrentDate     =   39185
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvoiceDetails 
      Height          =   3975
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   350
      ForeColor       =   12582912
      BackColorFixed  =   12582912
      ForeColorFixed  =   16777215
      BackColorSel    =   8388608
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvoices 
      Height          =   6615
      Left            =   6000
      TabIndex        =   11
      Top             =   240
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   350
      ForeColor       =   12582912
      BackColorFixed  =   12582912
      ForeColorFixed  =   16777215
      BackColorSel    =   8388608
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   7
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spread"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2220
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Con Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   1125
   End
End
Attribute VB_Name = "frmBilledCostSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_spread_Change()
LoadInvoices
End Sub

Private Sub Form_Load()
LoadSpreads
flex_title
End Sub
Private Function LoadSpreads()
    Dim tr As New ADODB.Recordset
    If tr.State Then tr.Close
    tr.Open "select DISTINCT(p.prgs_spread_code),s.spread_desc   from progressdurationdetails p,spreadmaster s where p.prgs_spread_code=s.spread_code order by prgs_spread_code", Cn, 3, 2
    While Not tr.EOF
    cbo_spread.AddItem tr(0) & "  -  " & tr(1)
    tr.MoveNext
    Wend
    tr.Close
End Function
Private Function LoadInvoices()
    Dim rsInvoices As New ADODB.Recordset
    If rsInvoices.State Then rsInvoices.Close
    rsInvoices.Open "select DISTINCT(p.prgs_spread_code),s.spread_desc   from progressdurationdetails p,spreadmaster s where p.prgs_spread_code=s.spread_code order by prgs_spread_code", Cn, 3, 2
    While Not tr.EOF
    cbo_spread.AddItem rsInvoices(0) & "  -  " & rsInvoices(1)
    rsInvoices.MoveNext
    Wend
    rsInvoices.Close
End Function
Public Sub flex_title()
On Error Resume Next
    With flxInvoices
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Sub-Con Name"
        .ColWidth(1) = 3500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Invoice No."
        .ColWidth(2) = 2500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Date."
        .ColWidth(3) = 800
        .TextMatrix(0, 4) = "Amount"
        .ColWidth(4) = 1200
        .ColAlignment(4) = 0
    End With
        With flxInvoiceDetails
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "JobCharge"
        .ColWidth(1) = 500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Cost Code"
        .ColWidth(2) = 500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Amount"
        .ColWidth(3) = 500
    End With
End Sub

