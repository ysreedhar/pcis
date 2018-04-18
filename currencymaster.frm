VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form currencymaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exchange Rate"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Exchange Rate"
      TabPicture(0)   =   "currencymaster.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "currencymaster.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   -75000
         TabIndex        =   12
         Top             =   300
         Width           =   5415
         Begin VB.TextBox txt_notes 
            Height          =   2655
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   13
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   5415
         Begin VB.TextBox txt_exchangerate 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1035
            Width           =   1815
         End
         Begin VB.TextBox txt_currencydesc 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   2640
            Width           =   4815
         End
         Begin VB.ComboBox txt_currency 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker DTP_currency 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   64290817
            CurrentDate     =   37982
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64290817
            CurrentDate     =   38733
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange Rate"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currecy - Description"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exchange Description"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   2400
            Width           =   1560
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   2040
            TabIndex        =   7
            Top             =   1560
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "currencymaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(c_name),c_desc from currency order by c_name", Cn, 3, 2
While Not rs.EOF
txt_currency.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
End Sub

Private Sub txt_currency_Click()
Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select * from currency where c_name='" & txt_currency.Text & "' ", Cn, 3, 2
If Not cr.EOF Then
txt_currencydesc.Text = cr!c_desc
End If
End Sub
