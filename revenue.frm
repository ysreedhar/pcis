VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form revenue 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Revenue Details"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
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
      TabCaption(0)   =   "Revenue"
      TabPicture(0)   =   "revenue.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "revenue.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   -75000
         TabIndex        =   20
         Top             =   300
         Width           =   8775
         Begin VB.TextBox txt_notes 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   21
            Top             =   120
            Width           =   8055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   0
         TabIndex        =   10
         Top             =   300
         Width           =   8775
         Begin VB.TextBox txt_perc 
            Height          =   315
            Left            =   7320
            TabIndex        =   22
            Text            =   "100"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cbo_revtype 
            Height          =   315
            ItemData        =   "revenue.frx":0038
            Left            =   240
            List            =   "revenue.frx":003A
            TabIndex        =   0
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txt_invoice 
            Height          =   315
            Left            =   4800
            TabIndex        =   7
            Top             =   1200
            Width           =   1935
         End
         Begin VB.ComboBox cbo_curcy 
            Height          =   315
            Left            =   1416
            TabIndex        =   1
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txt_amount 
            Height          =   315
            Left            =   2472
            TabIndex        =   2
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txt_exchange 
            Height          =   315
            Left            =   4368
            TabIndex        =   3
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txt_totalamount 
            Height          =   315
            Left            =   5424
            TabIndex        =   4
            Top             =   480
            Width           =   1815
         End
         Begin VB.ComboBox cbo_jobno 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   1200
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTP_inv 
            Height          =   315
            Left            =   6840
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   67502081
            CurrentDate     =   38040
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   3360
            TabIndex        =   6
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   67502081
            CurrentDate     =   38733
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VO(+) %"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7320
            TabIndex        =   23
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revenue Type"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblinv 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4800
            TabIndex        =   18
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbldate 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6840
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1416
            TabIndex        =   16
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2472
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "XRate"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4368
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5424
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbljobno 
            BackStyle       =   0  'Transparent
            Caption         =   "Job No"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   11
            Top             =   960
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "revenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_curcy_Click()

Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select * from currencymaster where cur_currency='" & cbo_curcy.Text & "' ", Cn, 3, 2
If Not cr.EOF Then
txt_exchange.Text = cr!cur_xchgrate
End If
End Sub

Private Sub cbo_revtype_Change()
If cbo_revtype.Text = "BGT" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
txt_perc.Visible = False
Label2.Visible = False
lbljobno.Visible = True
 
ElseIf cbo_revtype.Text = "VO(+)" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
txt_perc.Visible = True
Label2.Visible = True
lbljobno.Visible = True
 
ElseIf cbo_revtype.Text = "VO(-)" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
 
lbljobno.Visible = True
 txt_perc.Visible = False
Label2.Visible = False
ElseIf cbo_revtype.Text = "BLD" Then
txt_invoice.Visible = True
DTP_inv.Visible = True
cbo_jobno.Visible = True
 
lbljobno.Visible = True
 
lblinv.Visible = True
lbldate.Visible = True
Label2.Visible = False
ElseIf cbo_revtype.Text = "BGT VO" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
txt_perc.Visible = False
Label2.Visible = False
lbljobno.Visible = True

End If
End Sub

Private Sub cbo_revtype_Click()
If cbo_revtype.Text = "BGT" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True

lbljobno.Visible = True
txt_perc.Visible = False
Label2.Visible = False
ElseIf cbo_revtype.Text = "VO(+)" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
 
lbljobno.Visible = True
txt_perc.Visible = True
Label2.Visible = True
ElseIf cbo_revtype.Text = "VO(-)" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
 
lbljobno.Visible = True
txt_perc.Visible = False
Label2.Visible = False
ElseIf cbo_revtype.Text = "BLD" Then
txt_invoice.Visible = True
DTP_inv.Visible = True
cbo_jobno.Visible = True
 
lbljobno.Visible = True

lblinv.Visible = True
lbldate.Visible = True
txt_perc.Visible = False
Label2.Visible = False
ElseIf cbo_revtype.Text = "BGT VO" Then
txt_invoice.Text = "-"
txt_invoice.Visible = False
DTP_inv.Visible = False
lblinv.Visible = False
lbldate.Visible = False
cbo_jobno.Visible = True
txt_perc.Visible = False
Label2.Visible = False
lbljobno.Visible = True

End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
DTP_inv.Value = Format(Date, "dd/MM/yyyy")
cbo_revtype.AddItem "BGT"
cbo_revtype.AddItem "VO(+)"
cbo_revtype.AddItem "VO(-)"
cbo_revtype.AddItem "BLD"
cbo_revtype.AddItem "BGT VO"

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(cur_currency)  from currencymaster order by cur_currency", Cn, 3, 2
While Not rs.EOF
cbo_curcy.AddItem rs(0)
rs.MoveNext
Wend
ghj = Split(frm_revenue.cbo_projcode.Text, "  -  ", Len(frm_revenue.cbo_projcode.Text), vbTextCompare)
Dim jn As New ADODB.Recordset
If jn.State Then jn.Close
jn.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & ghj(0) & "'  order by jobno_code", Cn, 3, 2
While Not jn.EOF
cbo_jobno.AddItem jn(0) & "  -  " & jn(1)
jn.MoveNext
Wend
jn.Close

txt_perc.Visible = False
Label2.Visible = False
End Sub

Private Sub txt_amount_Change()
On Error Resume Next
txt_totalamount.Text = CDbl(txt_amount.Text) * CDbl(txt_exchange.Text)
End Sub

Private Sub txt_amount_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_totalamount.Text = CDbl(txt_amount.Text) * CDbl(txt_exchange.Text)
End Sub

Private Sub txt_exchange_Change()
On Error Resume Next
txt_totalamount.Text = CDbl(txt_amount.Text) * CDbl(txt_exchange.Text)
End Sub

