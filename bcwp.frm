VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form bcwp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab11 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Spread"
      TabPicture(0)   =   "bcwp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "JobCharge"
      TabPicture(1)   =   "bcwp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame21"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Details"
      TabPicture(2)   =   "bcwp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame41"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Unit Rate"
      TabPicture(3)   =   "bcwp.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame51"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame21 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox cbo_jobchargebcwp 
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txt_jobdescbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   23
            Top             =   1200
            Width           =   4215
         End
         Begin VB.Label Label2 
            Caption         =   "Jobcharge Code"
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Jobcharge Description"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1935
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   5535
         Begin VB.ComboBox cbo_spreadbcwp 
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Spread Code"
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame41 
         Height          =   1695
         Left            =   -74640
         TabIndex        =   12
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txt_totdaysbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_daysbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_qtybcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label41 
            Caption         =   "Quantity"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label51 
            Caption         =   "Days"
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label61 
            Caption         =   "Total Quantity"
            Height          =   255
            Left            =   3600
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame51 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   5655
         Begin VB.TextBox txt_bcwpbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   28
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txt_percompbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   27
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txt_Extdamtbcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txt_Xratebcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txt_unitratebcwp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   4
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox cbo_currbcwp 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cbo_uombcwp 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "% Complete"
            Height          =   255
            Left            =   2040
            TabIndex        =   30
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "BCWP Amount"
            Height          =   255
            Left            =   3480
            TabIndex        =   29
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "UOM"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Currency"
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Unit Rate"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Exchange Rate"
            Height          =   255
            Left            =   4320
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "BDGT Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "bcwp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_curr_Click()
Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
If Not cr1.EOF Then
txt_Xrate.Text = cr1!cur_xchgrate
End If
cr1.Close
End Sub

Private Sub cbo_jobcharge_Click()
Dim jd As New ADODB.Recordset
If jd.State Then jd.Close
jd.Open "select * from jobcharge where job_code='" & cbo_jobcharge.Text & "' ", Cn, 3, 2
If Not jd.EOF Then
txt_jobdesc.Text = jd!job_desc
End If


Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & cbo_jobcharge.Text & "' ", Cn, 3, 2
If Not bd.EOF Then
txt_days.Text = bd!bdgt_days
txt_percomp.Text = bd!bdgt_per_workcomplete
End If
If cbo_spread.Text = "NA" Then
txt_days.Text = ""
End If

End Sub

Private Sub cbo_resc_Click()
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select * from resourcemaster where resc_code='" & cbo_resc.Text & "' ", Cn, 3, 2
If Not rs1.EOF Then
txt_rescdesc.Text = rs1!resc_desc
End If
End Sub

Private Sub cbo_tranxtype_Click()
Dim tr1 As New ADODB.Recordset
If tr1.State Then tr1.Close
tr1.Open "select * from transactionmaster where tranx_code='" & cbo_tranxtype.Text & "' ", Cn, 3, 2
If Not tr1.EOF Then
txt_tranxdetails.Text = tr1!tranx_desc

End If

End Sub

Private Sub cbo_spread_Click()
If cbo_spread.Text = "NA" Then
txt_days.Text = ""
Else
cbo_jobcharge_Click
End If
End Sub

Private Sub Form_Load()
Call connect

txt_unitrate.Text = frm_costtranx.txt_stdrate.Text

Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select * from jobcharge", Cn, 3, 2
While Not jc.EOF
cbo_jobcharge.AddItem jc!job_code
jc.MoveNext
Wend
jc.Close

Dim tr As New ADODB.Recordset
If tr.State Then tr.Close
tr.Open "select * from spreadmaster ", Cn, 3, 2
While Not tr.EOF
cbo_spread.AddItem tr!spread_code
tr.MoveNext
Wend
tr.Close

Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select * from currencymaster ", Cn, 3, 2
While Not cr.EOF
cbo_curr.AddItem cr!cur_currency
cr.MoveNext
Wend
cr.Close

Dim um As New ADODB.Recordset
If um.State Then um.Close
um.Open "select * from resourcemaster where resc_code='" & frm_costtranx.textresccode.Text & "' ", Cn, 3, 2
If Not um.EOF Then
cbo_uom.Text = um!resc_uom
End If

Dim pk1 As New ADODB.Recordset
If pk1.State Then pk1.Close
pk1.Open "select * from UOM", Cn, 3, 2
While Not pk1.EOF
cbo_uom.AddItem pk1("uom_uom")
pk1.MoveNext
Wend
pk1.Close
End Sub

Private Sub txt_days_Change()
On Error Resume Next
txt_totdays.Text = txt_qty * txt_days.Text
End Sub

Private Sub txt_Extdamt_Change()
txt_bcwp.Text = CDbl(txt_Extdamt.Text) * CDbl(txt_percomp.Text / 100)
End Sub

Private Sub txt_qty_Change()
On Error Resume Next
txt_totdays.Text = txt_qty * txt_days.Text
End Sub

Private Sub txt_qty_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_totdays.Text = txt_qty * txt_days.Text
End Sub

Private Sub txt_Xrate_Change()
On Error Resume Next
txt_Extdamt.Text = (txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
End Sub

