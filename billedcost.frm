VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form billedcost 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Billed Cost"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5530
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
      TabCaption(0)   =   "ACWP Details"
      TabPicture(0)   =   "billedcost.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "billedcost.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   12015
         Begin VB.ComboBox cbo_tranx 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txt_inv 
            Height          =   285
            Left            =   1320
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   10215
            Begin VB.TextBox txt_Extdamt 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   8400
               TabIndex        =   14
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txt_Xrate 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6960
               TabIndex        =   13
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txt_unitrate 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5160
               TabIndex        =   12
               Top             =   480
               Width           =   1695
            End
            Begin VB.ComboBox cbo_curr 
               Height          =   315
               Left            =   3360
               TabIndex        =   11
               Top             =   480
               Width           =   1695
            End
            Begin VB.ComboBox cbo_uom 
               Height          =   315
               Left            =   1920
               TabIndex        =   10
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txt_totdays 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   255
               Left            =   1920
               TabIndex        =   20
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               Height          =   255
               Left            =   3360
               TabIndex        =   19
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Rate"
               Height          =   255
               Left            =   5160
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Exchange Rate"
               Height          =   255
               Left            =   6960
               TabIndex        =   17
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Extd Amount"
               Height          =   255
               Left            =   8400
               TabIndex        =   16
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Quantity"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.ComboBox cbo_costcode 
            Height          =   315
            Left            =   7200
            TabIndex        =   7
            Top             =   1320
            Width           =   3135
         End
         Begin VB.ComboBox cbo_jobcharge 
            Height          =   315
            Left            =   3720
            TabIndex        =   6
            Top             =   1320
            Width           =   3375
         End
         Begin VB.ComboBox cbo_resc 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   3495
         End
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   4800
            TabIndex        =   4
            Top             =   600
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker DTP_inv 
            Height          =   315
            Left            =   3360
            TabIndex        =   21
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64094209
            CurrentDate     =   38733
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   8880
            TabIndex        =   24
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64094209
            CurrentDate     =   38733
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "TranX Type"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No."
            Height          =   195
            Left            =   1320
            TabIndex        =   31
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
            Height          =   195
            Left            =   3360
            TabIndex        =   30
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Jobcharge Code"
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Costcode"
            Height          =   255
            Left            =   7200
            TabIndex        =   28
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   8880
            TabIndex        =   27
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Resource"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor"
            Height          =   255
            Left            =   4800
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2850
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   10695
         Begin VB.TextBox txt_notes 
            Height          =   2175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   360
            Width           =   10335
         End
      End
   End
End
Attribute VB_Name = "billedcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_curr_Change()
Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
If Not cr1.EOF Then
txt_Xrate.Text = cr1!cur_xchgrate
End If
cr1.Close
End Sub

Private Sub cbo_curr_Click()
Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
If Not cr1.EOF Then
txt_Xrate.Text = cr1!cur_xchgrate
End If
cr1.Close
End Sub

 

 

Private Sub cbo_resc_Click()
nm = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
Dim cc As New ADODB.Recordset
If cc.State Then cc.Close
cc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & nm(0) & "' ", Cn, 3, 2
While Not cc.EOF
cbo_costcode.AddItem cc(0) & "  -  " & cc(1)
cc.MoveNext
Wend
cc.Close
Dim crd As New ADODB.Recordset
If crd.State Then crd.Close
crd.Open "select * from resourcedetails where dresc_code='" & nm(0) & "'   ", Cn, 3, 2
If Not crd.EOF Then
cbo_curr.Text = crd!dresc_curcy
End If
crd.Close
End Sub

Private Sub Form_Load()
On Error Resume Next

DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
DTP_inv.Value = Format(Date, "dd/MM/yyyy")
 
Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select * from currencymaster ", Cn, 3, 2
While Not cr.EOF
cbo_curr.AddItem cr!cur_currency

cr.MoveNext
Wend
cr.Close


Dim pk1 As New ADODB.Recordset
If pk1.State Then pk1.Close
pk1.Open "select * from UOM", Cn, 3, 2
While Not pk1.EOF
cbo_uom.AddItem pk1("uom_uom")
pk1.MoveNext
Wend
pk1.Close

gg = Split(frm_billedcost.cbo_pproj.Text, "  -  ", Len(frm_billedcost.cbo_pproj.Text), vbTextCompare)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(rd.dresc_code),rm.resc_desc from resourcedetails rd ,resourcemaster rm  where  rm.resc_id=rd.resc_id   and dresc_proj='" & gg(0) & "' order by rd.dresc_code", Cn, 3, 2
While Not rs.EOF
billedcost.cbo_resc.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close

Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & gg(0) & "' order by job_code ", Cn, 3, 2
While Not jc.EOF
cbo_jobcharge.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(vendor_code),vendor_desc from vendormaster order by vendor_code", Cn, 3, 2
While Not vn.EOF
cbo_vendor.AddItem vn(0) & "  -  " & vn(1)
vn.MoveNext
Wend
vn.Close
 
cbo_tranx.AddItem "ME"
cbo_tranx.AddItem "AJ"
End Sub

Private Sub txt_totdays_Change()
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)
End Sub

Private Sub txt_totdays_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)
End Sub

Private Sub txt_unitrate_Change()
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)

End Sub

Private Sub txt_unitrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)

End Sub

Private Sub txt_Xrate_Change()
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)

End Sub

Private Sub txt_Xrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = CDbl(txt_unitrate.Text) * CDbl(txt_totdays.Text) * CDbl(txt_Xrate.Text)

End Sub
