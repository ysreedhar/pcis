VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form budgetedcost 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Resource"
      TabPicture(0)   =   "costtransaction.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "JobCharge"
      TabPicture(1)   =   "costtransaction.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Transaction"
      TabPicture(2)   =   "costtransaction.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "QTY"
      TabPicture(3)   =   "costtransaction.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Unit Rate"
      TabPicture(4)   =   "costtransaction.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "BCWP"
      TabPicture(5)   =   "costtransaction.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "ECTC"
      TabPicture(6)   =   "costtransaction.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_ECTCrm 
            Height          =   285
            Left            =   2400
            TabIndex        =   25
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txt_ECTCdays 
            Height          =   285
            Left            =   2400
            TabIndex        =   24
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_BCWP 
            Height          =   285
            Left            =   2280
            TabIndex        =   23
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txt_wrkper 
            Height          =   285
            Left            =   2280
            TabIndex        =   22
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_Extdamt 
            Height          =   285
            Left            =   5160
            TabIndex        =   21
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txt_Xrate 
            Height          =   285
            Left            =   2760
            TabIndex        =   20
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txt_amount 
            Height          =   285
            Left            =   480
            TabIndex        =   19
            Top             =   1680
            Width           =   1695
         End
         Begin VB.ComboBox cbo_curr 
            Height          =   315
            Left            =   2760
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cbo_uom 
            Height          =   315
            Left            =   480
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_totdays 
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txt_days 
            Height          =   285
            Left            =   2520
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txt_qty 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_tranxdetails 
            Height          =   285
            Left            =   2280
            TabIndex        =   13
            Top             =   1560
            Width           =   4095
         End
         Begin VB.ComboBox cbo_tranxtype 
            Height          =   315
            Left            =   2280
            TabIndex        =   12
            Top             =   480
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_jobdesc 
            Height          =   285
            Left            =   1920
            TabIndex        =   11
            Top             =   1800
            Width           =   4215
         End
         Begin VB.ComboBox cbo_jobcharge 
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8175
         Begin VB.TextBox txt_rescdesc 
            Height          =   285
            Left            =   2280
            TabIndex        =   9
            Top             =   1560
            Width           =   4815
         End
         Begin VB.ComboBox cbo_resc 
            Height          =   315
            Left            =   2280
            TabIndex        =   8
            Top             =   720
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "budgetedcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_jobcharge_Click()
Dim jd As New ADODB.Recordset
If jd.State Then jd.Close
jd.Open "select * from jobcharge where job_code='" & cbo_jobcharge.Text & "' ", Cn, 3, 2
If Not jd.EOF Then
txt_jobdesc.Text = jd!job_desc
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

Private Sub Form_Load()
Call connect

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from resourcemaster", Cn, 3, 2
While Not rs.EOF
cbo_resc.AddItem rs!resc_code
rs.MoveNext
Wend
rs.Close

Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select * from jobcharge", Cn, 3, 2
While Not jc.EOF
cbo_jobcharge.AddItem jc!job_code
jc.MoveNext
Wend

Dim tr As New ADODB.Recordset
If tr.State Then tr.Close
tr.Open "select * from transactionmaster ", Cn, 3, 2
While Not tr.EOF
cbo_tranxtype.AddItem tr!tranx_code
tr.MoveNext
Wend


End Sub
