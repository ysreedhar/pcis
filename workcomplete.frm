VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form workcomplete 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "% Work Complete(Spread)"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
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
      TabCaption(0)   =   "Spread"
      TabPicture(0)   =   "workcomplete.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "workcomplete.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   5895
         Begin VB.ComboBox cbo_spreadcode 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   4455
         End
         Begin VB.ComboBox cbo_jobcharge 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1140
            Width           =   5415
         End
         Begin VB.TextBox txt_bdgtdays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txt_per_wrkcmpltd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   195
            Left            =   1800
            TabIndex        =   4
            Top             =   2040
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   3000
            TabIndex        =   8
            Top             =   2010
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   49283073
            CurrentDate     =   38733
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spread Code - Description"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1860
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   3000
            TabIndex        =   12
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JobCharge - Description"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bdgt Days"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% WC"
            ForeColor       =   &H00800080&
            Height          =   195
            Left            =   1800
            TabIndex        =   9
            Top             =   1800
            Width           =   450
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   5895
         Begin VB.TextBox txt_remarks 
            Appearance      =   0  'Flat
            Height          =   2205
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   240
            Width           =   5415
         End
      End
   End
End
Attribute VB_Name = "workcomplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_jobcharge_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_spreadcode_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
cbo_spreadcode.Text = frm_workcomplete.cbo_spr.Text

Dim spr As New ADODB.Recordset
Dim jc As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(spread_code),spread_desc from spreadmaster where spread_code <>'NA' order by spread_code", Cn, 3, 2
While Not spr.EOF
cbo_spreadcode.AddItem spr(0) & "  -  " & spr(1)
spr.MoveNext
Wend
spr.Close
If jc.State Then jc.Close
jc.Open "select DISTINCT(job_code), job_desc from jobcharge order by job_code", Cn, 3, 2
While Not jc.EOF
cbo_jobcharge.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close

End Sub



