VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tranxcost 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8070
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
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Project Transaction"
      TabPicture(0)   =   "tranxcost.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "tranxcost.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   5535
         Begin VB.ComboBox cbo_job 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   4335
         End
         Begin VB.ComboBox cbo_projkey 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   4335
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Project To Date"
            Height          =   1215
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   2535
            Begin VB.TextBox txt_lye_cost 
               Height          =   300
               Left            =   120
               TabIndex        =   8
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Year End- Cost"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Year To Date"
            Height          =   1215
            Left            =   2760
            TabIndex        =   4
            Top             =   1800
            Width           =   2175
            Begin VB.TextBox txt_lme_cost 
               Height          =   300
               Left            =   120
               TabIndex        =   5
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Month End-Cost"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   1695
            End
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   3435
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   28246017
            CurrentDate     =   38733
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Job"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   3240
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Project Key"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   5535
         Begin VB.TextBox txt_notes 
            Height          =   2535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   240
            Width           =   4815
         End
      End
   End
End
Attribute VB_Name = "tranxcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub cbo_projkey_Click()
bh = Split(cbo_projkey.Text, "  -  ", Len(cbo_projkey.Text), vbTextCompare)
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & bh(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
cbo_job.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
End Sub

Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_projkey.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close

End Sub
