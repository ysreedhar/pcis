VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form baseline 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BusinessPlan Budget"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5318
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
      TabCaption(0)   =   "BP Budget"
      TabPicture(0)   =   "baseline.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "baseline.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   6015
         Begin VB.TextBox txt_cost 
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ComboBox cbo_job 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   3615
         End
         Begin VB.ComboBox cbo_project 
            Height          =   315
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txt_revn 
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   1920
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   4200
            TabIndex        =   5
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64094209
            CurrentDate     =   38733
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost(RM)"
            Height          =   195
            Left            =   2280
            TabIndex        =   13
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JobNo."
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   4200
            TabIndex        =   8
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revn(RM)"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Key"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   120
            Width           =   810
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   6015
         Begin VB.TextBox txt_notes 
            Height          =   2175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   240
            Width           =   5535
         End
      End
   End
End
Attribute VB_Name = "baseline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_project_Click()
cbo_job.Clear
sk = Split(cbo_project.Text, "  -  ", Len(cbo_project.Text), vbTextCompare)
Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & sk(0) & "' order by jobno_code", Cn, 3, 2
While Not jc.EOF
cbo_job.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
End Sub

Private Sub Form_Load()
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project  and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
cbo_project.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
End Sub

