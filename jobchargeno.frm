VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form jobchargeno 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JobNo"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
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
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Job No"
      TabPicture(0)   =   "jobchargeno.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "jobchargeno.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   -75000
         TabIndex        =   10
         Top             =   300
         Width           =   6015
         Begin VB.TextBox txt_notes 
            Height          =   2535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   11
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   0
         TabIndex        =   4
         Top             =   300
         Width           =   6015
         Begin VB.ComboBox cboRevnCalc 
            Height          =   315
            ItemData        =   "jobchargeno.frx":0038
            Left            =   3720
            List            =   "jobchargeno.frx":0042
            TabIndex        =   16
            Text            =   "P"
            Top             =   2520
            Width           =   855
         End
         Begin VB.ComboBox cbo_type 
            Height          =   315
            ItemData        =   "jobchargeno.frx":0050
            Left            =   1920
            List            =   "jobchargeno.frx":005A
            TabIndex        =   14
            Text            =   "MAIN"
            Top             =   2520
            Width           =   1695
         End
         Begin VB.ComboBox cbo_projstatus 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Text            =   "Active"
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txt_jobchargeno 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   1155
            Width           =   3975
         End
         Begin VB.TextBox txt_jobdescno 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   1800
            Width           =   5415
         End
         Begin VB.ComboBox cbo_job 
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   4200
            TabIndex        =   5
            Top             =   1155
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   38733
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RevnCalc"
            Height          =   195
            Left            =   3720
            TabIndex        =   17
            Top             =   2280
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            Height          =   195
            Left            =   1920
            TabIndex        =   15
            Top             =   2280
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   450
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job No"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   510
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   4200
            TabIndex        =   7
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Project Key"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "jobchargeno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_job.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close

cbo_projstatus.AddItem "Active"
cbo_projstatus.AddItem "InActive"
cbo_projstatus.AddItem "WithHeld"
cbo_projstatus.AddItem "Terminated"
' RevnCalc
cboRevnCalc.Clear
cboRevnCalc.AddItem "P"
cboRevnCalc.AddItem "Y"
End Sub
