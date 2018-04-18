VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form jobcharge 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JobCharge"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
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
      TabCaption(0)   =   "JobCharge"
      TabPicture(0)   =   "jobcharge.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "jobcharge.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   -75000
         TabIndex        =   16
         Top             =   300
         Width           =   6015
         Begin VB.TextBox txt_notes 
            Height          =   3375
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   17
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   6015
         Begin VB.TextBox txt_jobdesc 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   3480
            Width           =   5535
         End
         Begin VB.TextBox txt_jobcharge 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2115
            Width           =   4335
         End
         Begin VB.ComboBox cbo_projkey 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   4335
         End
         Begin VB.ComboBox cbo_projstatus 
            Height          =   315
            ItemData        =   "jobcharge.frx":0038
            Left            =   120
            List            =   "jobcharge.frx":003A
            TabIndex        =   4
            Text            =   "Active"
            Top             =   2760
            Width           =   1695
         End
         Begin VB.ComboBox cbo_jobno 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   4335
         End
         Begin VB.ComboBox cbo_subjobno 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   1560
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   3120
            TabIndex        =   8
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64946177
            CurrentDate     =   38733
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JobCharge Description"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   3240
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Charge"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Key"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job no"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Job no"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   3120
            TabIndex        =   9
            Top             =   2520
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "jobcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_jobno_Click()
On Error Resume Next
kl = Split(cbo_jobno.Text, "  -  ", Len(cbo_jobno.Text), vbTextCompare)
kl1 = Split(cbo_subjobno.Text, "  -  ", Len(cbo_subjobno.Text), vbTextCompare)
txt_jobcharge.Text = kl(0) & "-" & kl1(0)
End Sub

Private Sub cbo_jobno_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cbo_projkey_Change()
cbo_jobno.Clear
sk = Split(cbo_projkey.Text, "  -  ", Len(cbo_projkey.Text), vbTextCompare)
Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & sk(0) & "' order by jobno_code", Cn, 3, 2
While Not jc.EOF
cbo_jobno.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
End Sub

Private Sub cbo_projkey_Click()
cbo_jobno.Clear
sk = Split(cbo_projkey.Text, "  -  ", Len(cbo_projkey.Text), vbTextCompare)
Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & sk(0) & "' order by jobno_code", Cn, 3, 2
While Not jc.EOF
cbo_jobno.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
End Sub

Private Sub cbo_projkey_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cbo_projstatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_subjobno_Click()
On Error Resume Next
kl = Split(cbo_jobno.Text, "  -  ", Len(cbo_jobno.Text), vbTextCompare)
kl1 = Split(cbo_subjobno.Text, "  -  ", Len(cbo_subjobno.Text), vbTextCompare)
txt_jobcharge.Text = kl(0) & "-" & kl1(0)
End Sub

Private Sub cbo_subjobno_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
cbo_projkey.Text = frm_jobcharge.cbo_proj.Text
cbo_projkey.Enabled = False
Dim ld As New ADODB.Recordset
If ld.State Then ld.Close
ld.Open "select DISTINCT(proj_key),proj_desc from projectmaster order by proj_key", Cn, 3, 2
While Not ld.EOF
cbo_projkey.AddItem ld(0) & "  -  " & ld(1)
ld.MoveNext
Wend
ld.Close
cbo_projstatus.AddItem "Active"
cbo_projstatus.AddItem "InActive"
cbo_projstatus.AddItem "WithHeld"
cbo_projstatus.AddItem "Terminated"


Dim sjc As New ADODB.Recordset
If sjc.State Then sjc.Close
sjc.Open "select DISTINCT(subjobno_code),subjobno_desc from subjobno order by subjobno_code", Cn, 3, 2
While Not sjc.EOF
cbo_subjobno.AddItem sjc(0) & "  -  " & sjc(1)
sjc.MoveNext
Wend
sjc.Close
End Sub
