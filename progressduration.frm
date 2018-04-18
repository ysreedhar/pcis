VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form progressduration 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress Duration"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
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
      TabPicture(0)   =   "progressduration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "progressduration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3100
         Left            =   -75000
         TabIndex        =   18
         Top             =   300
         Width           =   6495
         Begin VB.TextBox txt_remarks 
            Appearance      =   0  'Flat
            Height          =   2085
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   19
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3100
         Left            =   0
         TabIndex        =   8
         Top             =   300
         Width           =   6495
         Begin VB.TextBox txt_type 
            Height          =   285
            Left            =   5640
            TabIndex        =   24
            Text            =   "A"
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txt_days 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            TabIndex        =   20
            Top             =   1980
            Width           =   975
         End
         Begin VB.ComboBox cbo_jobcharge 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1260
            Width           =   5295
         End
         Begin VB.ComboBox cbo_spreadcode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   600
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64421889
            CurrentDate     =   38733
         End
         Begin MSComCtl2.DTPicker DTP_startdate 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1965
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy H:mm:ss"
            Format          =   64421891
            CurrentDate     =   37987
         End
         Begin MSComCtl2.DTPicker DTP_enddate 
            Height          =   375
            Left            =   2760
            TabIndex        =   15
            Top             =   1965
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy H:mm:ss"
            Format          =   64421891
            CurrentDate     =   37987
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5640
            TabIndex        =   25
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5160
            TabIndex        =   21
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   17
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2760
            TabIndex        =   16
            Top             =   1800
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JobCharge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   11
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spread Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "   Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74160
         TabIndex        =   23
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spread"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox txt_tdays 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   3195
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_Adays 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3210
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_edays 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   3195
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total DAC"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Days"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ETC days"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "progressduration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_jobcharge_Change()
On Error Resume Next
nj = Split(cbo_spreadcode.Text, "  -  ", Len(cbo_spreadcode.Text), vbTextCompare)
Dim kl As New ADODB.Recordset
If kl.State Then kl.Close
kl.Open "select MAX(prgs_enddate) from progressdurationdetails where prgs_spread_code='" & nj(0) & "' ", Cn, 3, 2
If Not kl.EOF Then
DTP_startdate.Value = kl(0)
End If
If DTP_startdate.Value > DTP_enddate.Value Then
DTP_enddate.Value = DTP_startdate.Value
End If
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub cbo_jobcharge_Click()
On Error Resume Next
nj = Split(cbo_spreadcode.Text, "  -  ", Len(cbo_spreadcode.Text), vbTextCompare)
Dim kl As New ADODB.Recordset
If kl.State Then kl.Close
kl.Open "select MAX(prgs_enddate) from progressdurationdetails where prgs_spread_code='" & nj(0) & "' ", Cn, 3, 2
If Not kl.EOF Then
DTP_startdate.Value = kl(0)
End If
If DTP_startdate.Value > DTP_enddate.Value Then
DTP_enddate.Value = DTP_startdate.Value
End If
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub cbo_jobcharge_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_spreadcode_Change()
On Error Resume Next
nj = Split(cbo_spreadcode.Text, "  -  ", Len(cbo_spreadcode.Text), vbTextCompare)
Dim kl As New ADODB.Recordset
If kl.State Then kl.Close
kl.Open "select MAX(prgs_enddate) from progressdurationdetails where prgs_spread_code='" & nj(0) & "' ", Cn, 3, 2
If Not kl.EOF Then
DTP_startdate.Value = kl(0)
End If
If DTP_startdate.Value > DTP_enddate.Value Then
DTP_enddate.Value = DTP_startdate.Value
End If
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub cbo_spreadcode_Click()
On Error Resume Next
nj = Split(cbo_spreadcode.Text, "  -  ", Len(cbo_spreadcode.Text), vbTextCompare)
Dim kl As New ADODB.Recordset
If kl.State Then kl.Close
kl.Open "select MAX(prgs_enddate) from progressdurationdetails where prgs_spread_code='" & nj(0) & "' ", Cn, 3, 2
If Not kl.EOF Then
DTP_startdate.Value = kl(0)
End If
If DTP_startdate.Value > DTP_enddate.Value Then
DTP_enddate.Value = DTP_startdate.Value
End If
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub cbo_spreadcode_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub DTP_enddate_Change()
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub DTP_enddate_Click()
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub


Private Sub DTP_startdate_Change()
 
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub DTP_startdate_Click()
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub


Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")

cbo_spreadcode.Text = frm_progressdurationdetails.cbo_spr.Text

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
Dim l As Double
l = 0
l = Hour(DTP_startdate.Value)

 DTP_startdate.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
DTP_enddate.Value = Format(Date, "dd-MM-yyyy H:mm:ss")

End Sub


Private Sub txt_days_Change()
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_startdate.Value)
dt1 = CDbl(txt_days.Text)
dt2 = dt + dt1
DTP_enddate.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")
End Sub

 

Private Sub txt_days_GotFocus()
On Error Resume Next
txt_days.Text = DTP_enddate.Value - DTP_startdate.Value
End Sub

Private Sub txt_days_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_startdate.Value)
dt1 = CDbl(txt_days.Text)
dt2 = dt + dt1
DTP_enddate.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")
End Sub
