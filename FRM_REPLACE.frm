VERSION 5.00
Begin VB.Form FRM_REPLACE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_job 
      BackColor       =   &H00FF8080&
      Caption         =   "Continue To Replace Jobcharge........."
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "Continue To Replace Resource........."
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Replace Spread"
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Replace"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox TXT_SPREADO 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   3855
      End
      Begin VB.ComboBox TXT_SPREADN 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.ComboBox txt_job1 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Spread(New)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label DD 
         BackStyle       =   0  'Transparent
         Caption         =   "Spread(old)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "JobCharge"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox cbo_projcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key - Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FRM_REPLACE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_projcode_Click()
nm = Split(cbo_projcode.Text, "  -  ", Len(cbo_projcode.Text), vbTextCompare)
 Dim rc As New ADODB.Recordset
 
            


            
            txt_job1.Clear
            
            Dim rc1 As New ADODB.Recordset
            If rc1.State Then rc1.Close
            rc1.Open "select DISTINCT(j.job_code),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and j.job_proj_key = '" & nm(0) & "'    order by j.job_code", Cn, 3, 2
            While Not rc1.EOF
            txt_job1.AddItem rc1(0) & "  -  " & rc1(1)
            
            rc1.MoveNext
            Wend
            rc1.Close
End Sub

Private Sub cmd_job_Click()
updatejobcharge.Show
End Sub

Private Sub Command1_Click()
If cbo_projcode.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
If txt_job1.Text = "" Then
MsgBox "Select JobCharge"
Exit Sub
End If
If TXT_SPREADO.Text = "" Then
MsgBox "Select Existing Spread"
Exit Sub
End If

If TXT_SPREADN.Text = "" Then
MsgBox "Select New Spread"
Exit Sub
End If
nn1 = Split(TXT_SPREADN.Text, "  -  ", Len(TXT_SPREADN.Text), vbTextCompare)
nn2 = Split(TXT_SPREADO.Text, "  -  ", Len(TXT_SPREADO.Text), vbTextCompare)
nn3 = Split(txt_job1.Text, "  -  ", Len(txt_job1.Text), vbTextCompare)

Cn.Execute "update cost set bd_spread='" & nn1(0) & "' , bd_type='A'  where bd_spread='" & nn2(0) & "' AND BD_COSTTYPE='E'   and bd_jobcharge='" & nn3(0) & "'"
MsgBox "DONE"
End Sub

Private Sub command2_Click()
nn4 = Split(TXT_RESOURCEN.Text, "  -  ", Len(TXT_RESOURCEN.Text), vbTextCompare)
nn5 = Split(txt_spreado1.Text, "  -  ", Len(txt_spreado1.Text), vbTextCompare)
nn6 = Split(TXT_RESOURCEO.Text, "  -  ", Len(TXT_RESOURCEO.Text), vbTextCompare)
nn7 = Split(txt_job2.Text, "  -  ", Len(txt_job2.Text), vbTextCompare)
Cn.Execute "UPDATE COST SET BD_RESCCODE='" & nn4(0) & "' , bd_type='A' WHERE BD_SPREAD='" & nn5(0) & "' AND BD_RESCCODE='" & nn6(0) & "' AND BD_COSTTYPE='E'   and bd_jobcharge='" & nn7(0) & "'"
MsgBox "DONE"
End Sub

Private Sub Command3_Click()
frm_replaceresc.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 5
Me.Left = 5


Dim rv As New ADODB.Recordset
If rv.State Then rv.Close
rv.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rv.EOF
cbo_projcode.AddItem rv(0) & "  -  " & rv(1)
rv.MoveNext
Wend


End Sub

Private Sub txt_job1_Click()

TXT_SPREADO.Clear
TXT_SPREADN.Clear
nnw2 = Split(txt_job1.Text, "  -  ", Len(txt_job1.Text), vbTextCompare)
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(s.spread_code),s.spread_desc from spreadmaster s , cost c where s.spread_code=c.bd_spread and c.bd_jobcharge='" & nnw2(0) & "' and s.spread_code <>'NA' order by s.spread_code", Cn, 3, 2
While Not spr.EOF
 
TXT_SPREADO.AddItem spr(0) & "  -  " & spr(1)
 
spr.MoveNext
Wend
Dim spr1 As New ADODB.Recordset
If spr1.State Then spr1.Close
spr1.Open "select DISTINCT(s.spread_code),s.spread_desc from spreadmaster s , cost c where s.spread_code=c.bd_spread and  s.spread_code <>'NA' order by s.spread_code", Cn, 3, 2
While Not spr1.EOF
TXT_SPREADN.AddItem spr1(0) & "  -  " & spr1(1)
 
 
spr1.MoveNext
Wend
End Sub

  

 
Private Sub txt_job2_Change()

End Sub

Private Sub TXT_RESOURCEO_Change()

End Sub

