VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form projectremainder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reminder"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txt_notes 
      Height          =   3165
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
   End
   Begin MSComCtl2.DTPicker DTP_date 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   28377089
      CurrentDate     =   37982
   End
   Begin MSComCtl2.DTPicker dtp_remainder 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   28377089
      CurrentDate     =   37982
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remind on  Date"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Notes"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "projectremainder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
dtp_remainder.Visible = True
Label1.Visible = True
Else
dtp_remainder.Visible = False
Label1.Visible = False
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_remainder.Visible = False
Label1.Visible = False
dtp_remainder.Value = Date
DTP_date.Value = Date

If main.Label2.Caption = frm_projectremainder.cbo_users.Text Then
DTP_date.Enabled = True
txt_notes.Enabled = True
Check1.Enabled = True
dtp_remainder.Enabled = True
Else
DTP_date.Enabled = False
txt_notes.Enabled = False
Check1.Enabled = False
dtp_remainder.Enabled = False
End If


End Sub
