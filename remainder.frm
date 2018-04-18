VERSION 5.00
Begin VB.Form remainder 
   BackColor       =   &H00DC7E5A&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "0"
      Top             =   5520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00DC7E5A&
      Caption         =   "Check  if  YES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00DC7E5A&
      Caption         =   " Remind After"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DC7E5A&
      Caption         =   " Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DC7E5A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "remainder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iit As Double
Dim md As New ADODB.Recordset

Private Sub Check1_Click()

If md.State Then md.Close
md.Open "select * from projectremainder where p_id=" & iit, Cn, 3, 2
If Not md.EOF Then
 
If Check1.Value = 1 Then
Check1.Caption = "YES"
md!proj_remainder = "YES"

md!t_date = DateAdd("d", Text1.Text, md!t_date)
Else
Check1.Caption = "NO"
md!proj_remainder = "NO"
md!t_date = Null
End If
 
md.Update
md.Close
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
iit = 0
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from projectremainder where proj_user='" & main.Label2.Caption & "' and t_date='" & Format(Date, "MM/dd/yyyy ") & "' ", Cn, 3, 2
If Not rs.EOF Then
Label1.Caption = rs!proj_notes
iit = rs!p_id
End If

End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
MsgBox "Enter Numeric Value"
 Text1.SetFocus
 
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'''If IsNumeric(Text1.Text) = False Then
'''MsgBox "Enter Numeric Value"
'''Text1.Text = ""
'''Text1.SetFocus
'''End If
End Sub
