VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form projecttransaction 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Transaction"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox txt_notes 
         Height          =   315
         ItemData        =   "projecttransaction.frx":0000
         Left            =   2760
         List            =   "projecttransaction.frx":000A
         TabIndex        =   23
         Text            =   "MAIN"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txt_projdesc 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cbo_projkey 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Project To Date"
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2535
         Begin VB.TextBox txt_lye_cost 
            Height          =   300
            Left            =   240
            TabIndex        =   15
            Text            =   "0"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2295
            Begin VB.TextBox txt_lye_revn 
               Height          =   300
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txt_lye_revn1 
               Height          =   300
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Year End - Revn(B)"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Year End - Revn(UB)"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   1845
            End
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Year End- Cost"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year To Date"
         Height          =   2415
         Left            =   2760
         TabIndex        =   1
         Top             =   1560
         Width           =   2415
         Begin VB.TextBox txt_lme_cost 
            Height          =   300
            Left            =   240
            TabIndex        =   7
            Text            =   "0"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            Begin VB.TextBox txt_lme_revn 
               Height          =   300
               Left            =   120
               TabIndex        =   4
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txt_lme_revn1 
               Height          =   300
               Left            =   120
               TabIndex        =   3
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Month End- Revn(B)"
               Height          =   195
               Left            =   120
               TabIndex        =   6
               Top             =   120
               Width           =   1800
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Last Month End- Revn(UB)"
               Height          =   195
               Left            =   120
               TabIndex        =   5
               Top             =   720
               Width           =   1920
            End
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Month End-Cost"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin MSComCtl2.DTPicker DTP_tdate 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   1035
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67305473
         CurrentDate     =   38733
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MAIN/CO"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Project Transaction Description"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transaction Date"
         Height          =   195
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Project Key"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "projecttransaction"
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
cbo_projkey.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close
End Sub

