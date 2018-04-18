VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form projectmaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Master"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6000
   Begin VB.TextBox txt_notes 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox txt_projectkey 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txt_projectdesc 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1725
      Width           =   5415
   End
   Begin VB.TextBox txt_projecttitle 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1035
      Width           =   3615
   End
   Begin VB.ComboBox cbo_projstatus 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Active"
      Top             =   2400
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTP_tdate 
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      Top             =   1035
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64356353
      CurrentDate     =   38733
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      Height          =   195
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Description"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project  Title"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   885
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   450
   End
End
Attribute VB_Name = "projectmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
 cbo_projstatus.AddItem "Active"
cbo_projstatus.AddItem "InActive"
cbo_projstatus.AddItem "WithHeld"
cbo_projstatus.AddItem "Terminated"
End Sub
