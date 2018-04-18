VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form userid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Id"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_password 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "******************************"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txt_userid 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTP_tdate 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1755
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64225281
      CurrentDate     =   38733
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction Date"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "userid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
End Sub

