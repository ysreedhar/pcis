VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBusy 
   BackColor       =   &H00DC7E5A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar PBar 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   825
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblBusyString 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait Job Processing................................"
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
      Height          =   240
      Left            =   795
      TabIndex        =   0
      Top             =   510
      Width           =   4890
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'Project Name   :   PCIS
'File Name      :   frmBusy
'Description    :   This Form Used to show the application status like Saving records to
'                   Database, Application Busy with Processing and etc.
'                   Opening Connections
'Functions      :
'Procedures     :
'Date           :   28/04/2005
'Created by     :   Assad Sm
'***************************************************************************************************

'***************************************************************************************************
'Modification Details
'***************************************************************************************************
'Modification   :
'Description    :
'Functions      :
'Procedures     :
'Modified By    :
'Date           :
'***************************************************************************************************
Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000
End Sub
