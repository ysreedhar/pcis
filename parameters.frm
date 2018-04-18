VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form parameters 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txt_notes 
         Appearance      =   0  'Flat
         Height          =   2055
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtp_cdate 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   64946179
         CurrentDate     =   38135
      End
      Begin MSComCtl2.DTPicker dtp_edate 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2010
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   64946179
         CurrentDate     =   38135
      End
      Begin MSComCtl2.DTPicker dtp_sdate 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1260
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   64946179
         CurrentDate     =   38135
      End
      Begin VB.TextBox txt_yeardays 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cutt-Off Date"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year days"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
