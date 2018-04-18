VERSION 5.00
Begin VB.Form password 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_password 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txt_userid 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
