VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_incurredjob 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         Caption         =   "Close"
         Height          =   255
         Left            =   9000
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   255
         Left            =   9000
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
      Begin VB.CommandButton cmd_show 
         Caption         =   "View"
         Height          =   255
         Left            =   9000
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   3495
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8265
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   10785
      ExtentX         =   19024
      ExtentY         =   14579
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "rpt_incurredjob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
