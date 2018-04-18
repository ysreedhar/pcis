VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form othertransactionmaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Other TranX"
      TabPicture(0)   =   "othertransactionmaster.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "othertransactionmaster.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   5175
         Begin VB.ComboBox cbo_exp 
            Height          =   315
            ItemData        =   "othertransactionmaster.frx":0038
            Left            =   1200
            List            =   "othertransactionmaster.frx":0042
            TabIndex        =   10
            Text            =   "Expenditure"
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txt_Desc 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   1155
            Width           =   5055
         End
         Begin VB.TextBox txt_tranx 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   3840
            TabIndex        =   6
            Top             =   435
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   67239937
            CurrentDate     =   38276
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Income/Expenditure"
            Height          =   195
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TranX"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   3840
            TabIndex        =   7
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txt_notes 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   240
            Width           =   4575
         End
      End
   End
End
Attribute VB_Name = "othertransactionmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
