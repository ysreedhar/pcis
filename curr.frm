VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form curr 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Currency"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
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
      TabCaption(0)   =   "Currency "
      TabPicture(0)   =   "curr.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "curr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   -75000
         TabIndex        =   12
         Top             =   300
         Width           =   5535
         Begin VB.TextBox txt_notes 
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   13
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   5535
         Begin VB.TextBox txt_currencydesc 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   4815
         End
         Begin VB.TextBox txt_currencyname 
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   435
            Width           =   1695
         End
         Begin VB.TextBox txt_minor 
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txt_major 
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   1635
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   2160
            TabIndex        =   6
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64749569
            CurrentDate     =   38733
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   2160
            TabIndex        =   11
            Top             =   2040
            Width           =   1230
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   1470
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Minor Unit"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Major Unit"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "curr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
End Sub
