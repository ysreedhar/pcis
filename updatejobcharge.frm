VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form updatejobcharge 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Update JobCharge By EIC"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.CommandButton Command3 
         BackColor       =   &H00DC7E5A&
         Caption         =   "Continue To Replace Resource........."
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox TXT_SPREADN 
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   1680
         Width           =   3975
      End
      Begin VB.ComboBox TXT_SPREADO 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   4215
      End
      Begin VB.ComboBox cbo_projnew 
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cbo_newjob 
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   3975
      End
      Begin VB.ComboBox cbo_job 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8895
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update method"
         Height          =   1095
         Left            =   8640
         TabIndex        =   1
         Top             =   0
         Width           =   3015
         Begin VB.CommandButton Command1 
            BackColor       =   &H00DC7E5A&
            Height          =   615
            Left            =   1920
            Picture         =   "updatejobcharge.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Click to Apply"
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specific % Line"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Individual Line"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Line Items"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Spread"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label DD 
         BackStyle       =   0  'Transparent
         Caption         =   "New Spread"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "New Project"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Jobcharge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4560
         TabIndex        =   17
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobcharge to be Replaced"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Project"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   9960
         TabIndex        =   10
         Top             =   1155
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Existing Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   8880
         TabIndex        =   9
         Top             =   1155
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   8640
         X2              =   8640
         Y1              =   120
         Y2              =   1680
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "ar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "hlp"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6885
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   14
         Top             =   0
         Width           =   4215
      End
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   58
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":05C5
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":06D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":0B29
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":0F7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":13CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":181F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":7AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":7DD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":80ED
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":8687
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":8C21
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":91BB
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":9755
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":9867
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":9DA9
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":A343
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":A8DD
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B1B7
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B2C9
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B3DB
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B4ED
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B5FF
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B711
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":B823
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":BDBD
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":C357
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":C8F1
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":CE8B
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":CF9D
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":D0AF
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":D649
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":D75B
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":D86D
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":DE07
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":DF19
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":E4B3
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":EA4D
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":EB5F
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":F0F9
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":F693
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":FC2D
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":FD3F
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":102D9
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":103EB
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":104FD
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":1060F
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":10721
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":10833
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":10DCD
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":10EDF
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":10FF1
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":1158B
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":11B25
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":120BF
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":12659
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":12BF3
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":1318D
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "updatejobcharge.frx":13727
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   7095
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      TextStyle       =   3
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "updatejobcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_job_Click()
 

TXT_SPREADO.Clear
TXT_SPREADO.AddItem "NA  -  Not Applicable"
TXT_SPREADO.AddItem "NA  -  Progress"
nnw2 = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(s.spread_code),s.spread_desc from spreadmaster s , cost c where s.spread_code=c.bd_spread and c.bd_jobcharge='" & nnw2(0) & "' and s.spread_code <>'NA' order by s.spread_code", Cn, 3, 2
While Not spr.EOF
 
TXT_SPREADO.AddItem spr(0) & "  -  " & spr(1)
 
spr.MoveNext
Wend
flex_grid.Clear
Call flex_title
Call flex_data
End Sub
Private Sub cbo_job_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_newjob_Click()
TXT_SPREADN.Clear
 TXT_SPREADN.AddItem "NA  -  Not Applicable"
 TXT_SPREADN.AddItem "NA  -  Progress"
nnw2 = Split(cbo_newjob.Text, "  -  ", Len(cbo_newjob.Text), vbTextCompare)
Dim spr1 As New ADODB.Recordset
If spr1.State Then spr1.Close
spr1.Open "select DISTINCT(s.spread_code),s.spread_desc from spreadmaster s , cost c where s.spread_code=c.bd_spread and  s.spread_code <>'NA' order by s.spread_code", Cn, 3, 2
While Not spr1.EOF
TXT_SPREADN.AddItem spr1(0) & "  -  " & spr1(1)
 
 
spr1.MoveNext
Wend
End Sub

Private Sub cbo_proj_Click()
cbo_job.Clear
 
nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 nmm = Split(cbo_projnew.Text, "  -  ", Len(cbo_projnew.Text), vbTextCompare)
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & nm(0) & "' order by job_code", Cn, 3, 2
While Not rs1.EOF
cbo_job.AddItem rs1(0) & "  -  " & rs1(1)
 
rs1.MoveNext
Wend
rs1.Close
 
 
End Sub

Private Sub cbo_proj_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_projnew_Click()
cbo_newjob.Clear
 
 nmm = Split(cbo_projnew.Text, "  -  ", Len(cbo_projnew.Text), vbTextCompare)
 

Dim rs22 As New ADODB.Recordset
If rs22.State Then rs22.Close
rs22.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & nmm(0) & "' order by job_code", Cn, 3, 2
While Not rs22.EOF
cbo_newjob.AddItem rs22(0) & "  -  " & rs22(1)
rs22.MoveNext
Wend
 
 
End Sub

Private Sub Command1_Click()
On Error Resume Next
 
 


If Option3.Value = True Then
Dim r As Integer
r = 0
For r = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0

                                ji = Split(cbo_newjob.Text, "  -  ", Len(cbo_newjob.Text), vbTextCompare)
                                spd = Split(TXT_SPREADN.Text, "  -  ", Len(TXT_SPREADN.Text), vbTextCompare)
                                spd2 = Split(TXT_SPREADO.Text, "  -  ", Len(TXT_SPREADO.Text), vbTextCompare)
 spd1 = Split(flex_grid.TextMatrix(r, 2), "  -  ", Len(flex_grid.TextMatrix(r, 2)), vbTextCompare)
 If spd1(0) <> "NA" Then
 
 If spd1(0) = spd2(0) Then
                                Dim ty As New ADODB.Recordset
                                If ty.State Then ty.Close
                                ty.Open "select * from progressdurationdetails where prgs_type='" & flex_grid.TextMatrix(r, 1) & "' and  prgs_spread_code='" & spd(0) & "' and prgs_job_key='" & ji(0) & "' ", Cn, 3, 2
                                If ty.EOF Then
                                MsgBox "New JobCharge is not defined in Progress Duration - Pls. setup new jobcharge and perform replacement again"
                                Exit Sub
                                End If

'Current  row
 current = r
'                            For i = 1 To flex_grid.Cols - 1
'                            flex_grid.Col = i
                            flex_grid.CellBackColor = vbYellow
                            flex_grid.TextMatrix(current, 2) = TXT_SPREADN.Text
                            flex_grid.TextMatrix(current, 3) = cbo_newjob.Text
                            flex_grid.TextMatrix(current, 6) = ty!prgs_startdate
                            flex_grid.TextMatrix(current, 7) = ty!prgs_enddate
                            flex_grid.TextMatrix(current, 10) = cbo_projnew.Text
'                       Next i
 End If
    
 Else
 
 If spd1(0) = spd2(0) Then
 
 current = r
'                            For i = 1 To flex_grid.Cols - 1
'                            flex_grid.Col = i
                            flex_grid.CellBackColor = vbYellow
                           flex_grid.TextMatrix(current, 2) = TXT_SPREADN.Text
                          flex_grid.TextMatrix(current, 3) = cbo_newjob.Text
                       flex_grid.TextMatrix(current, 10) = cbo_projnew.Text
'
End If
  
  End If
            
 Next r
 


flex_grid.Col = 1
End If

End Sub

Private Sub Command3_Click()
frm_replaceresc.Show
End Sub

Private Sub flex_grid_Click()

On Error Resume Next
If cbo_newjob.Text = "" Then
MsgBox "Select JobCharge"
Exit Sub
End If
 
'back color

Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = False



Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If
ji = Split(cbo_newjob.Text, "  -  ", Len(cbo_newjob.Text), vbTextCompare)
spd = Split(TXT_SPREADN.Text, "  -  ", Len(TXT_SPREADN.Text), vbTextCompare)
spd2 = Split(TXT_SPREADO.Text, "  -  ", Len(TXT_SPREADO.Text), vbTextCompare)
spd1 = Split(flex_grid.TextMatrix(current, 2), "  -  ", Len(flex_grid.TextMatrix(current, 2)), vbTextCompare)
If spd1(0) <> "NA" Then
If spd1(0) = spd2(0) Then ''''''''''''
Dim ty As New ADODB.Recordset
If ty.State Then ty.Close
ty.Open "select * from progressdurationdetails where prgs_type='" & flex_grid.TextMatrix(current, 1) & "' and  prgs_spread_code='" & spd(0) & "' and prgs_job_key='" & ji(0) & "' ", Cn, 3, 2
If ty.EOF Then
MsgBox "New JobCharge is not defined in Progress Duration - Pls. setup new jobcharge and perform replacement again"
Exit Sub
End If

'Current  row
flex_grid.Row = current
            For i = 1 To flex_grid.Cols - 1
            flex_grid.Col = i
            flex_grid.CellBackColor = vbYellow
            flex_grid.TextMatrix(current, 2) = TXT_SPREADN.Text
            flex_grid.TextMatrix(current, 3) = cbo_newjob.Text
            flex_grid.TextMatrix(current, 6) = ty!prgs_startdate
            flex_grid.TextMatrix(current, 7) = ty!prgs_enddate
            flex_grid.TextMatrix(current, 10) = cbo_projnew.Text
            Next i
  End If ''''''''''''''''
 Else
 If spd1(0) = spd2(0) Then ''''''
flex_grid.Row = current
            For i = 1 To flex_grid.Cols - 1
            flex_grid.Col = i
            flex_grid.CellBackColor = vbYellow
            flex_grid.TextMatrix(current, 2) = TXT_SPREADN.Text
            flex_grid.TextMatrix(current, 3) = cbo_newjob.Text
            flex_grid.TextMatrix(current, 10) = cbo_projnew.Text

            Next
flex_grid.Col = 1
End If ''''''''''
End If
'Set flex_nob.CellPicture = ImageList1.ListImages(11).Picture

'---------------END------------------

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "UPDATE JOBCHARGE - EIC TRANSACTIONS"
Text1.Enabled = False
Text2.Enabled = False
Me.Top = 5
Me.Left = 5


Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(7).Enabled = False
cbo_proj.Clear
cbo_projnew.Clear
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_proj.AddItem rs(0) & "  -  " & rs(1)
cbo_projnew.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close

 


flex_grid.Clear
 
Call flex_title
 
'Call flex_data
 
End Sub
Public Sub flex_title()

On Error Resume Next

    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Type"
        .ColWidth(1) = 500
        
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Spread"
        .ColWidth(2) = 1500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "JobCharge"
        .ColWidth(3) = 3500
        
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "CostCode"
        .ColWidth(4) = 2400
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Rescource"
        .ColWidth(5) = 2000
        .TextMatrix(0, 6) = "Start Date"
        .ColWidth(6) = 2000
        .TextMatrix(0, 7) = "End Date"
        .ColWidth(7) = 2000
        
        .TextMatrix(0, 8) = "ACWP"
        .ColWidth(8) = 2000
        .TextMatrix(0, 9) = "ECTC"
        .ColWidth(9) = 2000
        .TextMatrix(0, 10) = "Project"
        .ColWidth(10) = 2000
 
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

 

Private Sub Option1_Click()
If Option1.Value = True Then
Text1.Enabled = True
Text2.Enabled = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text2.Enabled = False
Text1.Enabled = True
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text1.Enabled = True
Text2.Enabled = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
If Button.Caption = "Modify" Then

Toolbar1.Buttons(3).Enabled = False

For i = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0
hi = Split(flex_grid.TextMatrix(i, 3), "  -  ", Len(flex_grid.TextMatrix(i, 3)), vbTextCompare)
nmm1 = Split(flex_grid.TextMatrix(i, 10), "  -  ", Len(flex_grid.TextMatrix(i, 10)), vbTextCompare)
 
 nmm2 = Split(flex_grid.TextMatrix(i, 2), "  -  ", Len(flex_grid.TextMatrix(i, 2)), vbTextCompare)
            If flex_grid.TextMatrix(i, 0) = "" Then Exit Sub
            id1 = flex_grid.TextMatrix(i, 0)
                                Dim md As New ADODB.Recordset
                                If md.State Then md.Close
                                md.Open "select * from cost where  bd_id=" & id1, Cn, 3, 2
                                If Not md.EOF Then
                                    Dim rscc As New ADODB.Recordset
                                    If rscc.State Then rscc.Close
                                    rscc.Open "select * from resourcedetails where dresc_proj='" & nmm1(0) & "' and dresc_code='" & md!bd_resccode & "' ", Cn, 3, 2
                                    If Not rscc.EOF Then
                                    md!bd_year = rscc!dresc_year
 Else
nm = Split(cbo_projnew.Text, "  -  ", Len(cbo_projnew.Text), vbTextCompare)
pp = Split(cbo_projnew.Text, "20", Len(cbo_projnew.Text), vbTextCompare)
Dim hgg As String
hgg = Mid(pp(1), 1, 2)
 
                Dim sv As New ADODB.Recordset
                If sv.State Then sv.Close
                sv.Open "select * from resourcedetails where dresc_code='" & md!bd_resccode & "' and dresc_proj='" & nm(0) & "'", Cn, 3, 2
                If sv.EOF Then
                sv.AddNew
                sv!dresc_proj = nm(0)
                sv!dresc_code = md!bd_resccode
                sv!dresc_year = "20" & hgg
                sv!dresc_curcy = "RM"
                sv!dresc_rate = 0
                sv!dresc_ratetype = "BR"
                sv!dresc_notes = "-"
                Dim rd As New ADODB.Recordset
                If rd.State Then rd.Close
                rd.Open "select * from resourcemaster where resc_code='" & md!bd_resccode & "' ", Cn, 3, 2
                If Not rd.EOF Then
                sv!resc_id = rd!resc_id
                End If
                sv!t_date = Format(Date, "dd/MM/yyyy")
                sv!u_date = Now
                sv!t_user = main.Label2.Caption
                sv.Update
                sv.Close
                End If
                md!bd_year = "20" & hgg
                
 
  End If
                                    md!bd_projectkey = nmm1(0)
                                     
                                    md!bd_spread = nmm2(0)
                                    
                                    md!bd_jobcharge = hi(0)
                                    md!bd_sdate = flex_grid.TextMatrix(i, 6)
                                    md!bd_edate = flex_grid.TextMatrix(i, 7)
                                    
                                md.Update
                                md.Close
                                End If
'rescassad:
            
Next
MsgBox "Jobcharge Updated Successfully"
Call flex_data
Call flex_title

'to delete


ElseIf Button.Caption = "Close" Then
Unload Me
End If

End Sub

Public Sub flex_data()
'On Error Resume Next
'Call flex_title
nmt = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
nmd = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
With flex_grid
        .Rows = 1
 

Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost  where   bd_jobcharge='" & nmd(0) & "' and bd_projectkey='" & nmt(0) & "' and bd_costtype='E'   order by bd_spread ,bd_jobcharge,bd_costcode", Cn, 3, 2


    
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        .TextMatrix(.Rows - 1, 1) = fldata!bd_type
         
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread
        End If
        spr.Close
        Dim jc As New ADODB.Recordset
        If jc.State Then jc.Close
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 4) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 4) = fldata!bd_costcode
        End If
        cs.Close
        .TextMatrix(.Rows - 1, 5) = fldata!bd_resccode & "  -  " & fldata!bd_rescname
        
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_sdate, "dd/MM/yyyy H:mm:ss")
        
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_edate, "dd/MM/yyyy H:mm:ss")
        
        .TextMatrix(.Rows - 1, 8) = fldata!bd_extdamt
        
        .TextMatrix(.Rows - 1, 9) = fldata!bd_e_extdamt
                 Dim pm As New ADODB.Recordset
        If pm.State Then pm.Close
        pm.Open "select DISTINCT(proj_desc) from projectmaster where proj_key='" & fldata!bd_projectkey & "' ", Cn, 3, 2
        If Not pm.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_projectkey & "  -  " & pm(0)
        End If
        fldata.MoveNext
    Wend



 
End With
 
End Sub



