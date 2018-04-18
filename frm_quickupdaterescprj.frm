VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_quickupdaterescprj 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Quick Update/Recource/Project"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form2"
   ScaleHeight     =   9405
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8775
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15478
      _Version        =   393216
      Rows            =   3
      Cols            =   12
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   15255
      Begin VB.CommandButton cmd_clear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   5160
         TabIndex        =   28
         ToolTipText     =   "Clear"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txt_search 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton cmd_search 
         Caption         =   "Search"
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         ToolTipText     =   "Search"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transfer To Excel"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Click to Apply"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Projects By Date"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton opt_all 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton opt_nonspread 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Non Spread"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opt_spread 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spread"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cbo_year 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.ListBox lst_prj 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   330
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12375
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13440
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update method"
         Height          =   1095
         Left            =   12240
         TabIndex        =   3
         Top             =   0
         Width           =   3015
         Begin VB.ComboBox cbo_curr 
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specific  Line"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Individual Line"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Line Items"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Currency"
            Height          =   255
            Left            =   1800
            TabIndex        =   8
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DC7E5A&
         Height          =   615
         Left            =   14640
         Picture         =   "frm_quickupdaterescprj.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Click to Apply"
         Top             =   1200
         Width           =   615
      End
      Begin VB.ListBox lst_resc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   10800
         X2              =   10800
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource"
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
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   6360
         TabIndex        =   14
         Top             =   120
         Width           =   975
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
         Left            =   13440
         TabIndex        =   13
         Top             =   1155
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
         Left            =   12360
         TabIndex        =   12
         Top             =   1155
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   12240
         X2              =   12240
         Y1              =   120
         Y2              =   1800
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
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
         TabIndex        =   17
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
            Picture         =   "frm_quickupdaterescprj.frx":05C5
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":06D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":0B29
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":0F7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":13CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":181F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":7AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":7DD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":80ED
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":8687
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":8C21
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":91BB
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":9755
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":9867
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":9DA9
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":A343
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":A8DD
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B1B7
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B2C9
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B3DB
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B4ED
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B5FF
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B711
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":B823
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":BDBD
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":C357
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":C8F1
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":CE8B
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":CF9D
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":D0AF
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":D649
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":D75B
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":D86D
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":DE07
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":DF19
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":E4B3
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":EA4D
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":EB5F
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":F0F9
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":F693
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":FC2D
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":FD3F
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":102D9
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":103EB
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":104FD
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":1060F
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":10721
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":10833
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":10DCD
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":10EDF
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":10FF1
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":1158B
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":11B25
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":120BF
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":12659
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":12BF3
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":1318D
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_quickupdaterescprj.frx":13727
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_quickupdaterescprj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub cbo_year_Change()
lst_prj.Clear
Dim f As Integer
f = 0
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
lst_prj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close

If Check1.Value = 1 Then
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
End Sub

Private Sub cbo_year_Click()
lst_prj.Clear
Dim h As Integer
h = 0
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
lst_prj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
If Check1.Value = 1 Then
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If

End Sub

Private Sub Check1_Click()
Dim a As Integer
If Check1.Value = 1 Then


a = 0
For a = 0 To lst_prj.ListCount - 1
lst_prj.Selected(a) = True
Next a
lst_prj.Enabled = False
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
a = 0
For a = 0 To lst_prj.ListCount - 1
lst_prj.Selected(a) = False
Next a
lst_prj.Enabled = True

End If
End Sub

Private Sub Check2_Click()
End Sub

Private Sub cmd_clear_Click()
Dim Slsc As Double
Slsc = 0
For Slsc = 0 To lst_resc.ListCount - 1
lst_resc.Selected(Slsc) = False
Next Slsc
End Sub

Private Sub cmd_search_Click()
Dim Sls As Double
Sls = 0
For Sls = 0 To lst_resc.ListCount - 1
If InStr(lst_resc.List(Sls), txt_search.Text) Then
lst_resc.Selected(Sls) = True
End If

Next Sls
End Sub

Private Sub Command1_Click()
On Error Resume Next
If cbo_curr.Text = "" Then
MsgBox "Select Currency"
Exit Sub
End If
If Option1.Value = True Then
Dim i As Integer
i = 0
For i = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0
   If CDbl(flex_grid.TextMatrix(i, 7)) = Text2.Text Then
   flex_grid.TextMatrix(i, 7) = Format(Text1.Text, "###,##0.00")
   flex_grid.TextMatrix(i, 4) = cbo_curr.Text
    Dim cr3 As New ADODB.Recordset
If cr3.State Then cr3.Close
cr3.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr3.EOF Then
    flex_grid.TextMatrix(i, 5) = cr3!cur_xchgrate
    End If
cr3.Close
   End If
            
Next
ElseIf Option3.Value = True Then
Dim j As Integer
j = 0
For j = 1 To flex_grid.Rows - 1
 
   
   flex_grid.TextMatrix(j, 7) = Format(Text1.Text, "###,##0.00")
    flex_grid.TextMatrix(j, 4) = cbo_curr.Text
      Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr2.EOF Then
    flex_grid.TextMatrix(j, 5) = cr2!cur_xchgrate
    End If
cr2.Close
Next
 
 
 
End If
End Sub

Private Sub command2_Click()

End Sub

Private Sub Command3_Click()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
AppActivate "FlexGrid To Excel"
For i = 0 To flex_grid.Rows - 1
  flex_grid.Row = i
    For n = 0 To 11
        flex_grid.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flex_grid.Text
    Next
Next

End Sub

Private Sub flex_grid_Click()

On Error Resume Next
If Text1.Text = "" Then
MsgBox "Enter Unit Rate"
Exit Sub
End If
If cbo_curr.Text = "" Then
MsgBox "Select Currency"
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

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbYellow
flex_grid.TextMatrix(current, 7) = Text1.Text
 flex_grid.TextMatrix(current, 4) = cbo_curr.Text
 Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr1.EOF Then
    flex_grid.TextMatrix(current, 5) = cr1!cur_xchgrate
    End If
cr1.Close
Next
flex_grid.Col = 1
'Set flex_nob.CellPicture = ImageList1.ListImages(11).Picture

'---------------END------------------

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
On Error Resume Next
main.lbltitle.Caption = "UPDATE UNITRATE BY RESOURCE/PROJECT - EIC TRANSACTIONS"
Text1.Enabled = False
Text2.Enabled = False
Me.Top = 5
Me.Left = 5


Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(7).Enabled = False


Dim rs2 As New ADODB.Recordset
If rs2.State Then rs2.Close
rs2.Open "select DISTINCT(bd_resccode) from cost where bd_costtype='E'  order by bd_resccode", Cn, 3, 2
While Not rs2.EOF
Dim ki As New ADODB.Recordset
If ki.State Then ki.Close
ki.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rs2(0) & "' ", Cn, 3, 2
If Not ki.EOF Then
lst_resc.AddItem rs2(0) & "  -  " & ki(0)
Else
lst_resc.AddItem rs2(0)
End If
rs2.MoveNext
Wend
rs2.Close
cbo_year.Text = Year(Date)
Dim i As Integer
i = 0
For i = 2004 To 2050
cbo_year.AddItem i
Next i


flex_grid.Clear
cbo_curr.Clear
 Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from currencymaster order by cur_currency", Cn, 3, 2
While Not cr2.EOF
cbo_curr.AddItem cr2!cur_currency
cr2.MoveNext
Wend
cr2.Close




Call flex_title
'            If flex_grid.TextMatrix(0, 5) = "% WC" Then
'            flex_grid.CellBackColor = vbGreen
'            End If
If Check1.Value = 1 Then
lst_prj.Enabled = False
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
lst_prj.Enabled = True
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If

 Me.Width = 11415
 Me.Height = 9750
 
End Sub
Public Sub flex_title()

On Error Resume Next

    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "From"
        .ColWidth(1) = 1700
        
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "To"
        .ColWidth(2) = 1700
        .ColAlignment(2) = 0
        
        .TextMatrix(0, 3) = "JobCharge"
        .ColWidth(3) = 3500
        .ColAlignment(3) = 0
        
        .TextMatrix(0, 4) = "Currency"
        .ColWidth(4) = 700
        .CellBackColor = vbGreen
        
        .TextMatrix(0, 5) = "Xchg"
        .ColWidth(5) = 600
        
         
        .TextMatrix(0, 6) = "Qty"
        .ColWidth(6) = 500
        
        .TextMatrix(0, 7) = "Unit Price"
        .ColWidth(7) = 1200
        
        .TextMatrix(0, 8) = "Spread"
        .ColWidth(8) = 1200
        .ColAlignment(8) = 0
        
         .TextMatrix(0, 9) = "Type"
        .ColWidth(9) = 600
        .ColAlignment(9) = 0
        
         .TextMatrix(0, 10) = "CostCode"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        
         .TextMatrix(0, 11) = "Notes"
        .ColWidth(11) = 2400
        .ColAlignment(11) = 0
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
Private Sub lst_prj_Click()
Call flex_title
If Check1.Value = 1 Then
lst_prj.Enabled = False
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
lst_prj.Enabled = True
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
End Sub

Private Sub lst_resc_Click()
 Call flex_title
If Check1.Value = 1 Then
lst_prj.Enabled = False

frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
lst_prj.Enabled = True

frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
End Sub

Private Sub opt_all_Click()
Call flex_title
If Check1.Value = 1 Then
lst_prj.Enabled = False


frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusyElse
lst_prj.Enabled = True


frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
End Sub
Private Sub opt_nonspread_Click()
Call flex_title
If Check1.Value = 1 Then
lst_prj.Enabled = False
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
lst_prj.Enabled = True
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
End Sub
Private Sub opt_spread_Click()
Call flex_title
If Check1.Value = 1 Then
lst_prj.Enabled = False
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataall
Unload frmBusy
Else
lst_prj.Enabled = True
frmBusy.Show
SetParent frmBusy.HWnd, frm_quickupdaterescprj.HWnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_data
Unload frmBusy
End If
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
On Error Resume Next
If Button.Caption = "Modify" Then

Toolbar1.Buttons(3).Enabled = False

For i = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0
            If flex_grid.TextMatrix(i, 0) = "" Then Exit Sub
            id1 = flex_grid.TextMatrix(i, 0)
                                Dim md As New ADODB.Recordset
                                If md.State Then md.Close
                                md.Open "select * from cost where bd_id=" & id1, Cn, 3, 2
                                If Not md.EOF Then
                                    md!bd_curr = flex_grid.TextMatrix(i, 4)
                                    md!bd_unitrate = flex_grid.TextMatrix(i, 7)
                                    md!bd_xchg = flex_grid.TextMatrix(i, 5)
                                md.Update
                                md.Close
                                End If
             
            
Next
MsgBox "UnitRate Updated Successfully"
If Check1.Value = 1 Then
lst_prj.Enabled = False
Call flex_dataall
Else
lst_prj.Enabled = True
Call flex_data
End If
Call flex_title

'to delete


ElseIf Button.Caption = "Close" Then
Unload Me
End If

End Sub
Public Sub flex_data()
'On Error Resume Next
'Call flex_title
 
With flex_grid
        .Rows = 1
For i = 0 To lst_resc.ListCount - 1
If lst_resc.Selected(i) = True Then
nmm = Split(lst_resc.List(i), "  -  ", Len(lst_resc.List(i)), vbTextCompare)


 Dim j As Integer
 j = 0
 For j = 0 To lst_prj.ListCount - 1
 If lst_prj.Selected(j) = True Then
 nmd = Split(lst_prj.List(j), "  -  ", Len(lst_prj.List(j)), vbTextCompare)
 
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
Dim jc As New ADODB.Recordset
        
        If jc.State Then jc.Close
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
If opt_spread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey='" & nmd(0) & "' and bd_costtype='E' and bd_spread <> 'NA'   order by bd_sdate,bd_edate", Cn, 3, 2

   While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
             
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
    
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        cs.Close
        
        
        fldata.MoveNext
    Wend


ElseIf opt_nonspread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'   and bd_projectkey='" & nmd(0) & "' and bd_costtype='E' and bd_spread ='NA'  order by bd_jobcharge,bd_spread", Cn, 3, 2

    
 While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        
        
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
        
              
        
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        cs.Close
        
        
        fldata.MoveNext
    Wend



ElseIf opt_all.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'   and bd_projectkey='" & nmd(0) & "' and bd_costtype='E'   order by bd_sdate,bd_edate,bd_jobcharge,bd_spread", Cn, 3, 2

    
   While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        
        
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
        
              
        
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        cs.Close
        
        
        fldata.MoveNext
    Wend



Else
End If


End If
Next j

End If
Next i
End With
 
End Sub
Public Sub flex_dataall()
'On Error Resume Next
'Call flex_title
 i = 0
With flex_grid
        .Rows = 1
For i = 0 To lst_resc.ListCount - 1
If lst_resc.Selected(i) = True Then
nmm = Split(lst_resc.List(i), "  -  ", Len(lst_resc.List(i)), vbTextCompare)


 
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
Dim jc As New ADODB.Recordset
        
        If jc.State Then jc.Close
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
If opt_spread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "'  and bd_costtype='E' and bd_spread <> 'NA'   order by bd_sdate,bd_edate", Cn, 3, 2

   While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
             
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
    
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        cs.Close
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        
        fldata.MoveNext
    Wend


ElseIf opt_nonspread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "' and bd_costtype='E' and bd_spread ='NA'  order by bd_jobcharge,bd_spread", Cn, 3, 2

    
 While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        
        
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
        
              
        
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        cs.Close
        
        
        fldata.MoveNext
    Wend



ElseIf opt_all.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "' and bd_costtype='E'   order by bd_sdate,bd_edate", Cn, 3, 2

    
   While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        
        
        .TextMatrix(.Rows - 1, 1) = fldata!bd_sdate
        .TextMatrix(.Rows - 1, 2) = fldata!bd_edate
        
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        jc.Close
        .TextMatrix(.Rows - 1, 4) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 5) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_unitrate, "###,###,##0.00")
        
              
        
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_spread
        End If
        spr.Close
        .TextMatrix(.Rows - 1, 9) = fldata!bd_type
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 10) = fldata!bd_costcode
        End If
        .TextMatrix(.Rows - 1, 11) = fldata!bd_notes
        cs.Close
        
        
        fldata.MoveNext
    Wend



Else
End If




End If
Next i
End With
 
End Sub

