VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ImportJobChargeResources 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DUPLICATE FROM EXISTING JOBCHARGE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboSourceJobCharge 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      Top             =   360
      Width           =   7935
   End
   Begin VB.ComboBox cbo_spr 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   " "
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Duplicate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox cboTargetJobCharge 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   7935
   End
   Begin MSFlexGridLib.MSFlexGrid flxTargetResources 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   1
      Cols            =   20
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      TextStyle       =   3
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxSourceResources 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   22
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      TextStyle       =   3
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   360
      Top             =   1320
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
            Picture         =   "frm_ImportJobChargeResources.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   1320
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
            Picture         =   "frm_ImportJobChargeResources.frx":13274
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":13386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":137D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":13C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1407C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":144CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1A768
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1AA82
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1AD9C
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1B336
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1B8D0
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1BE6A
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1C404
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1C516
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1CA58
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1CFF2
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1D58C
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1DE66
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1DF78
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1E08A
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1E19C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1E2AE
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1E3C0
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1E4D2
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1EA6C
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1F006
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1F5A0
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1FB3A
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1FC4C
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":1FD5E
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":202F8
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":2040A
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":2051C
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":20AB6
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":20BC8
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":21162
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":216FC
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":2180E
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":21DA8
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":22342
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":228DC
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":229EE
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":22F88
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":2309A
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":231AC
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":232BE
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":233D0
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":234E2
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":23A7C
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":23B8E
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":23CA0
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":2423A
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":247D4
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":24D6E
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":25308
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":258A2
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":25E3C
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportJobChargeResources.frx":263D6
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTransactionType 
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Spread"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Source JobCharge"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Target JobCharge"
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
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
End
Attribute VB_Name = "frm_ImportJobChargeResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sprcode As String
Dim intRowsInserted As Integer
Dim jchrg As String
Dim ntotal As Double
'''''
Const strChecked = "þ"
Const strUnChecked = "q"

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
Function InsertResource(dblrowID As Double, strSpreadKey As String, strSourceJCKey As String, strTargetJobKey As String, strSourceType As String, strTargetType As String)
Set cmd = New ADODB.Command
If Cn.State Then Cn.Close
Cn.Open
cmd.ActiveConnection = Cn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "spImportResourcesFromJobCharge"
cmd.parameters("@bd_ID").Value = dblrowID
cmd.parameters("@Spread").Value = strSpreadKey
cmd.parameters("@jobchargeSource").Value = strSourceJCKey
cmd.parameters("@jobchargeTarget").Value = strTargetJobKey
cmd.parameters("@Sourcebd_type").Value = strSourceType
cmd.parameters("@Targetbd_type").Value = strTargetType
cmd.parameters("@t_user").Value = main.Label2.Caption
cmd.parameters("@result").Value = 0
cmd.Execute
intRowsInserted = intRowsInserted + cmd("@result")
Set cmd.ActiveConnection = Nothing
End Function
Function InsertResourceBudget(dblrowID As Double, strSpreadKey As String, strSourceJCKey As String, strTargetJobKey As String)
Set cmd = New ADODB.Command
If Cn.State Then Cn.Close
Cn.Open
cmd.ActiveConnection = Cn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "spImportResourcesFromJobChargeforBudget"
cmd.parameters("@bd_ID").Value = dblrowID
cmd.parameters("@Spread").Value = strSpreadKey
cmd.parameters("@jobchargeSource").Value = strSourceJCKey
cmd.parameters("@jobchargeTarget").Value = strTargetJobKey
cmd.parameters("@t_user").Value = main.Label2.Caption
cmd.parameters("@result").Value = 0
cmd.Execute
intRowsInserted = intRowsInserted + cmd("@result")
Set cmd.ActiveConnection = Nothing
End Function
Public Sub LoadFlexSourceJCSources()
On Error Resume Next
strJobKey = Split(cboSourceJobCharge.Text, "  -  ", Len(cboSourceJobCharge.Text), vbTextCompare)
strSpreadKey = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim ntotal As Double
ntotal = 0
Dim iddd As Double
iddd = 0
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
If lblTransactionType.Caption = "E" Then
fldata3.Open "select * from cost  where bd_jobcharge='" & strJobKey(0) & "' and bd_spread='" & strSpreadKey(0) & "' and bd_type = '" & strJobKey(2) & "' and bd_costtype= '" & lblTransactionType.Caption & "' order by bd_tranx,bd_spread,bd_sdate,bd_jobcharge,bd_costcode", Cn, 3, 2
ElseIf lblTransactionType.Caption = "B" Then
fldata3.Open "select * from cost  where bd_jobcharge='" & strJobKey(0) & "' and bd_spread='" & strSpreadKey(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' order by bd_tranx,bd_spread,bd_sdate,bd_jobcharge,bd_costcode", Cn, 3, 2
End If
With flxSourceResources
    .Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = fldata3(0)
        .TextMatrix(.Rows - 1, 2) = fldata3!bd_type
         Dim spd As New ADODB.Recordset
        If spd.State Then spd.Close
        spd.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata3!bd_spread & "' ", Cn, 3, 2
        If Not spd.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata3!bd_spread & "  -  " & spd(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata3!bd_spread
        End If
        spd.Close
                If .TextMatrix(.Rows - 1, 3) = "NA  -  Not Applicable" And fldata3!bd_chk1 = 1 Then
                .TextMatrix(.Rows - 1, 3) = "NA  -  Progress"
                End If
        
        Dim jcg As New ADODB.Recordset
        
Dim ki5 As New ADODB.Recordset
If ki5.State Then ki5.Close
ki5.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & fldata3!bd_resccode & "' ", Cn, 3, 2
If Not ki5.EOF Then
 .TextMatrix(.Rows - 1, 4) = fldata3!bd_resccode & "  -  " & ki5(0)
Else
 .TextMatrix(.Rows - 1, 4) = fldata3!bd_resccode
End If
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata3!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 6) = fldata3!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 6) = fldata3!bd_costcode
        End If
        cs.Close
         If IsNull(fldata3!bd_sdate) = True Then
         .TextMatrix(.Rows - 1, 7) = ""
         Else
        .TextMatrix(.Rows - 1, 7) = Format(fldata3!bd_sdate, "dd/MM/yyyy H:mm:ss")
        End If
        If IsNull(fldata3!bd_edate) = True Then
        .TextMatrix(.Rows - 1, 8) = ""
        Else
        .TextMatrix(.Rows - 1, 8) = Format(fldata3!bd_edate, "dd/MM/yyyy H:mm:ss")
        End If
        .TextMatrix(.Rows - 1, 9) = fldata3!bd_qty
        .TextMatrix(.Rows - 1, 10) = fldata3!bd_days
        .TextMatrix(.Rows - 1, 11) = fldata3!bd_tqty
        .TextMatrix(.Rows - 1, 12) = fldata3!bd_uom
        .TextMatrix(.Rows - 1, 13) = fldata3!bd_curr
        .TextMatrix(.Rows - 1, 14) = fldata3!bd_unitrate
        .TextMatrix(.Rows - 1, 15) = Format(fldata3!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 16) = Format(fldata3!bd_extdamt, "###,###,###,###,##0.00")
        gtotal = gtotal + fldata3!bd_extdamt
        .TextMatrix(.Rows - 1, 17) = fldata3!bd_e_days
        .TextMatrix(.Rows - 1, 18) = fldata3!bd_e_tqty
        .TextMatrix(.Rows - 1, 19) = Format(fldata3!bd_e_extdamt, "###,###,##0.00")
        ntotal = ntotal + fldata3!bd_e_extdamt
        .TextMatrix(.Rows - 1, 20) = fldata3!bd_notes
        .TextMatrix(.Rows - 1, 21) = fldata3!bd_id
 fldata3.MoveNext
    
            'define fields as checkbox
            For Y = 1 To .Rows - 1
                    .Row = Y
                    .Col = 1
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellAlignment = flexAlignCenterCenter
                    .Text = strChecked
            Next Y
    Wend
End With

Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
txt_btotal.Text = Format(ntotal, "###,###,##0.00")

'Check if any row is selected

CheckSourceSelected
End Sub
Function LoadSourceSpreads()
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(spread_code),spread_desc from spreadmaster where spread_code <>'NA' order by spread_code", Cn, 3, 2
While Not spr.EOF
cbo_spr.AddItem spr(0) & "  -  " & spr(1)
spr.MoveNext
Wend
spr.Close
End Function
Function LoadResourcesforTarget()
'On Error Resume Next
bh = Split(cboTargetJobCharge.Text, "  -  ", Len(cboTargetJobCharge.Text), vbTextCompare)

Dim gtotal As Double
gtotal = 0
Dim ntotal As Double
ntotal = 0
Dim iddd As Double
iddd = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close

If lblTransactionType.Caption = "E" Then
fldata.Open "select * from cost where bd_type = '" & bh(2) & "' and bd_jobcharge='" & bh(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' and bd_spread <>'NA' ", Cn, 3, 2
ElseIf lblTransactionType.Caption = "B" Then
fldata.Open "select * from cost where bd_jobcharge='" & bh(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' and bd_spread <>'NA' ", Cn, 3, 2
End If

    While Not fldata.EOF
     iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_jobcharge, "  -  ", Len(fldata!bd_jobcharge), vbTextCompare)
Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
If lblTransactionType.Caption = "E" Then
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If
Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close

fldata2.Open "select * from cost where bd_type = '" & bh(2) & "' and bd_resccode='" & bh(0) & "' and bd_jobcharge='" & fldata!bd_jobcharge & "' and bd_costtype= '" & lblTransactionType.Caption & "'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

    If Not fldata2.EOF Then

            fldata2!bd_sdate = dt1
            fldata2!bd_edate = dt2
                    If dt1 <= main.DTPcutdate1.Value And dt2 <= main.DTPcutdate1.Value Then
                    a = dt2 - dt1
                    c = 0
                    ElseIf dt1 <= main.DTPcutdate1.Value And dt2 >= main.DTPcutdate1.Value Then
                    a = main.DTPcutdate1.Value - dt1
                    c = dt2 - main.DTPcutdate1.Value

                    Else
                    a = 0
                    c = dt2 - dt1
                    End If
            Dim d As Double
            d = 0
            Dim f As Double
            f = 0
            fldata2!bd_days = a
            fldata2!bd_e_days = c
            d = CDbl(a) * CDbl(fldata!bd_qty)
            fldata2!bd_e_tqty = CDbl(c) * CDbl(fldata!bd_qty)
            fldata2!bd_tqty = d
            fldata2!bd_extdamt = CDbl(d) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
            fldata2!bd_e_extdamt = CDbl(fldata2!bd_e_tqty) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
            fldata2.Update
    End If

ElseIf lblTransactionType.Caption = "B" Then
pp.Open "select * from budgeteddurationdetails where bdgt_spread_code='" & fldata!bd_spread & "' and bdgt_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
End If
        fldata.MoveNext
    Wend
        
Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where bd_jobcharge='" & bh(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' and bd_spread ='NA' ", Cn, 3, 2
While Not cd.EOF
If cd!bd_chk = 1 Then
                    If cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate <= main.DTPcutdate1.Value Then
                    a = cd!bd_edate - cd!bd_sdate
                    c = 0
                    ElseIf cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate >= main.DTPcutdate1.Value Then
                    a = main.DTPcutdate1.Value - cd!bd_sdate
                    c = cd!bd_edate - main.DTPcutdate1.Value
                    
                    Else
                    a = 0
                    c = cd!bd_edate - cd!bd_sdate
                    End If
                    cd!bd_days = a
                    cd!bd_e_days = c
                    If IsNull(cd!bd_days) = True Then
                    cd!bd_tqty = cd!bd_qty
                    Else
                    cd!bd_tqty = cd!bd_qty * cd!bd_days
                    End If
                    cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
                    If IsNull(cd!bd_e_days) = True Then
                    cd!bd_e_tqty = cd!bd_qty
                    Else
                    cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
                    End If
                    cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
ElseIf cd!bd_chk = 0 Then
 
                If cd!bd_chk1 = 0 Then
                
                cd!bd_edate = cd!bd_sdate
                                  
                                   If cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate <= main.DTPcutdate1.Value Then
                                                   cd!bd_tqty = cd!bd_qty
                                                   cd!bd_days = Null
                                                   cd!bd_e_days = 0
                                                   cd!bd_e_tqty = 0
                                      Else
                                      
                                                   cd!bd_e_tqty = cd!bd_qty
                                                   cd!bd_e_days = Null
                                                   cd!bd_days = 0
                                                   cd!bd_tqty = 0
                                   End If
                                   If IsNull(cd!bd_days) = True Then
                                   cd!bd_tqty = cd!bd_qty
                                   Else
                                   cd!bd_tqty = cd!bd_qty * cd!bd_days
                                   End If
                                   cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
                                   If IsNull(cd!bd_e_days) = True Then
                                   cd!bd_e_tqty = cd!bd_qty
                                   Else
                                   cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
                                   End If
                                   cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
                
                ElseIf cd!bd_chk1 = 1 Then
                            If IsNull(cd!bd_days) Then
                            cd!bd_tqty = cd!bd_qty
                            Else
                            cd!bd_tqty = cd!bd_qty * cd!bd_days
                            End If
                    cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
                            If IsNull(cd!bd_e_days) Then
                            cd!bd_e_tqty = cd!bd_qty
                            Else
                            cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
                            End If
                    cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
                End If
 End If
cd.Update
cd.MoveNext
Wend
On Error Resume Next
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close

If lblTransactionType.Caption = "E" Then
fldata3.Open "select * from cost  where  bd_type = '" & bh(2) & "' and bd_jobcharge='" & bh(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' order by bd_tranx,bd_spread,bd_sdate,bd_jobcharge,bd_costcode", Cn, 3, 2
ElseIf lblTransactionType.Caption = "B" Then
fldata3.Open "select * from cost  where  bd_jobcharge='" & bh(0) & "' and bd_costtype= '" & lblTransactionType.Caption & "' order by bd_tranx,bd_spread,bd_sdate,bd_jobcharge,bd_costcode", Cn, 3, 2
End If
With flxTargetResources
    .Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata3(0)
        .TextMatrix(.Rows - 1, 1) = fldata3!bd_type
         Dim spd As New ADODB.Recordset
        If spd.State Then spd.Close
        spd.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata3!bd_spread & "' ", Cn, 3, 2
        If Not spd.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata3!bd_spread & "  -  " & spd(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata3!bd_spread
        End If
        spd.Close
                If .TextMatrix(.Rows - 1, 2) = "NA  -  Not Applicable" And fldata3!bd_chk1 = 1 Then
                .TextMatrix(.Rows - 1, 2) = "NA  -  Progress"
                End If
        
        Dim jcg As New ADODB.Recordset
        
Dim ki5 As New ADODB.Recordset
If ki5.State Then ki5.Close
ki5.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & fldata3!bd_resccode & "' ", Cn, 3, 2
If Not ki5.EOF Then
 .TextMatrix(.Rows - 1, 3) = fldata3!bd_resccode & "  -  " & ki5(0)
Else
 .TextMatrix(.Rows - 1, 3) = fldata3!bd_resccode
End If
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata3!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 5) = fldata3!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 5) = fldata3!bd_costcode
        End If
        cs.Close
         If IsNull(fldata3!bd_sdate) = True Then
         .TextMatrix(.Rows - 1, 6) = ""
         Else
        .TextMatrix(.Rows - 1, 6) = Format(fldata3!bd_sdate, "dd/MM/yyyy H:mm:ss")
        End If
        If IsNull(fldata3!bd_edate) = True Then
        .TextMatrix(.Rows - 1, 7) = ""
        Else
        .TextMatrix(.Rows - 1, 7) = Format(fldata3!bd_edate, "dd/MM/yyyy H:mm:ss")
        End If
        .TextMatrix(.Rows - 1, 8) = fldata3!bd_qty
        .TextMatrix(.Rows - 1, 9) = fldata3!bd_days
        .TextMatrix(.Rows - 1, 10) = fldata3!bd_tqty
        .TextMatrix(.Rows - 1, 11) = fldata3!bd_uom
        .TextMatrix(.Rows - 1, 12) = fldata3!bd_curr
        .TextMatrix(.Rows - 1, 13) = fldata3!bd_unitrate
        .TextMatrix(.Rows - 1, 14) = Format(fldata3!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 15) = Format(fldata3!bd_extdamt, "###,###,###,###,##0.00")
        gtotal = gtotal + fldata3!bd_extdamt
        .TextMatrix(.Rows - 1, 16) = fldata3!bd_e_days
        .TextMatrix(.Rows - 1, 17) = fldata3!bd_e_tqty
        .TextMatrix(.Rows - 1, 18) = Format(fldata3!bd_e_extdamt, "###,###,##0.00")
        ntotal = ntotal + fldata3!bd_e_extdamt
        .TextMatrix(.Rows - 1, 19) = fldata3!bd_notes
 fldata3.MoveNext
    Wend
End With
Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
txt_btotal.Text = Format(ntotal, "###,###,##0.00")
End Function

Public Sub flxtitleTargetResources()
On Error Resume Next
    With flxTargetResources
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
       .TextMatrix(0, 1) = "Type"
        .ColWidth(1) = 500
        .TextMatrix(0, 2) = "Spread "
        .ColWidth(2) = 1100
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Resource"
        .ColWidth(3) = 3300
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "OBS"
        .ColWidth(4) = 600
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "CostCode"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Start Date"
        .ColWidth(6) = 2000
        .TextMatrix(0, 7) = "End Date"
        .ColWidth(7) = 2000
        .TextMatrix(0, 8) = "Qty"
        .ColWidth(8) = 1000
        .TextMatrix(0, 9) = "A Days"
        .ColWidth(9) = 650
        .TextMatrix(0, 10) = "Tot Qty"
        .ColWidth(10) = 1000
        .TextMatrix(0, 11) = "UOM "
        .ColWidth(11) = 600
        .TextMatrix(0, 12) = "Curcy "
        .ColWidth(12) = 600
        .TextMatrix(0, 13) = "UnitRate"
        .ColWidth(13) = 1100
        .TextMatrix(0, 14) = "Xrate"
        .ColWidth(14) = 500
        .TextMatrix(0, 15) = "ACWP Amount"
        .ColWidth(15) = 1500
        .TextMatrix(0, 16) = "E Days"
        .ColWidth(16) = 650
        .TextMatrix(0, 17) = "Tot Qty"
        .ColWidth(17) = 1000
        .TextMatrix(0, 18) = "ECTC Amount"
        .ColWidth(18) = 1500
        .TextMatrix(0, 19) = "Notes"
        .ColWidth(19) = 2500
    End With
End Sub
Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With flxSourceResources
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
        End With
        'Check if any row is selected
        CheckSourceSelected
End Sub
Function CheckSourceSelected()
Dim intRowsSelected As Integer
intRowsSelected = 0
For Y = 1 To flxSourceResources.Rows - 1
    If flxSourceResources.TextMatrix(Y, 1) = strChecked Then intRowsSelected = intRowsSelected + 1
Next Y
If intRowsSelected > 0 Then cmdImport.Enabled = True Else cmdImport.Enabled = False
End Function

Private Sub cbo_spr_Change()
'LoadSourceJobCharges
LoadJobCharges
End Sub

Private Sub cbo_spr_Click()
'LoadSourceJobCharges
LoadJobCharges
End Sub

Private Sub cboTargetJobCharge_Change()
'LoadResourcesforTarget
End Sub

Private Sub cboTargetJobCharge_Click()
LoadResourcesforTarget
End Sub
Private Sub cmdImport_Click()
c = 0
intRowsInserted = 0
spreadkey = Split(cbo_spr.Text, " - ", Len(cbo_spr.Text), vbTextCompare)
jcskey = Split(cboSourceJobCharge.Text, "  -  ", Len(cboSourceJobCharge.Text), vbTextCompare)
jctkey = Split(cboTargetJobCharge.Text, "  -  ", Len(cboTargetJobCharge), vbTextCompare)
If flxTargetResources.Rows >= 2 Then
response = MsgBox("Do you want to append resources for the above jobcharge?", vbYesNo, App.Title & " - Import Resources")
Else
response = MsgBox("Do you want to create new transaction for the above jobcharge?", vbYesNo, App.Title & " - Import Resources")
End If
If response = vbYes Then
For InsertIterator = 1 To flxSourceResources.Rows - 1
If flxSourceResources.TextMatrix(InsertIterator, 1) = strChecked Then
rowid = flxSourceResources.TextMatrix(InsertIterator, 21)
If lblTransactionType.Caption = "E" Then
InsertResource CDbl(rowid), Trim(spreadkey(0)), Trim(jcskey(0)), Trim(jctkey(0)), Trim(jcskey(2)), Trim(jctkey(2))
c = c + 1
ElseIf lblTransactionType.Caption = "B" Then
InsertResourceBudget CDbl(rowid), Trim(spreadkey(0)), Trim(jcskey(0)), Trim(jctkey(0))
c = c + 1
End If
End If
Next
If intRowsInserted = c Then MsgBox c & " Resources Imported Successfully", vbInformation, App.Title & " - Import Resources"
LoadResourcesforTarget
End If
End Sub
Private Sub flxSourceResources_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
        With flxSourceResources
         If .ColSel = 1 Then Call TriggerCheckbox(.Row, .Col)
        End With
    End If
End Sub

Private Sub flxSourceResources_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With flxSourceResources
            If .MouseRow <> 0 And .MouseCol <> 0 Then
             If .ColSel = 1 Then Call TriggerCheckbox(.MouseRow, .MouseCol)
            End If
        End With
    End If
End Sub
Public Sub flexTitleSource()
On Error Resume Next
    With flxSourceResources
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
       .TextMatrix(0, 1) = "Select"
        .ColWidth(1) = 500
       .TextMatrix(0, 2) = "Type"
        .ColWidth(2) = 500
        .TextMatrix(0, 3) = "Spread "
        .ColWidth(3) = 1100
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Resource"
        .ColWidth(4) = 3300
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "OBS"
        .ColWidth(5) = 600
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "CostCode"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Start Date"
        .ColWidth(7) = 2000
        .TextMatrix(0, 8) = "End Date"
        .ColWidth(8) = 2000
        .TextMatrix(0, 9) = "Qty"
        .ColWidth(9) = 1000
        .TextMatrix(0, 10) = "A Days"
        .ColWidth(10) = 650
        .TextMatrix(0, 11) = "Tot Qty"
        .ColWidth(11) = 1000
        .TextMatrix(0, 12) = "UOM "
        .ColWidth(12) = 600
        .TextMatrix(0, 13) = "Curcy "
        .ColWidth(13) = 600
        .TextMatrix(0, 14) = "UnitRate"
        .ColWidth(14) = 1100
        .TextMatrix(0, 15) = "Xrate"
        .ColWidth(15) = 500
        .TextMatrix(0, 16) = "ACWP Amount"
        .ColWidth(16) = 1500
        .TextMatrix(0, 17) = "E Days"
        .ColWidth(17) = 650
        .TextMatrix(0, 18) = "Tot Qty"
        .ColWidth(18) = 1000
        .TextMatrix(0, 19) = "ECTC Amount"
        .ColWidth(19) = 1500
        .TextMatrix(0, 20) = "Notes"
        .ColWidth(20) = 2500
    End With
End Sub
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub DTP_cod_Change()
 On Error Resume Next
'Call flex_data1
End Sub
Private Sub DTP_cod_Click()
On Error Resume Next
'Call flex_data1
End Sub
Private Sub cboSourceJobCharge_Change()
'LoadFlexSourceJCSources
End Sub
Private Sub cboSourceJobCharge_Click()
LoadFlexSourceJCSources
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(spread_code),spread_desc from spreadmaster where spread_code <>'NA' order by spread_code", Cn, 3, 2
While Not spr.EOF
cbo_spr.AddItem spr(0) & "  -  " & spr(1)
spr.MoveNext
Wend
spr.Close

main.lbltitle.Caption = "IMPORT FROM EXISTING JOBCHARGE"
DTP_cod.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
dtpTargetJobChargeStartDate.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
Call flexTitleSource
Call flxtitleTargetResources
'Call flex_data1
Me.Top = 5
Me.Left = 5
End Sub
Function LoadSourceJobCharges()
On Error Resume Next
flxSourceResources.Rows = 1
strSpreadKey = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
cboSourceJobCharge.Clear
Dim rsSourceJobCharge As New ADODB.Recordset
If rsSourceJobCharge.State Then rsSourceJobCharge.Close
rsSourceJobCharge.Open "select job_code, job_desc, prgs_type from jobcharge where job_code in (select prgs_job_key from progressdurationdetails where prgs_spread_code = '" & strSpreadKey(0) & "')", Cn, 3, 2
While Not rsSourceJobCharge.EOF
cboSourceJobCharge.AddItem rsSourceJobCharge(0) & "  -  " & rsSourceJobCharge(1) & "  -  " & rsSourceJobCharge(2)
rsSourceJobCharge.MoveNext
Wend
rsSourceJobCharge.Close
End Function
Function LoadJobCharges()
'On Error Resume Next
strSpreadKey = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
cboTargetJobCharge.Clear
Dim rsTargetJobCharge As New ADODB.Recordset
If rsTargetJobCharge.State Then rsTargetJobCharge.Close
If lblTransactionType.Caption = "E" Then
rsTargetJobCharge.Open "select pdd.prgs_job_key, jc.job_desc, prgs_type from progressdurationdetails pdd, jobcharge jc where pdd.prgs_job_key = job_code and prgs_spread_code = '" & strSpreadKey(0) & "'", Cn, 3, 2
While Not rsTargetJobCharge.EOF
cboSourceJobCharge.AddItem rsTargetJobCharge(0) & "  -  " & rsTargetJobCharge(1) & "  -  " & rsTargetJobCharge(2)
cboTargetJobCharge.AddItem rsTargetJobCharge(0) & "  -  " & rsTargetJobCharge(1) & "  -  " & rsTargetJobCharge(2)
rsTargetJobCharge.MoveNext
Wend
If rsTargetJobCharge.State Then rsTargetJobCharge.Close
Else
rsTargetJobCharge.Open "select bdd.bdgt_job_key, jc.job_desc from budgeteddurationdetails bdd, jobcharge jc where bdd.bdgt_job_key = job_code and bdgt_spread_code = '" & strSpreadKey(0) & "'", Cn, 3, 2
While Not rsTargetJobCharge.EOF
cboSourceJobCharge.AddItem rsTargetJobCharge(0) & "  -  " & rsTargetJobCharge(1)
cboTargetJobCharge.AddItem rsTargetJobCharge(0) & "  -  " & rsTargetJobCharge(1)
rsTargetJobCharge.MoveNext
Wend
If rsTargetJobCharge.State Then rsTargetJobCharge.Close
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
main.lbltitle.Caption = ""
Unload estimatedincurredcost
Unload Me
End Sub

