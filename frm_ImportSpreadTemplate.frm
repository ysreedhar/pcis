VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ImportSpreadTemplate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "IMPORT SPREAD TEMPLATE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
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
      Left            =   10320
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox cbo_resc 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.ComboBox cboSpreadTemplate 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   240
      Top             =   8400
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
            Picture         =   "frm_ImportSpreadTemplate.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   8400
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
            Picture         =   "frm_ImportSpreadTemplate.frx":13274
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":13386
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":137D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":13C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1407C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":144CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1A768
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1AA82
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1AD9C
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1B336
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1B8D0
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1BE6A
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1C404
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1C516
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1CA58
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1CFF2
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1D58C
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1DE66
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1DF78
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1E08A
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1E19C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1E2AE
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1E3C0
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1E4D2
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1EA6C
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1F006
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1F5A0
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1FB3A
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1FC4C
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":1FD5E
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":202F8
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":2040A
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":2051C
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":20AB6
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":20BC8
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":21162
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":216FC
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":2180E
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":21DA8
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":22342
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":228DC
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":229EE
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":22F88
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":2309A
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":231AC
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":232BE
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":233D0
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":234E2
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":23A7C
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":23B8E
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":23CA0
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":2423A
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":247D4
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":24D6E
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":25308
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":258A2
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":25E3C
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ImportSpreadTemplate.frx":263D6
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   7815
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13785
      _Version        =   393216
      Rows            =   3
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
   Begin VB.Label Label16 
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
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Spread Template"
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
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frm_ImportSpreadTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sprcode As String
Dim jchrg As String
Dim ntotal As Double


Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub cbo_pproj_Change()
On Error Resume Next
kl1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
            gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
            cbo_resc.Clear
            Dim rs As New ADODB.Recordset
            rs.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & gg(0) & "' order by job_code", Cn, 3, 2
            While Not rs.EOF
            cbo_resc.AddItem rs(0) & "  -  " & rs(1)
            rs.MoveNext
            Wend
            rs.Close
End Sub

Private Sub cbo_pproj_Click()
 On Error Resume Next
        gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
        cbo_resc.Clear
        Dim rs As New ADODB.Recordset
        rs.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & gg(0) & "' order by job_code", Cn, 3, 2
        While Not rs.EOF
        cbo_resc.AddItem rs(0) & "  -  " & rs(1)
        rs.MoveNext
        Wend
        rs.Close
End Sub

Private Sub cbo_resc_Change()
On Error Resume Next
        textrescname.Text = ""
        textresccode.Text = ""
        textprojkey.Text = ""
        txt_projdesc.Text = ""
        txt_brate.Text = ""
        textcosttype.Text = ""
        txt_vendor.Text = ""
        txt_respcode.Text = ""
        txt_respname.Text = ""
        Text3.Text = ""
        txt_crate.Text = ""
        Text4.Text = ""
kl = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
nj = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & nj(0) & "' ", Cn, 3, 2
        
        If Not fl.EOF Then
        textrescname.Text = fl!resc_desc
        textresccode.Text = fl!resc_code
        textprojkey.Text = nj(0)
        txt_projdesc.Text = nj(1)
        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        textcosttype.Text = "E"
        txt_vendor.Text = fl!resc_vendorcode
        txt_respcode.Text = fl!resc_respcode
        Dim rr As New ADODB.Recordset
        If rr.State Then rr.Close
        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        txt_respname.Text = rr(0)
        End If
        Text3.Text = fl!dresc_curcy
        
         
        End If
fl.Close

Dim fl1 As New ADODB.Recordset
        If fl1.State Then fl1.Close
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close


Call flex_data1
End Sub

Public Sub flex_data1()
  'On Error Resume Next
bh = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
Pi = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim ntotal As Double
ntotal = 0
Dim iddd As Double
iddd = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where bd_jobcharge='" & bh(0) & "' and bd_projectkey='" & Pi(0) & "' and bd_year= '" & cbo_year.Text & "' and bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2


    While Not fldata.EOF
     iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_jobcharge, "  -  ", Len(fldata!bd_jobcharge), vbTextCompare)


Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If

Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
fldata2.Open "select * from cost where bd_resccode='" & bh(0) & "' and bd_year= '" & cbo_year.Text & "' and bd_jobcharge='" & fldata!bd_jobcharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

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
        fldata.MoveNext
    Wend
Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where bd_jobcharge='" & bh(0) & "'   and bd_year= '" & cbo_year.Text & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
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
fldata3.Open "select * from cost  where bd_jobcharge='" & bh(0) & "' and bd_projectkey='" & Pi(0) & "'  and bd_year= '" & cbo_year.Text & "' and bd_costtype='E' order by bd_tranx,bd_spread,bd_sdate,bd_jobcharge,bd_costcode", Cn, 3, 2
With flex_grid
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
End Sub

Public Sub flex_title1()
On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
If frm_estimated_layout.List1.Selected(0) = True Then
       .TextMatrix(0, 1) = "Type"
        .ColWidth(1) = 500
Else
       .ColWidth(1) = 0

End If
If frm_estimated_layout.List1.Selected(1) = True Then
        .TextMatrix(0, 2) = "Spread "
        .ColWidth(2) = 1100
        .ColAlignment(2) = 0
Else
       .ColWidth(2) = 0

End If
If frm_estimated_layout.List1.Selected(2) = True Then
        .TextMatrix(0, 3) = "Resource"
        .ColWidth(3) = 3300
        .ColAlignment(3) = 0
Else
       .ColWidth(3) = 0
End If
If frm_estimated_layout.List1.Selected(3) = True Then
        .TextMatrix(0, 4) = "OBS"
        .ColWidth(4) = 600
        .ColAlignment(4) = 0
Else
       .ColWidth(4) = 0

End If
If frm_estimated_layout.List1.Selected(4) = True Then
        .TextMatrix(0, 5) = "CostCode"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
Else
       .ColWidth(5) = 0

End If
If frm_estimated_layout.List1.Selected(5) = True Then
        .TextMatrix(0, 6) = "Start Date"
        .ColWidth(6) = 2000
Else
       .ColWidth(6) = 0

End If
If frm_estimated_layout.List1.Selected(6) = True Then
        .TextMatrix(0, 7) = "End Date"
        .ColWidth(7) = 2000
Else
      .ColWidth(7) = 0

End If
If frm_estimated_layout.List1.Selected(7) = True Then
        .TextMatrix(0, 8) = "Qty"
        .ColWidth(8) = 520
Else
      .ColWidth(8) = 0

End If
If frm_estimated_layout.List1.Selected(8) = True Then
        .TextMatrix(0, 9) = "A Days"
        .ColWidth(9) = 650
Else
      .ColWidth(9) = 0

End If
If frm_estimated_layout.List1.Selected(9) = True Then
        .TextMatrix(0, 10) = "Tot Qty"
        .ColWidth(10) = 600
Else
      .ColWidth(10) = 0

End If
If frm_estimated_layout.List1.Selected(10) = True Then
        .TextMatrix(0, 11) = "UOM "
        .ColWidth(11) = 600
Else
      .ColWidth(11) = 0

End If
If frm_estimated_layout.List1.Selected(11) = True Then
        .TextMatrix(0, 12) = "Curcy "
        .ColWidth(12) = 600
Else
      .ColWidth(12) = 0

End If
If frm_estimated_layout.List1.Selected(12) = True Then
        .TextMatrix(0, 13) = "UnitRate"
        .ColWidth(13) = 1100
Else
      .ColWidth(13) = 0

End If
If frm_estimated_layout.List1.Selected(13) = True Then
        .TextMatrix(0, 14) = "Xrate"
        .ColWidth(14) = 500
Else
      .ColWidth(14) = 0

End If
If frm_estimated_layout.List1.Selected(14) = True Then
        .TextMatrix(0, 15) = "ACWP Amount"
        .ColWidth(15) = 1500
Else
      .ColWidth(15) = 0

End If
If frm_estimated_layout.List1.Selected(15) = True Then
        .TextMatrix(0, 16) = "E Days"
        .ColWidth(16) = 650
Else
      .ColWidth(16) = 0

End If
If frm_estimated_layout.List1.Selected(16) = True Then
        .TextMatrix(0, 17) = "Tot Qty"
        .ColWidth(17) = 800
Else
      .ColWidth(17) = 0

End If
If frm_estimated_layout.List1.Selected(17) = True Then
        .TextMatrix(0, 18) = "ECTC Amount"
        .ColWidth(18) = 1500
Else
      .ColWidth(18) = 0

End If
If frm_estimated_layout.List1.Selected(18) = True Then
        .TextMatrix(0, 19) = "Notes"
        .ColWidth(19) = 2500
 Else
      .ColWidth(19) = 0

End If
    End With
End Sub
Public Sub flex_title()
On Error Resume Next
    With flex_grid
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

Private Sub cbo_resc_Click()
On Error Resume Next
        textrescname.Text = ""
        textresccode.Text = ""
        textprojkey.Text = ""
        txt_projdesc.Text = ""
        txt_brate.Text = ""
        textcosttype.Text = ""
        txt_vendor.Text = ""
        txt_respcode.Text = ""
        txt_respname.Text = ""
        Text3.Text = ""
        txt_crate.Text = ""
        Text4.Text = ""
kl = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
nj = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & nj(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        textrescname.Text = fl!resc_desc
        textresccode.Text = fl!resc_code
        textprojkey.Text = nj(0)
        txt_projdesc.Text = nj(1)
        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        textcosttype.Text = "E"
        txt_vendor.Text = fl!resc_vendorcode
        txt_respcode.Text = fl!resc_respcode
        Dim rr As New ADODB.Recordset
        If rr.State Then rr.Close
        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        txt_respname.Text = rr(0)
        End If
        Text3.Text = fl!dresc_curcy
        End If
fl.Close
Dim fl1 As New ADODB.Recordset
        If fl1.State Then fl1.Close
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close
Call flex_data1
End Sub

Private Sub cbo_year_Change()
On Error Resume Next
cbo_pproj.Clear
cbo_resc.Clear
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
    cbo_pproj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
End Sub

Private Sub cbo_year_Click()
On Error Resume Next
cbo_pproj.Clear
cbo_resc.Clear
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
    cbo_pproj.AddItem pr(0) & "  -  " & pr(1)
    pr.MoveNext
Wend
pr.Close
 
 
End Sub

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub cboSpreadTemplate_Change()
LoadFlexData
End Sub
Function LoadFlexData()
With flex_grid
If flex_grid.TextMatrix(.Rows - 1, 2) <> "-" Then
.Rows = .Rows + 1
.TextMatrix(.Rows - 1, 0) = "-"
.TextMatrix(.Rows - 1, 1) = "-"
.TextMatrix(.Rows - 1, 2) = "-"
.TextMatrix(.Rows - 1, 3) = "-"
.TextMatrix(.Rows - 1, 4) = "-"
.TextMatrix(.Rows - 1, 5) = "-"
.TextMatrix(.Rows - 1, 6) = "-"
.TextMatrix(.Rows - 1, 7) = "-"
.TextMatrix(.Rows - 1, 8) = "-"
.TextMatrix(.Rows - 1, 9) = "-"
.TextMatrix(.Rows - 1, 10) = "-"
.TextMatrix(.Rows - 1, 11) = "-"
.TextMatrix(.Rows - 1, 12) = "-"
.TextMatrix(.Rows - 1, 13) = "-"
.TextMatrix(.Rows - 1, 14) = "-"
.TextMatrix(.Rows - 1, 15) = "-"
.TextMatrix(.Rows - 1, 16) = "-"
.TextMatrix(.Rows - 1, 17) = "-"
.TextMatrix(.Rows - 1, 18) = "-"
.TextMatrix(.Rows - 1, 19) = "-"
.CellBackColor = vbCyan
.CellForeColor = vbWhite
End If
End With
End Function
Private Sub cboSpreadTemplate_Click()
LoadFlexData
End Sub
Private Sub DTP_cod_Change()
 On Error Resume Next
Call flex_data1
End Sub
Private Sub DTP_cod_Click()
 On Error Resume Next
Call flex_data1
End Sub

Private Sub flex_grid_Click()
On Error Resume Next
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
Next
flex_grid.Col = 1
vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
Unload estimatedincurredcost
estimatedincurredcost.Top = 3200
estimatedincurredcost.Left = 0
estimatedincurredcost.Height = 3915
estimatedincurredcost.Width = 9645
estimatedincurredcost.Show
' back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
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
Next
flex_grid.Col = 1
 
Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
'estimatedincurredcost.Show

Dim rsdd  As New ADODB.Recordset
If rsdd.State Then rsd.Close
rsdd.Open "select * from cost  where bd_id=" & ID, Cn, 3, 2
If Not rsdd.EOF Then
        cbo_year.Text = rsdd!bd_year
        textresccode.Text = rsdd!bd_resccode
        textrescname.Text = rsdd!bd_rescname
        txt_vendor.Text = rsdd!bd_vendor
        txt_brate.Text = Format(rsdd!bd_brate, "###,###,##0.00")
        txt_crate.Text = Format(rsdd!bd_crate, "###,###,##0.00")
        textprojkey.Text = rsdd!bd_projectkey
        txt_projdesc.Text = rsdd!bd_projectdesc
        textcosttype.Text = rsdd!bd_costtype
        txt_respcode.Text = rsdd!bd_respcode
        txt_respname.Text = rsdd!bd_respname
       ' DTP_cod.Value = rsdd!bd_cuttdate
        Dim spd1 As New ADODB.Recordset
        If spd1.State Then spd1.Close
        spd1.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & rsdd!bd_spread & "' ", Cn, 3, 2
        If Not spd1.EOF Then
        estimatedincurredcost.cbo_spread.Text = rsdd!bd_spread & "  -  " & spd1(0)
        Else
        estimatedincurredcost.cbo_spread.Text = rsdd!bd_spread
        End If
        spd1.Close
If estimatedincurredcost.cbo_spread.Text = "NA  -  Not Applicable" And rsdd!bd_chk1 = 1 Then
estimatedincurredcost.cbo_spread.Text = "NA  -  Progress"
End If
        estimatedincurredcost.cbo_tranx.Text = rsdd!bd_tranx
        Dim jcg1 As New ADODB.Recordset
        If jcg1.State Then jcg1.Close
        jcg1.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & rsdd!bd_jobcharge & "' ", Cn, 3, 2
        If Not jcg1.EOF Then
        estimatedincurredcost.cbo_jobcharge.Text = rsdd!bd_jobcharge & "  -  " & jcg1(0)
        Else
        estimatedincurredcost.cbo_jobcharge.Text = rsdd!bd_jobcharge
        End If
        jcg1.Close
         Dim cs1 As New ADODB.Recordset
        If cs1.State Then cs1.Close
        cs1.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & rsdd!bd_costcode & "' ", Cn, 3, 2
        If Not cs1.EOF Then
        estimatedincurredcost.cbo_costcode.Text = rsdd!bd_costcode & "  -  " & cs1(0)
        Else
        estimatedincurredcost.cbo_costcode.Text = rsdd!bd_costcode
        End If
        cs1.Close
        estimatedincurredcost.txt_qty.Text = rsdd!bd_qty
        estimatedincurredcost.txt_days.Text = rsdd!bd_days
        estimatedincurredcost.txt_tqty.Text = rsdd!bd_tqty
        estimatedincurredcost.cbo_uom.Text = rsdd!bd_uom
        estimatedincurredcost.cbo_curr.Text = rsdd!bd_curr
        estimatedincurredcost.txt_Xrate.Text = rsdd!bd_xchg
        estimatedincurredcost.txt_unitrate.Text = Format(rsdd!bd_unitrate, "###,###,##0.00")
        estimatedincurredcost.txt_Extdamt.Text = Format(rsdd!bd_extdamt, "###,###,##0.00")
        estimatedincurredcost.txt_note.Text = rsdd!bd_notes
        estimatedincurredcost.cbo_obs.Text = rsdd!bd_obs
                                If IsNull(rsdd!bd_e_days) = True Then
                                estimatedincurredcost.txt_edays.Text = ""
                                Else
                                estimatedincurredcost.txt_edays.Text = rsdd!bd_e_days
                                End If
        estimatedincurredcost.txt_etqty.Text = rsdd!bd_e_tqty
        estimatedincurredcost.txt_ectcamt.Text = Format(rsdd!bd_e_extdamt, "###,###,##0.00")
                        If IsNull(rsdd!bd_sdate) = False Then
                        estimatedincurredcost.DTP_ed.Value = rsdd!bd_edate
                        Else
                        estimatedincurredcost.DTP_ed.Value = Date
                        End If
                    If IsNull(rsdd!bd_edate) = False Then
                    estimatedincurredcost.DTP_sd.Value = rsdd!bd_sdate
                    Else
                    estimatedincurredcost.DTP_sd.Value = Date
                    End If
        
                If rsdd!bd_chk = 1 Then
                estimatedincurredcost.Check1.Value = 1
                Else
                estimatedincurredcost.Check1.Value = 0
                End If
               estimatedincurredcost.cbo_type.Text = rsdd!bd_type
End If
If estimatedincurredcost.cbo_spread.Text = "NA  -  Not Applicable" Then
                If estimatedincurredcost.Check1.Value = 0 Then
                estimatedincurredcost.DTP_ed.Enabled = 0
                estimatedincurredcost.Check2.Visible = True
                Else
                        estimatedincurredcost.DTP_sd.Enabled = True
                        estimatedincurredcost.DTP_ed.Enabled = True
                        estimatedincurredcost.Check1.Visible = True
                        estimatedincurredcost.lbl.Visible = True
                        estimatedincurredcost.Check2.Visible = True
                End If
ElseIf estimatedincurredcost.cbo_spread.Text = "NA  -  Progress" Then
        estimatedincurredcost.Check2.Value = 1
        estimatedincurredcost.Check2.Visible = True
        estimatedincurredcost.txt_days.Enabled = True
        estimatedincurredcost.txt_edays.Enabled = True
        estimatedincurredcost.Check1.Visible = True
        estimatedincurredcost.DTP_sd.Visible = True
        estimatedincurredcost.DTP_ed.Visible = True
Else
    estimatedincurredcost.DTP_sd.Enabled = False
    estimatedincurredcost.DTP_ed.Enabled = False
    estimatedincurredcost.Check1.Visible = False
    estimatedincurredcost.lbl.Visible = False
End If

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
LoadSpreadTemplates
main.lbltitle.Caption = "IMPORT SPREAD TEMPLATE"
DTP_cod.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
Call flex_title
Call flex_data1
Me.Top = 5
Me.Left = 5
Toolbar1.Visible = False
Dim i As Integer
i = 0
For i = 2000 To 2050
cbo_year.AddItem i
Next i
'load resource
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Width = 11415
Me.Height = 9750
End Sub
Function LoadSpreadTemplates()
On Error Resume Next
cboSpreadTemplate.Clear
Dim rsSpreadTemplates As New ADODB.Recordset
If rsSpreadTemplates.State Then rsSpreadTemplates.Close
rsSpreadTemplates.Open "select SpreadCode from tblSpreadTemplate", Cn, 3, 2
While Not rsSpreadTemplates.EOF
    cboSpreadTemplate.AddItem rsSpreadTemplates(0)
    rsSpreadTemplates.MoveNext
Wend
rsSpreadTemplates.Close
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
main.lbltitle.Caption = ""
Unload estimatedincurredcost
Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
If cbo_year.Text = "" Then
    MsgBox "Select Year"
    cbo_year.SetFocus
Exit Sub
End If
If cboJobCharge.Text = "" Then
    MsgBox "Select Job Charge"
    cboJobCharge.SetFocus
Exit Sub
End If
'If cboSpreadTemplate.Text = "" Then
'MsgBox "select Resource"
'cboSpreadTemplate.SetFocus
'Exit Sub
'End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload estimatedincurredcost
estimatedincurredcost.Top = 3200
estimatedincurredcost.Left = 0
estimatedincurredcost.Height = 3915
estimatedincurredcost.Width = 9645
estimatedincurredcost.Show
Dim uid As Double
uid = 0
' to save new record
ElseIf Button.Caption = "Save" Then
If estimatedincurredcost.cbo_spread.Text = "" Then
MsgBox "Select Spread"
estimatedincurredcost.cbo_spread.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_tranx.Text = "" Then
MsgBox "Select Tranx"
estimatedincurredcost.cbo_tranx.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_type.Text = "" Then
MsgBox "Select SUB-JC"
estimatedincurredcost.cbo_type.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_jobcharge.Text = "" Then
MsgBox "Select Jobcharge"
estimatedincurredcost.cbo_jobcharge.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_obs.Text = "" Then
MsgBox "Select OBS Code"
estimatedincurredcost.cbo_obs.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
estimatedincurredcost.cbo_costcode.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_qty.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_uom.Text = "" Then
MsgBox "Select UOM"
estimatedincurredcost.cbo_uom.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_curr.Text = "" Then
MsgBox "Select Currency"
estimatedincurredcost.cbo_curr.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_unitrate.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_unitrate.SetFocus
Exit Sub
End If
'On Error Resume Next
es = Split(estimatedincurredcost.cbo_spread.Text, "  -  ", Len(estimatedincurredcost.cbo_spread.Text), vbTextCompare)
es1 = Split(estimatedincurredcost.cbo_jobcharge.Text, "  -  ", Len(estimatedincurredcost.cbo_jobcharge.Text), vbTextCompare)
es2 = Split(estimatedincurredcost.cbo_costcode.Text, "  -  ", Len(estimatedincurredcost.cbo_costcode.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from cost ", Cn, 3, 2
sv.AddNew
        sv!bd_year = cbo_year.Text
        sv!bd_resccode = textresccode.Text
        sv!bd_rescname = textrescname.Text
        sv!bd_vendor = txt_vendor.Text
        sv!bd_brate = txt_brate.Text
        sv!bd_crate = txt_crate.Text
        sv!bd_projectkey = textprojkey.Text
        sv!bd_projectdesc = txt_projdesc.Text
        sv!bd_costtype = textcosttype.Text
        sv!bd_respcode = txt_respcode.Text
        sv!bd_respname = txt_respname.Text
        sv!bd_cuttdate = DTP_cod.Value
        sv!bd_spread = es(0)
        sv!bd_tranx = estimatedincurredcost.cbo_tranx.Text
        sv!bd_jobcharge = es1(0)
        sv!bd_costcode = es2(0)
        If estimatedincurredcost.txt_qty.Text = "" Then
        sv!bd_qty = 0
        Else
        sv!bd_qty = estimatedincurredcost.txt_qty.Text
        End If
        If estimatedincurredcost.txt_days.Text = "" Then
        sv!bd_days = 0
        Else
        sv!bd_days = estimatedincurredcost.txt_days.Text
        End If
        sv!bd_tqty = estimatedincurredcost.txt_tqty.Text
        sv!bd_uom = estimatedincurredcost.cbo_uom.Text
        sv!bd_curr = estimatedincurredcost.cbo_curr.Text
        sv!bd_xchg = estimatedincurredcost.txt_Xrate.Text
        sv!bd_unitrate = estimatedincurredcost.txt_unitrate.Text
        sv!bd_extdamt = estimatedincurredcost.txt_Extdamt.Text
        If estimatedincurredcost.txt_edays.Text = "" Then
        sv!bd_e_days = 0
        Else
        sv!bd_e_days = estimatedincurredcost.txt_edays.Text
        End If
        If estimatedincurredcost.txt_etqty.Text = "" Then
        sv!bd_e_tqty = 0
        Else
        sv!bd_e_tqty = estimatedincurredcost.txt_etqty.Text
        End If
        sv!bd_e_extdamt = estimatedincurredcost.txt_ectcamt.Text
        sv!bd_edate = estimatedincurredcost.DTP_ed.Value
        sv!bd_sdate = estimatedincurredcost.DTP_sd.Value
        sv!bd_notes = estimatedincurredcost.txt_note.Text
        If estimatedincurredcost.Check1.Value = 1 Then
        sv!bd_chk = 1
        Else
        sv!bd_chk = 0
        End If
        If estimatedincurredcost.Check2.Value = 1 Then
        sv!bd_chk1 = 1
        Else
        sv!bd_chk1 = 0
        End If
        sv!t_date = estimatedincurredcost.DTP_tdate.Value
        sv!u_date = Now
        sv!t_user = main.Label2.Caption
        sv!bd_type = estimatedincurredcost.cbo_type.Text
        sv!bd_obs = estimatedincurredcost.cbo_obs.Text
sv.Update
sv.Close
Call flex_data1
 'for BCWP
MsgBox "New Record Added Succesfully"
Unload estimatedincurredcost
 '----------------END--------------

'to modify existing record
ElseIf Button.Caption = "Modify" Then
If estimatedincurredcost.cbo_spread.Text = "" Then
MsgBox "Select Spread"
estimatedincurredcost.cbo_spread.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_tranx.Text = "" Then
MsgBox "Select Tranx"
estimatedincurredcost.cbo_tranx.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_type.Text = "" Then
MsgBox "Select SUB-JC"
estimatedincurredcost.cbo_type.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_jobcharge.Text = "" Then
MsgBox "Select Jobcharge"
estimatedincurredcost.cbo_jobcharge.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_obs.Text = "" Then
MsgBox "Select OBS Code"
estimatedincurredcost.cbo_obs.SetFocus
Exit Sub
End If

If estimatedincurredcost.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
estimatedincurredcost.cbo_costcode.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_qty.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_uom.Text = "" Then
MsgBox "Select UOM"
estimatedincurredcost.cbo_uom.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_curr.Text = "" Then
MsgBox "Select Currency"
estimatedincurredcost.cbo_curr.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_unitrate.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_unitrate.SetFocus
Exit Sub
End If
On Error Resume Next
es = Split(estimatedincurredcost.cbo_spread.Text, "  -  ", Len(estimatedincurredcost.cbo_spread.Text), vbTextCompare)
es1 = Split(estimatedincurredcost.cbo_jobcharge.Text, "  -  ", Len(estimatedincurredcost.cbo_jobcharge.Text), vbTextCompare)
es2 = Split(estimatedincurredcost.cbo_costcode.Text, "  -  ", Len(estimatedincurredcost.cbo_costcode.Text), vbTextCompare)

Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from cost  where bd_id=" & id1, Cn, 3, 2
If Not md.EOF Then
        md!bd_year = cbo_year.Text
        md!bd_resccode = textresccode.Text
        md!bd_rescname = textrescname.Text
        md!bd_vendor = txt_vendor.Text
        md!bd_brate = txt_brate.Text
        md!bd_crate = txt_crate.Text
        md!bd_projectkey = textprojkey.Text
        md!bd_projectdesc = txt_projdesc.Text
        md!bd_costtype = textcosttype.Text
        md!bd_respcode = txt_respcode.Text
        md!bd_respname = txt_respname.Text
        md!bd_cuttdate = DTP_cod.Value
        md!bd_spread = es(0)
        md!bd_tranx = estimatedincurredcost.cbo_tranx.Text
        md!bd_jobcharge = es1(0)
        md!bd_costcode = es2(0)
        If estimatedincurredcost.txt_qty.Text = "" Then
        md!bd_qty = 0
        Else
        md!bd_qty = estimatedincurredcost.txt_qty.Text
        End If
        If estimatedincurredcost.txt_days.Text = "" Then
        md!bd_days = 0
        Else
        md!bd_days = estimatedincurredcost.txt_days.Text
        End If
        md!bd_tqty = estimatedincurredcost.txt_tqty.Text
        md!bd_uom = estimatedincurredcost.cbo_uom.Text
        md!bd_curr = estimatedincurredcost.cbo_curr.Text
        md!bd_xchg = estimatedincurredcost.txt_Xrate.Text
        md!bd_unitrate = estimatedincurredcost.txt_unitrate.Text
        md!bd_extdamt = estimatedincurredcost.txt_Extdamt.Text
        If estimatedincurredcost.txt_edays.Text = "" Then
        md!bd_e_days = 0
        Else
        md!bd_e_days = estimatedincurredcost.txt_edays.Text
        End If
        If estimatedincurredcost.txt_etqty.Text = "" Then
        md!bd_e_tqty = 0
        Else
        md!bd_e_tqty = estimatedincurredcost.txt_etqty.Text
        End If
        md!bd_e_extdamt = estimatedincurredcost.txt_ectcamt.Text
        md!bd_edate = estimatedincurredcost.DTP_ed.Value
        md!bd_sdate = estimatedincurredcost.DTP_sd.Value
        md!bd_notes = estimatedincurredcost.txt_note.Text
        If estimatedincurredcost.Check1.Value = 1 Then
        md!bd_chk = 1
        Else
        md!bd_chk = 0
        End If
        If estimatedincurredcost.Check2.Value = 1 Then
        md!bd_chk1 = 1
        Else
        md!bd_chk1 = 0
        End If
        md!t_date = estimatedincurredcost.DTP_tdate.Value
        md!u_date = Now
        md!t_user = main.Label2.Caption
        md!bd_type = estimatedincurredcost.cbo_type.Text
        md!bd_obs = estimatedincurredcost.cbo_obs.Text
        md.Update
md.Close
End If
MsgBox "Selected Record Modified Successfully"
'Unload estimatedincurredcost
Call flex_data1


'-----------END----------
'to delete
ElseIf Button.Caption = "Delete" Then
Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
Dim id3 As Double
id3 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from cost where bd_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload estimatedincurredcost
Call flex_data1
Else
Unload estimatedincurredcost

End If
ElseIf Button.Caption = "Duplicate" Then
If estimatedincurredcost.cbo_spread.Text = "" Then
MsgBox "Select Spread"
estimatedincurredcost.cbo_spread.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_tranx.Text = "" Then
MsgBox "Select Tranx"
estimatedincurredcost.cbo_tranx.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_type.Text = "" Then
MsgBox "Select SUB-JC"
estimatedincurredcost.cbo_type.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_jobcharge.Text = "" Then
MsgBox "Select Jobcharge"
estimatedincurredcost.cbo_jobcharge.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_obs.Text = "" Then
MsgBox "Select OBS Code"
estimatedincurredcost.cbo_obs.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
estimatedincurredcost.cbo_costcode.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_qty.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_uom.Text = "" Then
MsgBox "Select UOM"
estimatedincurredcost.cbo_uom.SetFocus
Exit Sub
End If
If estimatedincurredcost.cbo_curr.Text = "" Then
MsgBox "Select Currency"
estimatedincurredcost.cbo_curr.SetFocus
Exit Sub
End If
If estimatedincurredcost.txt_unitrate.Text = "" Then
MsgBox "Enter Quantity"
estimatedincurredcost.txt_unitrate.SetFocus
Exit Sub
End If
'On Error Resume Next
es = Split(estimatedincurredcost.cbo_spread.Text, "  -  ", Len(estimatedincurredcost.cbo_spread.Text), vbTextCompare)
es1 = Split(estimatedincurredcost.cbo_jobcharge.Text, "  -  ", Len(estimatedincurredcost.cbo_jobcharge.Text), vbTextCompare)
es2 = Split(estimatedincurredcost.cbo_costcode.Text, "  -  ", Len(estimatedincurredcost.cbo_costcode.Text), vbTextCompare)
Dim svv As New ADODB.Recordset
If svv.State Then svv.Close
svv.Open "select * from cost ", Cn, 3, 2
svv.AddNew
        svv!bd_year = cbo_year.Text
        svv!bd_resccode = textresccode.Text
        svv!bd_rescname = textrescname.Text
        svv!bd_vendor = txt_vendor.Text
        svv!bd_brate = txt_brate.Text
        svv!bd_crate = txt_crate.Text
        svv!bd_projectkey = textprojkey.Text
        svv!bd_projectdesc = txt_projdesc.Text
        svv!bd_costtype = textcosttype.Text
        svv!bd_respcode = txt_respcode.Text
        svv!bd_respname = txt_respname.Text
        svv!bd_cuttdate = DTP_cod.Value
        svv!bd_spread = es(0)
        svv!bd_tranx = estimatedincurredcost.cbo_tranx.Text
        svv!bd_jobcharge = es1(0)
        svv!bd_costcode = es2(0)
        If estimatedincurredcost.txt_qty.Text = "" Then
        svv!bd_qty = 0
        Else
        svv!bd_qty = estimatedincurredcost.txt_qty.Text
        End If
        If estimatedincurredcost.txt_days.Text = "" Then
        svv!bd_days = 0
        Else
        svv!bd_days = estimatedincurredcost.txt_days.Text
        End If
        svv!bd_tqty = estimatedincurredcost.txt_tqty.Text
        svv!bd_uom = estimatedincurredcost.cbo_uom.Text
        svv!bd_curr = estimatedincurredcost.cbo_curr.Text
        svv!bd_xchg = estimatedincurredcost.txt_Xrate.Text
        svv!bd_unitrate = estimatedincurredcost.txt_unitrate.Text
        svv!bd_extdamt = estimatedincurredcost.txt_Extdamt.Text
        If estimatedincurredcost.txt_edays.Text = "" Then
        svv!bd_e_days = 0
        Else
        svv!bd_e_days = estimatedincurredcost.txt_edays.Text
        End If
        If estimatedincurredcost.txt_etqty.Text = "" Then
        svv!bd_e_tqty = 0
        Else
        svv!bd_e_tqty = estimatedincurredcost.txt_etqty.Text
        End If
        svv!bd_e_extdamt = estimatedincurredcost.txt_ectcamt.Text
        svv!bd_edate = estimatedincurredcost.DTP_ed.Value
        svv!bd_sdate = estimatedincurredcost.DTP_sd.Value
        svv!bd_notes = estimatedincurredcost.txt_note.Text
        If estimatedincurredcost.Check1.Value = 1 Then
        svv!bd_chk = 1
        Else
        svv!bd_chk = 0
        End If
            If estimatedincurredcost.Check2.Value = 1 Then
            svv!bd_chk1 = 1
            Else
            svv!bd_chk1 = 0
            End If
        svv!t_date = estimatedincurredcost.DTP_tdate.Value
        svv!u_date = Now
        svv!t_user = main.Label2.Caption
        svv!bd_type = estimatedincurredcost.cbo_type.Text
        svv!bd_obs = estimatedincurredcost.cbo_obs.Text
        
svv.Update
svv.Close

Call flex_data1


 'for BCWP
MsgBox "New Record Added Succesfully"
Unload estimatedincurredcost
 
 
 '----------------END--------------
ElseIf Button.Caption = "Close" Then
Unload Me
Unload estimatedincurredcost
ElseIf Button.Caption = "Disp Layout" Then
frm_layoutestr.Show
ElseIf Button.Caption = "App Layout" Then
Call flex_title1
Unload frm_layoutestr
ElseIf Button.Caption = "Excel" Then
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
    For n = 0 To 19
        flex_grid.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flex_grid.Text
    Next
Next
End If
End Sub


