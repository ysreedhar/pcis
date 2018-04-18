VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Project Cost Information System"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   2100
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   0
      ScaleHeight     =   7500
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.PictureBox Splitter 
         Height          =   8040
         Left            =   2865
         ScaleHeight     =   8040
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   -120
         Visible         =   0   'False
         Width           =   15
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   9540
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   16828
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   882
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlTree"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList51 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1376
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":17C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2116
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2430
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3308
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3622
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":393C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":18AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1ED48
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":24FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":252FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25456
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25770
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":25EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":261F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26648
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26962
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":26DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2720E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27528
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27842
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":27E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28190
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":285E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":28D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29068
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29382
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2969C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":29AEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   4155
      Top             =   3315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2A280
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2A6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2AF78
            Key             =   "employee"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B3CC
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B820
            Key             =   "open"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2BC74
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2C550
            Key             =   "report"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2C9A4
            Key             =   "shipper"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2D280
            Key             =   "group"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2D6D4
            Key             =   "supplier"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2DB28
            Key             =   "taxonomy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":302DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":305F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3090E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":30C28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   " TL   OFFSHORE  SDN  BHD"
            TextSave        =   " TL   OFFSHORE  SDN  BHD"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "03/05/2007"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:50 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   15435
         TabIndex        =   5
         Top             =   0
         Width           =   15500
         Begin MSComCtl2.DTPicker DTPcutdate1 
            Height          =   375
            Left            =   13080
            TabIndex        =   9
            Top             =   0
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy H:mm:ss"
            Format          =   48824323
            CurrentDate     =   38140
         End
         Begin MSComCtl2.DTPicker DTP_login 
            Height          =   300
            Left            =   5760
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            Format          =   48824321
            CurrentDate     =   38059
         End
         Begin VB.Label lbltitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   3720
            TabIndex        =   11
            Top             =   0
            Width           =   7935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cutt-Off Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   11760
            TabIndex        =   10
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   8640
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   11055
         End
      End
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   19800
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
            Picture         =   "main.frx":30F42
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":31054
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":314A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":318F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":31D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3219C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38436
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38750
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":38A6A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":39004
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3959E
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":39B38
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A0D2
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A1E4
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3A726
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3ACC0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3B25A
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BB34
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BC46
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BD58
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BE6A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3BF7C
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C08E
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C1A0
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3C73A
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3CCD4
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D26E
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D808
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3D91A
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3DA2C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3DFC6
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E0D8
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E1EA
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E784
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3E896
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3EE30
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3F3CA
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3F4DC
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3FA76
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40010
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":405AA
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":406BC
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40C56
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40D68
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40E7A
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":40F8C
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4109E
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":411B0
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4174A
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4185C
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4196E
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":41F08
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":424A2
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":42A3C
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":42FD6
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":43570
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":43B0A
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":440A4
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   4560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":441B6
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":442C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4471A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":44B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":44FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":45410
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4B6AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4B9C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4BCDE
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4C278
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4C812
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4CDAC
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D346
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D458
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4D99A
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4DF34
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4E4CE
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EDA8
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EEBA
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4EFCC
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F0DE
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F1F0
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F302
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F414
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F9AE
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4FF48
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":504E2
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50A7C
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50B8E
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":50CA0
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5123A
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5134C
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5145E
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":519F8
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":51B0A
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":520A4
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5263E
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":52750
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":52CEA
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53284
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5381E
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53930
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53ECA
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":53FDC
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":540EE
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54200
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54312
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54424
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":549BE
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54AD0
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":54BE2
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5517C
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":55716
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":55CB0
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5624A
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":567E4
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":56D7E
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":57318
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5742A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":613EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":72804
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82069
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":824BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82911
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":82D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":83214
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":836DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":83C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8410C
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":846C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":84B0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8502C
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":98F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":ACFB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B09BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B5173
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":B56B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":BCF80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnumaster 
      Caption         =   "&Masters"
      Begin VB.Menu mnu_projectcodes 
         Caption         =   "&Project Codes"
         Begin VB.Menu mnu_projectkey 
            Caption         =   "P&roject Key"
         End
         Begin VB.Menu mnu_jobno 
            Caption         =   "J&ob No"
         End
         Begin VB.Menu mnu_subjobno 
            Caption         =   "S&ub Job No"
         End
         Begin VB.Menu mnu_jobchargeno 
            Caption         =   "J&obcharge No"
         End
         Begin VB.Menu mnu_obscode 
            Caption         =   "O&BS Code"
         End
         Begin VB.Menu mnu_costcode 
            Caption         =   "C&ost Code"
         End
      End
      Begin VB.Menu mnu_resourcecodes 
         Caption         =   "&Resource Codes"
         Begin VB.Menu mnu_resourcetypecodes 
            Caption         =   "R&esource Type Codes"
         End
         Begin VB.Menu mnu_resourceresponsibilitycodes 
            Caption         =   "R&esource Responsibility Codes"
         End
         Begin VB.Menu mnu_resourcevendorcodes 
            Caption         =   "R&esource Vendor Codes"
         End
         Begin VB.Menu mnu_resourcecode 
            Caption         =   "R&esource Code"
         End
         Begin VB.Menu mnu_resourcemaptoprojectkey1 
            Caption         =   "R&esource Map To ProjectKey"
         End
      End
      Begin VB.Menu mnu_othercodes 
         Caption         =   "&Other Codes"
         Begin VB.Menu mnu_spreadcode 
            Caption         =   "S&pread Code"
         End
         Begin VB.Menu mnu_costtypecode 
            Caption         =   "C&ost Type Code"
         End
         Begin VB.Menu mnu_uom 
            Caption         =   "U&.O.M"
         End
         Begin VB.Menu mnu_currencycode 
            Caption         =   "C&urrency Code"
         End
         Begin VB.Menu mnu_exchangerate 
            Caption         =   "E&xchange Rate"
         End
         Begin VB.Menu mnu_ohpiitemcode 
            Caption         =   "O&H/PI Item Code"
         End
         Begin VB.Menu mnu_othertranxcoces 
            Caption         =   "O&ther TranX Codes"
         End
         Begin VB.Menu mnuL0NotesSignatures 
            Caption         =   "L0 Notes and Si&gnatures"
         End
      End
   End
   Begin VB.Menu mnu_Transactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnu_budgeteddetails 
         Caption         =   "&Budgeted Details"
         Begin VB.Menu mnu_budgeteddurationbyspread 
            Caption         =   "B&udgeted Duration By Spread"
         End
         Begin VB.Menu mnu_budgetedcostdetails 
            Caption         =   "B&udgeted Cost Details"
            Begin VB.Menu mnu_bcbyresource 
               Caption         =   "B&C By Resource"
            End
            Begin VB.Menu mnu_bcbyjobcharge 
               Caption         =   "B&C By Jobcharge"
            End
         End
      End
      Begin VB.Menu mnu_generateeicdetailsbybudget 
         Caption         =   "&Generate EIC Details From Budget"
         Begin VB.Menu mnu_generateeictransactions 
            Caption         =   "G&enerate EIC Transactions"
         End
         Begin VB.Menu mnu_editposttransactions 
            Caption         =   "E&dit/Post Transactions"
         End
      End
      Begin VB.Menu mnu_estimateddetails 
         Caption         =   "&Estimated Details"
         Begin VB.Menu mnu_estimatedprogressdurationbyspread 
            Caption         =   "E&stimated Progress Duration By Spread"
         End
         Begin VB.Menu mnu_estimatedincurredcostdetails 
            Caption         =   "E&stimated Incurred Cost Details"
            Begin VB.Menu mnu_eicbyresource 
               Caption         =   "E&IC By Resource"
            End
            Begin VB.Menu mnu_eicbyjobcharge 
               Caption         =   "E&IC By Jobcharge"
            End
         End
      End
      Begin VB.Menu mnu_otherdetails 
         Caption         =   "&Other Details"
         Begin VB.Menu mnu_revenuebdgtvoadjbilledunbilled 
            Caption         =   "R&evenue-Bdgt/VO/Adj/Billed/Unbilled"
         End
         Begin VB.Menu mnu_otherincexpoverheadestrecovery 
            Caption         =   "O&ther Inc/Exp & Overhead  - Est/Recovery"
         End
         Begin VB.Menu mnu_variationorderunrealized 
            Caption         =   "V&ariation Order - Unrealized/Potential"
         End
         Begin VB.Menu mnu_billedcost 
            Caption         =   "B&illed Cost"
         End
         Begin VB.Menu mnu_projectdiary 
            Caption         =   "Project Diary"
         End
         Begin VB.Menu mnu_bpbdgt 
            Caption         =   "BusinessPlan Budget"
         End
      End
      Begin VB.Menu mnu_quickupdates 
         Caption         =   "&Quick Updates"
         Begin VB.Menu mnu_updateworkcomplete 
            Caption         =   "U&pdate % Work Complete"
         End
         Begin VB.Menu mnu_updateunitrate 
            Caption         =   "U&pdate Unit Rate"
            Begin VB.Menu mnu_bctransactions 
               Caption         =   "B&C Transactions"
            End
            Begin VB.Menu mnu_eictransactions 
               Caption         =   "E&IC Transactions"
            End
            Begin VB.Menu mnu_eicbyresc 
               Caption         =   "E&IC By Resource"
            End
         End
         Begin VB.Menu mnu_updatedatesfornaeic 
            Caption         =   "U&pdate Dates For NA-EIC"
         End
         Begin VB.Menu mnu_updateqty 
            Caption         =   "Update Qty"
         End
         Begin VB.Menu mnu_updatejcb 
            Caption         =   "Update Jobcharge/BC"
         End
         Begin VB.Menu mnu_updatejobcharge 
            Caption         =   "Update Jobcharge/EIC"
         End
      End
      Begin VB.Menu mnu_periodendupdates 
         Caption         =   "&Period End Updates"
         Begin VB.Menu mnu_revenueprojectkeylevel 
            Caption         =   "R&evenue @ProjectKey Level"
         End
         Begin VB.Menu mnu_costjoblevel 
            Caption         =   "C&ost @ Job Level"
         End
      End
   End
   Begin VB.Menu mnu_reports 
      Caption         =   "&Reports"
      Begin VB.Menu mnu_masterlists 
         Caption         =   "&Master Lists"
         Begin VB.Menu mnu_projectrep 
            Caption         =   "P&roject"
            Begin VB.Menu mnu_projectlistrep 
               Caption         =   "Pr&oject List"
            End
            Begin VB.Menu mnu_jobnolistrep 
               Caption         =   "Jo&b No List"
            End
            Begin VB.Menu mnu_subjobnolistrep 
               Caption         =   "Su&b Job No List"
            End
            Begin VB.Menu mnu_jobchargenolistrep 
               Caption         =   "Jo&bcharge No List"
            End
            Begin VB.Menu mnu_obscoderep 
               Caption         =   "OB&S Code List"
            End
            Begin VB.Menu mnu_costcodelistrep 
               Caption         =   "Co&st Code List"
            End
         End
         Begin VB.Menu mnu_resourcerep 
            Caption         =   "R&esource"
            Begin VB.Menu mnu_resourcetypecodelistrep 
               Caption         =   "Re&source Type Code List"
            End
            Begin VB.Menu mnu_resourceresponsibiltycodelistrep 
               Caption         =   "Re&source Responsibility Code List"
            End
            Begin VB.Menu mnu_resourcevendorcodelistrep 
               Caption         =   "Re&source Vendor Code List"
            End
            Begin VB.Menu mnu_resourcecodelistrep 
               Caption         =   "Re&source Code List"
            End
         End
         Begin VB.Menu mnu_othesrep 
            Caption         =   "O&thers"
            Begin VB.Menu mnu_spreadcodelistrep 
               Caption         =   "Sp&read Code List"
            End
            Begin VB.Menu mnu_costtypecodelistrep 
               Caption         =   "Co&st type Code List"
            End
            Begin VB.Menu mnu_uomrep 
               Caption         =   "U.O.&M"
            End
            Begin VB.Menu mnu_currencycodelistrep 
               Caption         =   "Cu&rrency Code List"
            End
            Begin VB.Menu mnu_exchangeratelistrep 
               Caption         =   "Ex&change Rate List"
            End
            Begin VB.Menu mnu_tranxidforoverheadpitemslistrep 
               Caption         =   "Tr&anX ID For OverHead & P/Items List"
            End
         End
      End
      Begin VB.Menu mnu_budgetedreportsrep 
         Caption         =   "&Budget Reports"
         Begin VB.Menu mnu_budgetedduartionbyspreadrep 
            Caption         =   "B&udgeted Duration By Spread"
         End
         Begin VB.Menu mnu_bdjc 
            Caption         =   "B&udgeted Duration By JobCharge"
         End
         Begin VB.Menu mnu_budgetedcostdetailsrep1 
            Caption         =   "B&udgeted Cost Details"
            Begin VB.Menu mnu_bcbyresourcerep 
               Caption         =   "BC& By Resource"
            End
            Begin VB.Menu mnu_bcbyresourcecostcoderep 
               Caption         =   "BC& By Resource/CostCode"
            End
            Begin VB.Menu mnu_bcbyjobchargerep 
               Caption         =   "BC& By Jobcharge"
            End
            Begin VB.Menu mnu_bcbyobsrep 
               Caption         =   "BC& By OBS"
            End
         End
      End
      Begin VB.Menu mnu_estimatedincurredreportsrep 
         Caption         =   "&Estimated Incurred Reports"
         Begin VB.Menu mnu_estimatedprogressdurationbyspreadrep 
            Caption         =   "E&stimated Progress Duration By Spread"
         End
         Begin VB.Menu mnu_eicjc 
            Caption         =   "E&stimated Progress Duration By JobCharge"
         End
         Begin VB.Menu mnu_dvar 
            Caption         =   "D&uration Vairance By Project"
         End
         Begin VB.Menu mnu_estimatedincurredcostdetailsrep 
            Caption         =   "E&stimated Incurred Cost Details"
            Begin VB.Menu mnu_estimatedincurredcostbyresourcerep 
               Caption         =   "EI&C By Resource"
            End
            Begin VB.Menu mnu_estimatedincurredcostbyjobchargerep 
               Caption         =   "EI&C By JobCharge"
            End
            Begin VB.Menu mnu_estimatedincurredcostbyobsrep 
               Caption         =   "EI&C By OBS"
            End
         End
      End
      Begin VB.Menu mnu_managementreports 
         Caption         =   "&Management Reports"
         Begin VB.Menu mnu_l0rep 
            Caption         =   "L&0 - PRCR @ Company Level - All Projects"
         End
         Begin VB.Menu mnu_l1rep 
            Caption         =   "L&1 - PRCR @ Project key Level - All Projects"
         End
         Begin VB.Menu mnu_l2rep 
            Caption         =   "L&2 - PRCR @ JobCharge Level - By Project Key"
         End
         Begin VB.Menu mnu_l3rep 
            Caption         =   "L&3 - PRCR @ Details Level - By ProjectKey & Job"
         End
      End
      Begin VB.Menu mnu_miscelleneousrep 
         Caption         =   "&Miscelleneous Reports"
         Begin VB.Menu mnu_revenuedetailsrep 
            Caption         =   "R&evenue Details"
         End
         Begin VB.Menu mnu_budgetedrevenuevariationorderrep 
            Caption         =   "B&udgeted Revenue/Variation Order"
         End
         Begin VB.Menu mnu_revenuebilledunbilledrep 
            Caption         =   "R&evenue Billed/Unbilled"
         End
         Begin VB.Menu mnu_costaccruallistrep 
            Caption         =   "S&pread Cost Summary"
         End
         Begin VB.Menu mnu_costsummarybyresourcerep 
            Caption         =   "C&ost Summary By Resource"
         End
         Begin VB.Menu mnu_estimatebilledcostrep 
            Caption         =   "E&stimate Vs Billed Cost"
         End
         Begin VB.Menu mnu_tablesrep 
            Caption         =   "&Tables"
            Begin VB.Menu mnu_tablesbcrep 
               Caption         =   "BC&  Details"
            End
            Begin VB.Menu mnu_tableseicrep 
               Caption         =   "EI&C Details"
            End
         End
      End
   End
   Begin VB.Menu mnu_utilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnu_BackUp 
         Caption         =   "B&ackUp"
      End
      Begin VB.Menu mnu_restore 
         Caption         =   "R&estore"
      End
      Begin VB.Menu mnu_sendmessage 
         Caption         =   "Send Message"
      End
   End
   Begin VB.Menu mnu_administration 
      Caption         =   "&Administration"
      Begin VB.Menu mnu_companyparameter 
         Caption         =   "C&ompany Parameter"
      End
      Begin VB.Menu mnu_createPassword 
         Caption         =   "C&reate Password"
      End
      Begin VB.Menu mnu_userrights 
         Caption         =   "U&ser Rights"
      End
      Begin VB.Menu mnu_rulesvalidations 
         Caption         =   "R&ules & Validations"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_dataflow 
         Caption         =   "D&ata Flow"
      End
      Begin VB.Menu mnu_formhelp 
         Caption         =   "F&orm Help"
      End
      Begin VB.Menu mnu_logout 
         Caption         =   "L&ogout"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim f As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_THICKFRAME = &H40000
Private Const WS_EX_STATICEDGE = &H20000
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Dim Flag As Boolean
Private Sub DTPcutdate1_Change()
'frmBusy.Show
On Error Resume Next
'oitran.Show
If fab = 1 Then
Dim ppr1 As Double
Dim ppsd1 As Date
Dim pped1 As Date
Dim dys1 As Double
dys1 = 0
ppr1 = 0
Dim prf As New ADODB.Recordset
If prf.State Then prf.Close
prf.Open "select * from parameters", Cn, 3, 2
If Not prf.EOF Then
ppr1 = prf!p_ydays
ppsd1 = prf!p_sdate
pped1 = prf!p_edate
'ppcd = pr!p_cdate
End If
ppcd1 = DTPcutdate1.Value
If (ppcd1 - ppsd1) < 0 Then
oitran.txt_bcwpdays.Text = 0
Else
oitran.txt_bcwpdays.Text = Round(CDbl(ppcd1 - ppsd1), 2)
End If
If (pped1 - ppcd1) < 0 Then
oitran.txt_etcdays.Text = 0
Else
oitran.txt_etcdays.Text = Round(CDbl(pped1 - ppcd1), 2)
End If

 dys1 = 0
dys1 = DTPcutdate1.Value - oitran.dtp_asat.Value
 
oitran.txt_bcwp.Text = Round(CDbl(oitran.txt_bcwpbl.Text) * CDbl(oitran.txt_bcwpdays.Text), 2)
oitran.txt_etcbl.Text = Format(Round(CDbl(oitran.txt_bdgt.Text) / ppr, 2), "###,###,##0.00")
oitran.txt_acwpadj.Text = dys1
oitran.txt_acwp.Text = Format(Round(CDbl(oitran.txt_acwpacc.Text) + CDbl(oitran.txt_acwpbl.Text * oitran.txt_acwpadj.Text), 2), "###,###,##0")
txt_acwpbl.Text = txt_rateaft.Text
End If
'Call progcost
'Unload frmBusy
End Sub

Private Sub DTPcutdate1_Click()
On Error Resume Next
If fab = 1 Then
Dim ppr1 As Double
Dim ppsd1 As Date
Dim pped1 As Date
Dim dys1 As Double
dys1 = 0
ppr1 = 0
Dim prf As New ADODB.Recordset
If prf.State Then prf.Close
prf.Open "select * from parameters", Cn, 3, 2
If Not prf.EOF Then
ppr1 = prf!p_ydays
ppsd1 = prf!p_sdate
pped1 = prf!p_edate
'ppcd = pr!p_cdate
End If
ppcd1 = DTPcutdate1.Value
If (ppcd1 - ppsd1) < 0 Then
oitran.txt_bcwpdays.Text = 0
Else
oitran.txt_bcwpdays.Text = Round(CDbl(ppcd1 - ppsd1), 2)
End If
If (pped1 - ppcd1) < 0 Then
oitran.txt_etcdays.Text = 0
Else
oitran.txt_etcdays.Text = Round(CDbl(pped1 - ppcd1), 2)
End If
 dys1 = 0
dys1 = DTPcutdate1.Value - oitran.dtp_asat.Value
oitran.txt_bcwp.Text = Round(CDbl(oitran.txt_bcwpbl.Text) * CDbl(oitran.txt_bcwpdays.Text), 2)
oitran.txt_etcbl.Text = Round(CDbl(oitran.txt_bdgt.Text) / ppr1, 2)
oitran.txt_acwpbl.Text = Round((CDbl(oitran.txt_bdgt.Text) / ppr1) * dys1, 2)
oitran.txt_acwp.Text = Round(CDbl(oitran.txt_acwpacc.Text) + CDbl(oitran.txt_acwpbl.Text) + CDbl(oitran.txt_acwpadj.Text), 2)
oitran.Label6.Caption = "(B/L Bdgt)/" & ppr1
oitran.Label3.Caption = "((B/L)/" & ppr1 & ") * " & dys1
oitran.Label15.Caption = "(B/L Bdgt)/" & ppr1
oitran.Label3.Caption = "((B/L)/" & ppr1 & " * " & dys1
End If
'Call progcost
End Sub
Private Sub MDIForm_Load()
On Error Resume Next
Option1.Value = True
Call connect
fab = 0
Dim rst As New ADODB.Recordset
If rst.State Then rst.Close
rst.Open "select * from userrights where u_name='" & frm_login.cbo_userid.Text & "' ", Cn, 3, 2
If Not rst.EOF Then
a = rst!mforms
b = rst!tforms
c = rst!mreports
d = rst!treports
f = rst!others
End If
 Call tree
 'Call userinvisible
 Call mnuuser
LoadDragEvents
StatusBar1.Panels(1).Text = GetCompanyName
End Sub

Private Sub picLeftPane_Resize()
If Flag = False Then Exit Sub
'Change the Width of the Treeview and reset the flag
TreeView1.Width = picLeftPane.Width - 10
Flag = True
End Sub

Function LoadDragEvents()
SetWindowLong picLeftPane.hwnd, GWL_STYLE, GetWindowLong(picLeftPane.hwnd, GWL_STYLE) Or WS_THICKFRAME
SetWindowLong picLeftPane.hwnd, GWL_EXSTYLE, GetWindowLong(picLeftPane.hwnd, GWL_EXSTYLE) Or WS_EX_STATICEDGE
SetWindowPos picLeftPane.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
' set the flag
Flag = True
End Function
Private Function tree()

TreeView1.Nodes.Add , , "l", "PROJECT COST INFORMATION SYSTEM", 14

'Company master
If Mid(a, 1, 1) = 1 Then
TreeView1.Nodes.Add "l", tvwChild, "CMPmaster", UCase("MASTERS"), 5, 6
If Mid(a, 2, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "pcodes", ("Project Codes"), 5, 6
If Mid(a, 3, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "projectmaster", ("Project Key"), 5, 6
End If
If Mid(a, 4, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "Jobno", ("Job No"), 5, 6
End If
If Mid(a, 5, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "SubJobno", ("Sub Job No"), 5, 6
End If
If Mid(a, 6, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "JobChargemaster", ("Job Charge No"), 5, 6
End If
If Mid(a, 7, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "respd", ("OBS Code"), 5, 6
End If
If Mid(a, 8, 1) = 1 Then
TreeView1.Nodes.Add "pcodes", tvwChild, "cc", ("CostCode"), 5, 6
End If
End If

If Mid(a, 9, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "rcodes", ("Resource Codes"), 5, 6
If Mid(a, 10, 1) = 1 Then
TreeView1.Nodes.Add "rcodes", tvwChild, "rty", ("Resource Type Codes"), 5, 6
End If
If Mid(a, 11, 1) = 1 Then
TreeView1.Nodes.Add "rcodes", tvwChild, "resp", ("Resource Responsiblility Code"), 5, 6
End If
If Mid(a, 12, 1) = 1 Then
TreeView1.Nodes.Add "rcodes", tvwChild, "vendor", ("Resource Vendor Code"), 5, 6
End If
If Mid(a, 13, 1) = 1 Then
TreeView1.Nodes.Add "rcodes", tvwChild, "Resourcemaster", ("Resource Code"), 5, 6
End If
If Mid(a, 14, 1) = 1 Then
TreeView1.Nodes.Add "rcodes", tvwChild, "rp", ("Resource Map To ProjectKey"), 5, 6
End If
End If
If Mid(a, 15, 1) = 1 Then
TreeView1.Nodes.Add "CMPmaster", tvwChild, "scodes", ("Other Codes"), 5, 6
If Mid(a, 16, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "Spread", ("Spread Code"), 5, 6
End If
If Mid(a, 17, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "Tranxtype", ("Cost Type Code"), 5, 6
End If
If Mid(a, 18, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "UOMmaster", ("U.O.M"), 5, 6
End If
If Mid(a, 19, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "Currencymaster", ("Currency Code"), 5, 6
End If
If Mid(a, 20, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "Currencyexchange", ("Exchange Rate"), 5, 6
End If
If Mid(a, 21, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "ohpi", ("OH / PI Item Code"), 5, 6
End If
If Mid(a, 22, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "otranx", ("Other TranX Codes"), 5, 6
End If
If Mid(a, 22, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "Chargetype", ("Charge Type Code"), 5, 6
End If
If Mid(a, 22, 1) = 1 Then
TreeView1.Nodes.Add "scodes", tvwChild, "L0Notes", ("L0 Notes and Signatures"), 5, 6
End If
End If
End If 'masters

' Transaction
If Mid(b, 1, 1) = 1 Then
TreeView1.Nodes.Add "l", tvwChild, "Tranx", UCase("TRANSACTIONS"), 5, 6
If Mid(b, 2, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "bddetails", ("Budget Details"), 5, 6
If Mid(b, 3, 1) = 1 Then
TreeView1.Nodes.Add "bddetails", tvwChild, "DurationSchedule", ("Budgeted Duration By Spread"), 5, 6
End If
If Mid(b, 4, 1) = 1 Then
TreeView1.Nodes.Add "bddetails", tvwChild, "bdcost", ("Budgeted Cost Details"), 5, 6
If Mid(b, 5, 1) = 1 Then
TreeView1.Nodes.Add "bdcost", tvwChild, "budget1", ("BC By Resource"), 5, 6
End If
If Mid(b, 6, 1) = 1 Then
TreeView1.Nodes.Add "bdcost", tvwChild, "budget2", ("BC By JobCharge"), 5, 6
TreeView1.Nodes.Add "bdcost", tvwChild, "BudgetImport", ("Budget by Resource Import"), 5, 6
TreeView1.Nodes.Add "BudgetImport", tvwChild, "ImportExistingJobCharge_B", ("Duplicate from Existing JobCharge"), 5, 6
End If
End If
End If
If Mid(b, 7, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "EIC", ("Generate EIC Details From Budget"), 5, 6
If Mid(b, 8, 1) = 1 Then
TreeView1.Nodes.Add "EIC", tvwChild, "budtran", ("Generate EIC Transactions"), 5, 6
End If
If Mid(b, 9, 1) = 1 Then
TreeView1.Nodes.Add "EIC", tvwChild, "estpost", ("Edit/Post Transactions"), 5, 6
End If
End If
If Mid(b, 10, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "estdetails", ("Estimated Details"), 5, 6
If Mid(b, 11, 1) = 1 Then
TreeView1.Nodes.Add "estdetails", tvwChild, "progressschedule", ("Estimated Progress Duration By Spread"), 5, 6
End If
If Mid(b, 12, 1) = 1 Then
TreeView1.Nodes.Add "estdetails", tvwChild, "estcost", ("Estimated Incurred Cost Details"), 5, 6
    If Mid(b, 13, 1) = 1 Then
    TreeView1.Nodes.Add "estcost", tvwChild, "incurredProjectbyResource", ("EIC By Resource"), 5, 6
    TreeView1.Nodes.Add "estcost", tvwChild, "incurred1", ("EIC By Project By Resource"), 5, 6
    
    End If
    If Mid(b, 14, 1) = 1 Then
    TreeView1.Nodes.Add "estcost", tvwChild, "incurred2", ("EIC Project By JobCharge"), 5, 6
    TreeView1.Nodes.Add "estcost", tvwChild, "estCostImport", ("EIC by Resource Import"), 5, 6
    TreeView1.Nodes.Add "estCostImport", tvwChild, "ImportExistingJobCharge_E", ("Duplicate from Existing JobCharge"), 5, 6
    TreeView1.Nodes.Add "estCostImport", tvwChild, "ImportUsingExcelSheet", ("Import from Excel Worksheet for EIC/Budget"), 5, 6
    End If
End If
End If
If Mid(b, 15, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "revtranx", ("Other Details"), 5, 6
If Mid(b, 16, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "billunbillrev", ("Revenue-Bdgt/VO/Adj/Billed/Unbilled"), 5, 6
End If
If Mid(b, 17, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "oitran", ("Other Inc/Exp & Overhead - Est/Recovery"), 5, 6
End If
If Mid(b, 18, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "pi", ("Variation Order - Unrealized/Potential"), 5, 6
End If
If Mid(b, 19, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "billed", ("Billed Cost"), 5, 6
End If
If Mid(b, 20, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "pdiary", ("Project Diary"), 5, 6
End If
If Mid(b, 21, 1) = 1 Then
TreeView1.Nodes.Add "revtranx", tvwChild, "bpbdgt", ("BusinessPlan Budget"), 5, 6
End If
End If
If Mid(b, 22, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "qupdates", ("Quick updates"), 5, 6
    If Mid(b, 23, 1) = 1 Then
    TreeView1.Nodes.Add "qupdates", tvwChild, "perwork", ("Update % Work Complete"), 5, 6
    End If
        If Mid(b, 24, 1) = 1 Then
        TreeView1.Nodes.Add "qupdates", tvwChild, "urate", ("Update Unit Rate"), 5, 6
            If Mid(b, 25, 1) = 1 Then
            TreeView1.Nodes.Add "urate", tvwChild, "urateb", ("BC Transactions"), 5, 6
            End If
            If Mid(b, 26, 1) = 1 Then
            TreeView1.Nodes.Add "urate", tvwChild, "uratee", ("EIC Transactions"), 5, 6
            End If
            If Mid(b, 27, 1) = 1 Then
            TreeView1.Nodes.Add "urate", tvwChild, "qupdr", ("EIC Transactions By Resc"), 5, 6
            End If
        End If
    If Mid(b, 28, 1) = 1 Then
    TreeView1.Nodes.Add "qupdates", tvwChild, "una", ("Update Dates for NA-EIC"), 5, 6
    End If
    If Mid(b, 29, 1) = 1 Then
    TreeView1.Nodes.Add "qupdates", tvwChild, "uqty", ("Update QTY"), 5, 6
    End If
    If Mid(b, 30, 1) = 1 Then
    TreeView1.Nodes.Add "qupdates", tvwChild, "ujcb", ("Update JobCharge/BC"), 5, 6
    End If
    If Mid(b, 31, 1) = 1 Then
    TreeView1.Nodes.Add "qupdates", tvwChild, "ujc", ("Update JobCharge/EIC"), 5, 6
    End If
End If
If Mid(b, 31, 1) = 1 Then
TreeView1.Nodes.Add "Tranx", tvwChild, "tTranx", ("Period End Updates"), 5, 6
If Mid(b, 32, 1) = 1 Then
TreeView1.Nodes.Add "tTranx", tvwChild, "ptranx", ("Revenue @ Projectkey Level"), 5, 6
End If
If Mid(b, 33, 1) = 1 Then
TreeView1.Nodes.Add "tTranx", tvwChild, "ctranx", ("Cost @ Job Level"), 5, 6
End If
End If
End If 'transactions
' Reports
If Mid(c, 1, 1) = 1 Then
TreeView1.Nodes.Add "l", tvwChild, "Rep", UCase("REPORTS"), 5, 6
If Mid(c, 2, 1) = 1 Then
TreeView1.Nodes.Add "Rep", tvwChild, "Master", ("Master Lists"), 5, 6
    If Mid(c, 3, 1) = 1 Then
        TreeView1.Nodes.Add "Master", tvwChild, "pj", ("Project"), 5, 6
        If Mid(c, 4, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "pl", ("Project List"), 5, 6
        End If
        If Mid(c, 5, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "jn", ("Job No List"), 5, 6
        End If
        If Mid(c, 6, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "sjn", ("Sub Job No List"), 5, 6
        End If
        If Mid(c, 7, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "jcl", ("Job Charge No List"), 5, 6
        End If
        If Mid(c, 8, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "obscd", ("OBS Code List"), 5, 6
        End If
        If Mid(c, 9, 1) = 1 Then
        TreeView1.Nodes.Add "pj", tvwChild, "ccl", ("Cost Code List"), 5, 6
        End If
    End If
If Mid(c, 10, 1) = 1 Then
        TreeView1.Nodes.Add "Master", tvwChild, "rs", ("Resource"), 5, 6
        If Mid(c, 11, 1) = 1 Then
        TreeView1.Nodes.Add "rs", tvwChild, "rtcl", ("Resource Type Code List"), 5, 6
        End If
        If Mid(c, 12, 1) = 1 Then
        TreeView1.Nodes.Add "rs", tvwChild, "rrcl", ("Resource Responsibility Code List"), 5, 6
        End If
        If Mid(c, 13, 1) = 1 Then
        TreeView1.Nodes.Add "rs", tvwChild, "rvcl", ("Resource Vendor Code List"), 5, 6
        End If
        If Mid(c, 14, 1) = 1 Then
        TreeView1.Nodes.Add "rs", tvwChild, "rl", ("Resource Code List"), 5, 6
        End If
End If
    If Mid(c, 15, 1) = 1 Then
    TreeView1.Nodes.Add "Master", tvwChild, "ott", ("Others"), 5, 6
            If Mid(c, 16, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "sl", ("Spread Code List"), 5, 6
            End If
            If Mid(c, 17, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "ctcl", ("Cost type Code List"), 5, 6
            End If
            If Mid(c, 18, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "ul", ("U.O.M List"), 5, 6
            End If
            If Mid(c, 19, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "cl", ("Currency Code List"), 5, 6
            End If
            If Mid(c, 20, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "erl", ("Exchange Rate List"), 5, 6
            End If
            If Mid(c, 21, 1) = 1 Then
            TreeView1.Nodes.Add "ott", tvwChild, "ttl", ("TranX ID For Overhead & P/Items List"), 5, 6
            End If
    End If
End If 'master
End If
If Mid(d, 1, 1) = 1 Then
TreeView1.Nodes.Add "Rep", tvwChild, "pcpr", ("Budget Reports"), 5, 6
    If Mid(d, 2, 1) = 1 Then
    TreeView1.Nodes.Add "pcpr", tvwChild, "bdur", ("Budgeted Duration By Spread"), 5, 6
    End If
        If Mid(d, 3, 1) = 1 Then
    TreeView1.Nodes.Add "pcpr", tvwChild, "bdurp", ("Budgeted Duration By Project"), 5, 6
    End If
        If Mid(d, 4, 1) = 1 Then
            TreeView1.Nodes.Add "pcpr", tvwChild, "bcost", ("Budgeted Cost Details"), 5, 6
            If Mid(d, 5, 1) = 1 Then
            TreeView1.Nodes.Add "bcost", tvwChild, "bdgt", ("BC By Resource"), 5, 6
            End If
            If Mid(d, 6, 1) = 1 Then
            TreeView1.Nodes.Add "bcost", tvwChild, "bdgtcost1", ("BC By Resource/Costcode"), 5, 6
            End If
            If Mid(d, 7, 1) = 1 Then
            TreeView1.Nodes.Add "bcost", tvwChild, "bdgtj", ("BC By Jobcharge"), 5, 6
            End If
            If Mid(d, 8, 1) = 1 Then
            TreeView1.Nodes.Add "bcost", tvwChild, "obsb", ("BC By OBS"), 5, 6
            End If
        End If
   End If
        If Mid(d, 9, 1) = 1 Then
        TreeView1.Nodes.Add "Rep", tvwChild, "eir", ("Estimated Incurred Reports"), 5, 6
        If Mid(d, 10, 1) = 1 Then
        TreeView1.Nodes.Add "eir", tvwChild, "pdur", ("Estimated Progress Duration By Spread"), 5, 6
        End If

                If Mid(d, 11, 1) = 1 Then
                TreeView1.Nodes.Add "eir", tvwChild, "pdurp", ("Estimated Progress Duration By Project"), 5, 6
                End If
                
                If Mid(d, 12, 1) = 1 Then
                TreeView1.Nodes.Add "eir", tvwChild, "durvar", ("Duration Variance By Project"), 5, 6
                End If
            If Mid(d, 13, 1) = 1 Then
            TreeView1.Nodes.Add "eir", tvwChild, "eicd", ("Estimated Incurred Cost Details"), 5, 6
            If Mid(d, 14, 1) = 1 Then
            TreeView1.Nodes.Add "eicd", tvwChild, "estres", ("EIC By Resource"), 5, 6
            TreeView1.Nodes.Add "eicd", tvwChild, "estbyProjbyres", ("EIC BY Project By Resource"), 5, 6
            End If
            If Mid(d, 15, 1) = 1 Then
            TreeView1.Nodes.Add "eicd", tvwChild, "estjob", ("EIC By Jobcharge"), 5, 6
            End If
            If Mid(d, 16, 1) = 1 Then
            TreeView1.Nodes.Add "eicd", tvwChild, "obse", ("EIC By OBS"), 5, 6
            End If
            End If
        
        End If
            If Mid(d, 17, 1) = 1 Then
            TreeView1.Nodes.Add "Rep", tvwChild, "mgmt", ("Management Reports"), 5, 6
            If Mid(d, 18, 1) = 1 Then
            TreeView1.Nodes.Add "mgmt", tvwChild, "lo", ("L0 - PRCR @ Company Level - All Projects"), 5, 6
            End If
            If Mid(d, 19, 1) = 1 Then
            TreeView1.Nodes.Add "mgmt", tvwChild, "psr", ("L1 - PRCR @ Project Key Level - All Projects"), 5, 6
            End If
            If Mid(d, 20, 1) = 1 Then
            TreeView1.Nodes.Add "mgmt", tvwChild, "csr", ("L2 - PRCR @ JobCharge Level - By Project Key"), 5, 6
            End If
            If Mid(d, 21, 1) = 1 Then
            TreeView1.Nodes.Add "mgmt", tvwChild, "cdr", ("L3 - PRCR @ Details Level - By Project Key & Job"), 5, 6
            End If
            End If
            
If Mid(d, 22, 1) = 1 Then
TreeView1.Nodes.Add "Rep", tvwChild, "misc", ("Miscelleneous Reports"), 5, 6
If Mid(d, 23, 1) = 1 Then
TreeView1.Nodes.Add "misc", tvwChild, "rdte", ("Revenue Details"), 5, 6
    If Mid(d, 24, 1) = 1 Then
    TreeView1.Nodes.Add "rdte", tvwChild, "ral", ("Budgeted Revenue / Variation Order"), 5, 6
    End If
    If Mid(d, 25, 1) = 1 Then
    TreeView1.Nodes.Add "rdte", tvwChild, "rbu", ("Revenue Billed/Unbilled"), 5, 6
    End If
End If
If Mid(d, 26, 1) = 1 Then
TreeView1.Nodes.Add "misc", tvwChild, "cal", ("Spread Cost Summary"), 5, 6
End If
If Mid(d, 27, 1) = 1 Then
TreeView1.Nodes.Add "misc", tvwChild, "csrr", ("EIC By Resource/Project"), 5, 6
End If
If Mid(d, 28, 1) = 1 Then
TreeView1.Nodes.Add "misc", tvwChild, "ebaResc", ("Estimate Vs Billed Cost By Resource"), 5, 6
TreeView1.Nodes.Add "misc", tvwChild, "ebaJC", ("Estimate Vs Billed Cost By JobCharge"), 5, 6
TreeView1.Nodes.Add "misc", tvwChild, "VISUMM", ("Vendor Invoice Summary"), 5, 6

End If
If Mid(d, 29, 1) = 1 Then
TreeView1.Nodes.Add "misc", tvwChild, "tbls", ("Tables"), 5, 6
If Mid(d, 30, 1) = 1 Then
TreeView1.Nodes.Add "tbls", tvwChild, "buddet", ("Budget Cost Details"), 5, 6
End If
If Mid(d, 31, 1) = 1 Then
TreeView1.Nodes.Add "tbls", tvwChild, "estdet", ("EIC Cost Details"), 5, 6
End If
End If
End If
If Mid(f, 1, 1) = 1 Then
TreeView1.Nodes.Add "l", tvwChild, "othe", UCase("Others"), 5, 6
' utilities
If Mid(f, 2, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "util", UCase("Utilities"), 5, 6
If Mid(f, 3, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "backup1", ("Backup"), 5, 6
End If
If Mid(f, 4, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "Restore1", ("Restore"), 5, 6
End If
If Mid(f, 5, 1) = 1 Then
TreeView1.Nodes.Add "util", tvwChild, "msg", ("Send Message"), 5, 6
End If
End If
' Administration
If Mid(f, 6, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "admin", UCase("Administration"), 5, 6
If Mid(f, 7, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "cmppara", ("Company Parameter"), 5, 6
End If
If Mid(f, 8, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "Chngpswd", ("Create Password"), 5, 6
End If
If Mid(f, 9, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "User", ("User Rights"), 5, 6
End If
If Mid(f, 10, 1) = 1 Then
TreeView1.Nodes.Add "admin", tvwChild, "ibdgt", ("Import Budget"), 5, 6
End If
End If
' Help
If Mid(f, 11, 1) = 1 Then
TreeView1.Nodes.Add "othe", tvwChild, "hlp", UCase("Help"), 5, 6
If Mid(f, 12, 1) = 1 Then
TreeView1.Nodes.Add "hlp", tvwChild, "Dataflow", ("Data Flow"), 5, 6
End If
If Mid(f, 13, 1) = 1 Then
TreeView1.Nodes.Add "hlp", tvwChild, "frmhelp", ("Form Help"), 5, 6
End If
End If
End If
TreeView1.Nodes.Add "l", tvwChild, "logo", UCase("logout"), 5, 6
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Dim rss As New ADODB.Recordset
If rss.State Then rss.Close
rss.Open "select * from login where l_userid='" & Label2.Caption & "' and (l_intime)='" & DTP_login.Value & "'", Cn, 5, 6
If Not rss.EOF Then
rss!l_outtime = Now
rss.Update
End If
End Sub
Private Sub mnu_bdgt_dur_Click()
frm_budgetedduration.Show
End Sub
Private Sub mnu_budg_Click()
frm_costtranx.Show
End Sub
Private Sub mnu_currency_Click()
frm_currencymaster.Show
End Sub
Private Sub mnu_jobcharge_Click()
frm_jobcharge.Show
End Sub
Private Sub mnu_new_Click()
'If frm_jobcharge.Show = True Then
'frm_j
'Else
'MsgBox "Sorry! Form Not Selected"
'End If
End Sub
Private Sub mnu_prgs_Click()
frm_progressdurationdetails.Show
End Sub
Private Sub mnu_projmast_Click()
frm_projectmaster.Show
End Sub
Private Sub mnu_rescmaster_Click()
frm_resourcemaster.Show
End Sub
Private Sub mnu_spread_Click()
frm_Spreadmaster.Show
End Sub
Private Sub mnu_tranx_Click()
frm_transactiontype.Show
End Sub
Private Sub mnu_estimatedincurredcostbyresource_Click()
End Sub
Private Sub mnu_estimatedincurredcostbyobs_Click()
End Sub
Private Sub mnu_budgetedcostdetailsrep_Click()
End Sub
Private Sub mnu_BackUp_Click()
Cn.Execute "BACKUP DATABASE PCMS to disk='c:\PCMS" & Format(Date, "dd-MMM-yyyy") & ".bak'"
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End Sub
Private Sub mnu_bcbyjobcharge_Click()
frm_budgetedcost.Show
End Sub
Private Sub mnu_bcbyjobchargerep_Click()
rpt_budgetbyjobcharge.Show
End Sub

Private Sub mnu_bcbyobsrep_Click()
rpt_obsbudget.Show
End Sub

Private Sub mnu_bcbyresource_Click()
frm_costtranx.Show
End Sub

Private Sub mnu_bcbyresourcecostcoderep_Click()
rpt_budgetbycost.Show
End Sub

Private Sub mnu_bcbyresourcerep_Click()
rpt_budgetbyresource.Show
End Sub

Private Sub mnu_bctransactions_Click()
frm_updateunitrate.Show
End Sub

Private Sub mnu_bdjc_Click()
rpt_budgeteddurationjb.Show
End Sub

Private Sub mnu_billedcost_Click()
frm_billedcost.Show
End Sub

Private Sub mnu_bpbdgt_Click()
frm_baseline.Show
End Sub

Private Sub mnu_budgetedduartionbyspreadrep_Click()
rpt_budgetedduration.Show
End Sub

Private Sub mnu_budgeteddurationbyspread_Click()
frm_budgetedduration.Show
End Sub

Private Sub mnu_budgetedrevenuevariationorderrep_Click()
rpt_revenuebu.Show
End Sub

Private Sub mnu_companyparameter_Click()
frm_parameters.Show
End Sub

Private Sub mnu_costaccruallistrep_Click()
rpt_spreadcostsummary.Show
End Sub

Private Sub mnu_costcode_Click()
frm_costcode.Show
End Sub

Private Sub mnu_costcodelistrep_Click()
rpt_costcode.Show
End Sub

Private Sub mnu_costjoblevel_Click()
frm_tranxcost.Show
End Sub

Private Sub mnu_costtypecode_Click()
frm_transactiontype.Show
End Sub

Private Sub mnu_costtypecodelistrep_Click()
rpt_transactiontypelist.Show
End Sub

Private Sub mnu_createPassword_Click()
frm_password.Show
End Sub

Private Sub mnu_currencycode_Click()
frm_curr.Show
End Sub

Private Sub mnu_currencycodelistrep_Click()
rpt_currencylist.Show
End Sub

Private Sub mnu_dvar_Click()
rpt_variance.Show
End Sub

Private Sub mnu_editposttransactions_Click()
frm_estpost.Show
End Sub

Private Sub mnu_eicbyjobcharge_Click()
frm_estimatedcost.Show
End Sub

Private Sub mnu_eicbyresc_Click()
frm_quickupdateunitratebyresc.Show
End Sub

Private Sub mnu_eicbyresource_Click()
frm_progresstranx.Show
End Sub

Private Sub mnu_eicjc_Click()
rpt_progressdurationjb.Show
End Sub

Private Sub mnu_eictransactions_Click()
frm_eupdateunitprice.Show
End Sub

Private Sub mnu_estimatedincurredcostbyjobchargerep_Click()
rpt_incurredbyjobcharge.Show
End Sub

Private Sub mnu_estimatedincurredcostbyobsrep_Click()
rpt_obsestimate.Show
End Sub

Private Sub mnu_estimatedincurredcostbyresourcerep_Click()
rpt_incurredbyresource.Show
End Sub

Private Sub mnu_estimatedprogressdurationbyspread_Click()
frm_progressdurationdetails.Show
End Sub

Private Sub mnu_estimatedprogressdurationbyspreadrep_Click()
rpt_progressduration.Show
End Sub

Private Sub mnu_exchangerate_Click()
frm_currencymaster.Show
End Sub

Private Sub mnu_generateeictransactions_Click()
frm_budtran.Show
End Sub

Private Sub mnu_jobchargeno_Click()
frm_jobcharge.Show
End Sub

Private Sub mnu_jobchargenolistrep_Click()
rpt_jobchargelist.Show
End Sub

Private Sub mnu_jobno_Click()
frm_job.Show
End Sub

Private Sub mnu_jobnolistrep_Click()
rpt_jobno.Show
End Sub

Private Sub mnu_l0rep_Click()
rpt_l0.Show
SetParent rpt_l0.hwnd, main.hwnd
End Sub

Private Sub mnu_l1rep_Click()
rpt_l1.Show
SetParent rpt_l1.hwnd, main.hwnd
End Sub

Private Sub mnu_l2rep_Click()
rpt_l2main.Show
SetParent rpt_l2.hwnd, main.hwnd
End Sub

Private Sub mnu_l3rep_Click()
rpt_costdetails.Show
SetParent rpt_costdetails.hwnd, main.hwnd
End Sub

Private Sub mnu_logout_Click()
Unload Me
End Sub

Private Sub mnu_obscode_Click()
frm_respdetails.Show
End Sub

Private Sub mnu_obscoderep_Click()
rpt_obscode.Show
End Sub

Private Sub mnu_ohpiitemcode_Click()
frm_ohpi_itemmaster.Show
End Sub

Private Sub mnu_otherincexpoverheadestrecovery_Click()
frm_l0.Show
End Sub

Private Sub mnu_othertranxcoces_Click()
frm_tranl0.Show
End Sub

Private Sub mnu_projectdiary_Click()
frm_projectremainder.Show
End Sub

Private Sub mnu_projectkey_Click()
frm_projectmaster.Show
End Sub

Private Sub mnu_projectlistrep_Click()
rpt_projectlist.Show
End Sub

Private Sub mnu_resourcecode_Click()
frm_resourcemaster.Show
End Sub

Private Sub mnu_resourcecodelistrep_Click()
rpt_resourcelist.Show
End Sub

Private Sub mnu_resourcemaptoprojectkey1_Click()
frm_resourcemapping.Show
End Sub

Private Sub mnu_resourceresponsibilitycodes_Click()
frm_respdetails.Show
End Sub

Private Sub mnu_resourceresponsibiltycodelistrep_Click()
rpt_rescresp.Show
End Sub

Private Sub mnu_resourcetypecodelistrep_Click()
rpt_resourcetype.Show
End Sub

Private Sub mnu_resourcetypecodes_Click()
frm_resourcetype.Show
End Sub

Private Sub mnu_resourcevendorcodelistrep_Click()
rpt_vendor.Show
End Sub

Private Sub mnu_resourcevendorcodes_Click()
frm_vendor.Show
End Sub

Private Sub mnu_restore_Click()
frm_restore.Show
End Sub

Private Sub mnu_revenuebdgtvoadjbilledunbilled_Click()
frm_revenue.Show
End Sub

Private Sub mnu_revenuedetailsrep_Click()
rpt_revenue.Show
End Sub

Private Sub mnu_revenueprojectkeylevel_Click()
frm_projecttransaction.Show
End Sub

Private Sub mnu_rulesvalidations_Click()
FRM_REPLACE.Show
End Sub

Private Sub mnu_sendmessage_Click()
Form1.Show
End Sub

Private Sub mnu_spreadcode_Click()
frm_Spreadmaster.Show
End Sub

Private Sub mnu_spreadcodelistrep_Click()
rpt_spreadlist.Show
End Sub

Private Sub mnu_subjobno_Click()
frm_subjob.Show
End Sub

Private Sub mnu_subjobnolistrep_Click()
rpt_subjob.Show
End Sub

Private Sub mnu_tablesbcrep_Click()
rpt_buddetail.Show
End Sub

Private Sub mnu_tableseicrep_Click()
rpt_estdetail.Show
End Sub

Private Sub mnu_tranxidforoverheadpitemslistrep_Click()
rpt_othertran.Show
End Sub

Private Sub mnu_uom_Click()
frm_uom.Show
End Sub

Private Sub mnuutilbackup_Click()
 
 
    Cn.Execute "BACKUP DATABASE PCMS to disk='D:\PCMS" & Format(Date, "dd-MM") & ".bak'"
 
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End Sub

Private Sub mnu_uomrep_Click()
rpt_uomlist.Show
End Sub

Private Sub mnu_updatedatesfornaeic_Click()
frm_updatena.Show
End Sub

Private Sub mnu_updatejcb_Click()
updatejobchargebybc.Show
End Sub

Private Sub mnu_updatejobcharge_Click()
updatejobcharge.Show
End Sub

Private Sub mnu_updateqty_Click()
frm_updateqty.Show
End Sub

Private Sub mnu_updateworkcomplete_Click()
frm_na.Show
End Sub

Private Sub mnu_userrights_Click()
frm_userrights.Show
End Sub

Private Sub mnu_variationorderunrealized_Click()
frm_potentialitems.Show
End Sub

Private Sub Option1_Click()
 
TreeView1.Visible = True
End Sub

Private Sub Option2_Click()
TreeView1.Visible = False

End Sub

Private Sub Option3_Click()
TreeView1.Visible = False
 
End Sub

Private Sub mnuL0NotesSignatures_Click()
frmL0Notes.Show
End Sub
Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Key = "Jobno" Then frm_job.Show
If TreeView1.SelectedItem.Key = "SubJobno" Then frm_subjob.Show
If TreeView1.SelectedItem.Key = "JobChargemaster" Then frm_jobcharge.Show
If TreeView1.SelectedItem.Key = "projectmaster" Then frm_projectmaster.Show
If TreeView1.SelectedItem.Key = "Resourcemaster" Then frm_resourcemaster.Show
If TreeView1.SelectedItem.Key = "Currencyexchange" Then frm_currencymaster.Show
If TreeView1.SelectedItem.Key = "Currencymaster" Then frm_curr.Show

If TreeView1.SelectedItem.Key = "otranx" Then frm_tranl0.Show
If TreeView1.SelectedItem.Key = "Chargetype" Then frm_chargetype.Show
If TreeView1.SelectedItem.Key = "L0Notes" Then frmL0Notes.Show

If TreeView1.SelectedItem.Key = "Tranxtype" Then frm_transactiontype.Show
If TreeView1.SelectedItem.Key = "Spread" Then frm_Spreadmaster.Show
If TreeView1.SelectedItem.Key = "vendor" Then frm_vendor.Show
If TreeView1.SelectedItem.Key = "resp" Then frm_responsible.Show
If TreeView1.SelectedItem.Key = "respd" Then frm_respdetails.Show
If TreeView1.SelectedItem.Key = "UOMmaster" Then frm_uom.Show
If TreeView1.SelectedItem.Key = "cc" Then frm_costcode.Show
If TreeView1.SelectedItem.Key = "rty" Then frm_resourcetype.Show
If TreeView1.SelectedItem.Key = "rp" Then frm_resourcemapping.Show
'If TreeView1.SelectedItem.Key = "pty" Then frm_pertype.Show

If TreeView1.SelectedItem.Key = "DurationSchedule" Then frm_budgetedduration.Show
If TreeView1.SelectedItem.Key = "progressschedule" Then frm_progressdurationdetails.Show
If TreeView1.SelectedItem.Key = "budget1" Then frm_costtranx.Show
If TreeView1.SelectedItem.Key = "budget2" Then frm_budgetedcost.Show
If TreeView1.SelectedItem.Key = "ImportExistingJobCharge_B" Then frm_ImportJobChargeResources.Show: frm_ImportJobChargeResources.lblTransactionType.Caption = "B"
If TreeView1.SelectedItem.Key = "incurred1" Then frm_progresstranx.Show
If TreeView1.SelectedItem.Key = "ImportExistingJobCharge_E" Then frm_ImportJobChargeResources.Show: frm_ImportJobChargeResources.lblTransactionType.Caption = "E"
If TreeView1.SelectedItem.Key = "ImportUsingExcelSheet" Then frm_import.Show
If TreeView1.SelectedItem.Key = "incurredProjectbyResource" Then frm_incurredbyResource.Show
If TreeView1.SelectedItem.Key = "incurred2" Then frm_estimatedcost.Show
If TreeView1.SelectedItem.Key = "billed" Then frm_billedcost.Show
'If TreeView1.SelectedItem.Key = "perworkcomp" Then frm_workcomplete.Show
If TreeView1.SelectedItem.Key = "perwork" Then frm_na.Show
If TreeView1.SelectedItem.Key = "urateb" Then frm_updateunitrate.Show
'If TreeView1.SelectedItem.Key = "uratee" Then frm_eupdateunitprice.Show
If TreeView1.SelectedItem.Key = "uratee" Then
frm_quickupdaterescprj.Show
SetParent frm_quickupdaterescprj.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "billunbillrev" Then frm_revenue.Show
If TreeView1.SelectedItem.Key = "uqty" Then frm_updateqty.Show
If TreeView1.SelectedItem.Key = "oitran" Then frm_l0.Show

If TreeView1.SelectedItem.Key = "bdgt" Then rpt_budgetbyresource.Show
If TreeView1.SelectedItem.Key = "bdgtj" Then rpt_budgetbyjobcharge.Show
If TreeView1.SelectedItem.Key = "obsb" Then rpt_obsbudget.Show
If TreeView1.SelectedItem.Key = "obse" Then rpt_obsestimate.Show

If TreeView1.SelectedItem.Key = "estjob" Then rpt_incurredbyjobcharge.Show
If TreeView1.SelectedItem.Key = "estres" Then rpt_incurredbyresource.Show
If TreeView1.SelectedItem.Key = "estbyProjbyres" Then rpt_incurredbyprojectbyresource.Show
'reports
If TreeView1.SelectedItem.Key = "pl" Then rpt_projectlist.Show
If TreeView1.SelectedItem.Key = "sjn" Then rpt_subjob.Show
If TreeView1.SelectedItem.Key = "jn" Then rpt_jobno.Show
If TreeView1.SelectedItem.Key = "ccl" Then rpt_costcode.Show
If TreeView1.SelectedItem.Key = "jcl" Then rpt_jobchargelist.Show
If TreeView1.SelectedItem.Key = "rl" Then rpt_resourcelist.Show
If TreeView1.SelectedItem.Key = "sl" Then rpt_spreadlist.Show
'If TreeView1.SelectedItem.Key = "ttl" Then rpt_transactiontypelist.Show
If TreeView1.SelectedItem.Key = "cl" Then rpt_currencylist.Show
If TreeView1.SelectedItem.Key = "ul" Then rpt_uomlist.Show

If TreeView1.SelectedItem.Key = "obscd" Then rpt_obscode.Show

If TreeView1.SelectedItem.Key = "rtcl" Then rpt_resourcetype.Show
If TreeView1.SelectedItem.Key = "rrcl" Then rpt_rescresp.Show
If TreeView1.SelectedItem.Key = "rvcl" Then rpt_vendor.Show

If TreeView1.SelectedItem.Key = "ctcl" Then rpt_costtype.Show
If TreeView1.SelectedItem.Key = "ttl" Then rpt_othertran.Show

If TreeView1.SelectedItem.Key = "cdr" Then
rpt_costdetails.Show
SetParent rpt_costdetails.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "csr" Then
rpt_l2main.Show
SetParent rpt_l2main.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "msg" Then Form1.Show
If TreeView1.SelectedItem.Key = "lo" Then
rpt_l0.Show
SetParent rpt_l0.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "ohpi" Then frm_ohpi_itemmaster.Show
If TreeView1.SelectedItem.Key = "pi" Then frm_potentialitems.Show
If TreeView1.SelectedItem.Key = "ohi" Then frm_overheaditem.Show
If TreeView1.SelectedItem.Key = "ptranx" Then frm_projecttransaction.Show
If TreeView1.SelectedItem.Key = "ctranx" Then frm_tranxcost.Show
If TreeView1.SelectedItem.Key = "Chngpswd" Then frm_password.Show
If TreeView1.SelectedItem.Key = "psr" Then
rpt_l1.Show
SetParent rpt_l1.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "cmppara" Then frm_parameters.Show
If TreeView1.SelectedItem.Key = "una" Then frm_updatena.Show
If TreeView1.SelectedItem.Key = "bdgtcost1" Then rpt_budgetbycost.Show
If TreeView1.SelectedItem.Key = "buddet" Then rpt_buddetail.Show
If TreeView1.SelectedItem.Key = "estdet" Then rpt_estdetail.Show

If TreeView1.SelectedItem.Key = "ral" Then rpt_revenue.Show

If TreeView1.SelectedItem.Key = "rbu" Then rpt_revenuebu.Show

If TreeView1.SelectedItem.Key = "pdur" Then rpt_progressduration.Show
'''If TreeView1.SelectedItem.Key = "pdurp" Then rpt_progressdurationjb.Show
If TreeView1.SelectedItem.Key = "pdurp" Then rpt_progressdurationjb.Show
If TreeView1.SelectedItem.Key = "durvar" Then rpt_variance.Show
If TreeView1.SelectedItem.Key = "bdur" Then rpt_budgetedduration.Show
If TreeView1.SelectedItem.Key = "bdurp" Then rpt_budgeteddurationjb.Show
If TreeView1.SelectedItem.Key = "budtran" Then frm_budtran.Show
If TreeView1.SelectedItem.Key = "estpost" Then frm_estpost.Show
If TreeView1.SelectedItem.Key = "cal" Then rpt_spreadcostsummary.Show
If TreeView1.SelectedItem.Key = "qupdr" Then frm_quickupdateunitratebyresc.Show

If TreeView1.SelectedItem.Key = "bpbdgt" Then frm_baseline.Show

If TreeView1.SelectedItem.Key = "csrr" Then
report_eicresource.Show
SetParent report_eicresource.hwnd, main.hwnd
End If
If TreeView1.SelectedItem.Key = "ebaResc" Then rpt_estimatedvsbilledbyresource.Show
If TreeView1.SelectedItem.Key = "ebaJC" Then rpt_estimatedvsbilledbyJC.Show
If TreeView1.SelectedItem.Key = "VISUMM" Then rpt_VendorInvoiceSummary.Show
If TreeView1.SelectedItem.Key = "User" Then frm_userrights.Show

If TreeView1.SelectedItem.Key = "backup1" Then
Cn.Execute "BACKUP DATABASE PCMS to disk='D:\PCMS" & Format(Date, "dd-MMM-yyyy") & ".bak'"
 
    If Err.Number = 0 Then MsgBox "Backup Succeded", vbInformation
End If

If TreeView1.SelectedItem.Key = "pdiary" Then frm_projectremainder.Show
If TreeView1.SelectedItem.Key = "ujc" Then updatejobcharge.Show
If TreeView1.SelectedItem.Key = "ujcb" Then updatejobchargebybc.Show
If TreeView1.SelectedItem.Key = "ibdgt" Then frm_export.Show
If TreeView1.SelectedItem.Key = "logo" Then
Unload Me
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
 

'''Public Sub progcost()
'''
'''Dim gtotal As Double
'''gtotal = 0
'''Dim ntotal As Double
'''ntotal = 0
'''Dim iddd As Integer
'''iddd = 0
'''Dim fldata As New ADODB.Recordset
'''If fldata.State Then fldata.Close
'''fldata.Open "select * from cost where  bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2
'''
'''
'''    While Not fldata.EOF
'''
'''     iddd = fldata!bd_id
'''mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
'''mmm = Split(fldata!bd_jobcharge, "  -  ", Len(fldata!bd_jobcharge), vbTextCompare)
'''mmmm = Split(fldata!bd_resccode, "  -  ", Len(fldata!bd_resccode), vbTextCompare)
'''
'''Dim dt1 As Date
'''Dim dt2 As Date
'''Dim pp As New ADODB.Recordset
'''If pp.State Then pp.Close
'''pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
'''If Not pp.EOF Then
'''dt1 = pp!prgs_startdate
'''dt2 = pp!prgs_enddate
'''End If
'''
'''Dim fldata2 As New ADODB.Recordset
'''If fldata2.State Then fldata2.Close
'''fldata2.Open "select * from cost where    bd_jobcharge='" & fldata!bd_jobcharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'
'''
'''    If Not fldata2.EOF Then
'''
'''
'''
'''            fldata2!bd_sdate = dt1
'''            fldata2!bd_edate = dt2
'''                    If dt1 <= DTPcutdate1.Value And dt2 <= DTPcutdate1.Value Then
'''                    a = dt2 - dt1
'''                    c = 0
'''                    ElseIf dt1 <= DTPcutdate1.Value And dt2 >= DTPcutdate1.Value Then
'''                    a = DTPcutdate1.Value - dt1
'''                    c = dt2 - DTPcutdate1.Value
'''
'''                    Else
'''                    a = 0
'''                    c = dt2 - dt1
'''                    End If
'''            Dim d As Double
'''            d = 0
'''            Dim f As Double
'''            f = 0
'''            fldata2!bd_days = a
'''            fldata2!bd_e_days = c
'''            d = CDbl(a) * CDbl(fldata!bd_qty)
'''            fldata2!bd_e_tqty = CDbl(c) * CDbl(fldata!bd_qty)
'''            fldata2!bd_tqty = d
'''            fldata2!bd_extdamt = CDbl(d) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
'''            fldata2!bd_e_extdamt = CDbl(fldata2!bd_e_tqty) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
'''            fldata2.Update
'''
'''    End If
'''
'''        fldata.MoveNext
'''    Wend
'''
'''
'''Dim cid As Integer
'''Dim cd As New ADODB.Recordset
'''If cd.State Then cd.Close
'''cd.Open "select * from cost where bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
'''While Not cd.EOF
'''
'''
'''If cd!bd_chk = 1 Then
'''
'''
'''                    If cd!bd_sdate <= DTPcutdate1.Value And cd!bd_edate <= DTPcutdate1.Value Then
'''                    a = cd!bd_edate - cd!bd_sdate
'''                    c = 0
'''                    ElseIf cd!bd_sdate <= DTPcutdate1.Value And cd!bd_edate >= DTPcutdate1.Value Then
'''                    a = DTPcutdate1.Value - cd!bd_sdate
'''                    c = cd!bd_edate - DTPcutdate1.Value
'''
'''                    Else
'''                    a = 0
'''                    c = cd!bd_edate - cd!bd_sdate
'''                    End If
'''                    cd!bd_days = a
'''                    cd!bd_e_days = c
'''                    If IsNull(cd!bd_days) = True Then
'''                    cd!bd_tqty = cd!bd_qty
'''                    Else
'''                    cd!bd_tqty = cd!bd_qty * cd!bd_days
'''                    End If
'''                    cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
'''                    If IsNull(cd!bd_e_days) = True Then
'''                    cd!bd_e_tqty = cd!bd_qty
'''                    Else
'''                    cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
'''                    End If
'''                    cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
''' ElseIf cd!bd_chk = 0 Then
'''
'''cd!bd_edate = cd!bd_sdate
'''
'''                    If cd!bd_sdate <= DTPcutdate1.Value And cd!bd_edate <= DTPcutdate1.Value Then
'''                                    cd!bd_tqty = cd!bd_qty
'''                                    cd!bd_days = Null
'''                                    cd!bd_e_days = 0
'''                                    cd!bd_e_tqty = 0
'''                       Else
'''
'''                                    cd!bd_e_tqty = cd!bd_qty
'''                                    cd!bd_e_days = Null
'''                                    cd!bd_days = 0
'''                                    cd!bd_tqty = 0
'''                    End If
'''
'''
'''                    If IsNull(cd!bd_days) = True Then
'''                    cd!bd_tqty = cd!bd_qty
'''                    Else
'''                    cd!bd_tqty = cd!bd_qty * cd!bd_days
'''                    End If
'''                    cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
'''                    If IsNull(cd!bd_e_days) = True Then
'''                    cd!bd_e_tqty = cd!bd_qty
'''                    Else
'''                    cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
'''                    End If
'''                    cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
'''
'''
''' End If
'''cd.Update
'''
'''cd.MoveNext
'''Wend
'''End Sub
'''

Public Sub mnuuser()
On Error Resume Next
    If Mid(a, 1, 1) = 1 Then
    mnumaster.Visible = True
        If Mid(a, 2, 1) = 1 Then
        mnu_projectcodes.Visible = True
                If Mid(a, 3, 1) = 1 Then
                mnu_projectkey.Visible = True
                Else
                mnu_projectkey.Visible = False
                End If
                If Mid(a, 4, 1) = 1 Then
                mnu_jobno.Visible = True
                Else
                mnu_jobno.Visible = False
                End If
                If Mid(a, 5, 1) = 1 Then
                mnu_subjobno.Visible = True
                Else
                mnu_subjobno.Visible = False
                End If
                If Mid(a, 6, 1) = 1 Then
                mnu_jobchargeno.Visible = True
                Else
                mnu_jobchargeno.Visible = False
                End If
                If Mid(a, 7, 1) = 1 Then
                mnu_obscode.Visible = True
                Else
                mnu_obscode.Visible = False
                End If
        Else
        mnu_projectcodes.Visible = False
        End If
        
        
If Mid(a, 9, 1) = 1 Then
mnu_resourcecodes.Visible = True
    If Mid(a, 10, 1) = 1 Then
    mnu_resourcetypecodes.Visible = True
    Else
    mnu_resourcetypecodes.Visible = False
    End If
    If Mid(a, 11, 1) = 1 Then
    mnu_resourceresponsibilitycodes.Visible = True
    Else
    mnu_resourceresponsibilitycodes.Visible = False
    End If
    If Mid(a, 12, 1) = 1 Then
    mnu_resourcevendorcodes.Visible = True
    Else
    mnu_resourcevendorcodes.Visible = False
    End If
    If Mid(a, 13, 1) = 1 Then
    mnu_resourcecode.Visible = True
    Else
    mnu_resourcecode.Visible = False
    End If
    If Mid(a, 14, 1) = 1 Then
    mnu_resourcemaptoprojectkey1.Visible = True
    Else
    mnu_resourcemaptoprojectkey1.Visible = False
    End If
Else
mnu_resourcecodes.Visible = False
End If
                                
                If Mid(a, 15, 1) = 1 Then
                mnu_othercodes.Visible = True
                    If Mid(a, 16, 1) = 1 Then
                    mnu_spreadcode.Visible = True
                    Else
                    mnu_spreadcode.Visible = False
                    End If
                    If Mid(a, 17, 1) = 1 Then
                    mnu_costtypecode.Visible = True
                    Else
                    mnu_costtypecode.Visible = False
                    End If
                    If Mid(a, 18, 1) = 1 Then
                    mnu_uom.Visible = True
                    Else
                    mnu_uom.Visible = False
                    End If
                    If Mid(a, 19, 1) = 1 Then
                    mnu_currencycode.Visible = True
                    Else
                    mnu_currencycode.Visible = False
                    End If
                    If Mid(a, 20, 1) = 1 Then
                    mnu_exchangerate.Visible = True
                    Else
                    mnu_exchangerate.Visible = False
                    End If
                    If Mid(a, 21, 1) = 1 Then
                    mnu_ohpiitemcode.Visible = True
                    Else
                    mnu_ohpiitemcode.Visible = False
                    End If
                    If Mid(a, 22, 1) = 1 Then
                    mnu_othertranxcoces.Visible = True
                    Else
                    mnu_othertranxcoces.Visible = False
                    End If
                Else
                mnu_othercodes.Visible = False
                End If
    Else
    mnumaster.Visible = False
    End If 'end of master forms
    
    
    
    
    
    
If Mid(b, 1, 1) = 1 Then
mnu_Transactions.Visible = True
    If Mid(b, 2, 1) = 1 Then
    mnu_budgeteddetails.Visible = True
            If Mid(b, 3, 1) = 1 Then
            mnu_budgeteddurationbyspread.Visible = True
            End If
            If Mid(b, 4, 1) = 1 Then
            mnu_budgetedcostdetails.Visible = True
                If Mid(b, 5, 1) = 1 Then
                mnu_bcbyresource.Visible = True
                Else
                mnu_bcbyresource.Visible = False
                End If
                If Mid(b, 6, 1) = 1 Then
                mnu_bcbyjobcharge.Visible = True
                Else
                mnu_bcbyjobcharge.Visible = False
                End If
           Else
           mnu_budgetedcostdetails.Visible = False
           End If
    
    Else
    mnu_budgeteddetails.Visible = False
    End If


        If Mid(b, 7, 1) = 1 Then
        mnu_generateeicdetailsbybudget.Visible = True
                        If Mid(b, 8, 1) = 1 Then
                        mnu_generateeictransactions.Visible = True
                        Else
                        mnu_generateeictransactions.Visible = False
                        End If
                        If Mid(b, 9, 1) = 1 Then
                        mnu_editposttransactions.Visible = True
                        Else
                        mnu_editposttransactions.Visible = False
                        End If
        Else
        mnu_generateeicdetailsbybudget.Visible = False
        End If
                If Mid(b, 10, 1) = 1 Then
                mnu_estimateddetails.Visible = True
                If Mid(b, 11, 1) = 1 Then
                mnu_estimatedprogressdurationbyspread.Visible = True
                Else
                mnu_estimatedprogressdurationbyspread.Visible = False
                End If
                            If Mid(b, 12, 1) = 1 Then
                            mnu_estimatedincurredcostdetails.Visible = True
                                            If Mid(b, 13, 1) = 1 Then
                                            mnu_eicbyresource.Visible = True
                                            Else
                                            mnu_eicbyresource.Visible = False
                                            End If
                                            If Mid(b, 14, 1) = 1 Then
                                            mnu_eicbyresource.Visible = True
                                            Else
                                            mnu_eicbyresource.Visible = False
                                            End If
                            Else
                            mnu_estimatedincurredcostdetails.Visible = False
                            End If
                Else
                mnu_estimateddetails.Visible = False
                End If
 

If Mid(b, 15, 1) = 1 Then
mnu_otherdetails.Visible = True
            If Mid(b, 16, 1) = 1 Then
            mnu_revenuebdgtvoadjbilledunbilled.Visible = True
            Else
            mnu_revenuebdgtvoadjbilledunbilled.Visible = False
            End If
            If Mid(b, 17, 1) = 1 Then
            mnu_otherincexpoverheadestrecovery.Visible = True
            Else
            mnu_otherincexpoverheadestrecovery.Visible = False
            End If
            If Mid(b, 18, 1) = 1 Then
            mnu_variationorderunrealized.Visible = True
            Else
            mnu_variationorderunrealized.Visible = False
            End If
            If Mid(b, 19, 1) = 1 Then
            mnu_billedcost.Visible = True
            Else
            mnu_billedcost.Visible = False
            End If
            If Mid(b, 20, 1) = 1 Then
            mnu_projectdiary.Visible = True
            Else
            mnu_projectdiary.Visible = False
            End If
            If Mid(b, 21, 1) = 1 Then
            mnu_bpbdgt.Visible = True
            Else
            mnu_bpbdgt.Visible = False
            End If
 Else
 mnu_otherdetails.Visible = False
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Mid(b, 22, 1) = 1 Then
mnu_quickupdates.Visible = True
                If Mid(b, 23, 1) = 1 Then
                mnu_updateworkcomplete.Visible = True
                Else
                mnu_updateworkcomplete.Visible = False
                End If
                If Mid(b, 24, 1) = 1 Then
                mnu_updateunitrate.Visible = True
                            If Mid(b, 25, 1) = 1 Then
                            mnu_bctransactions.Visible = True
                            Else
                            mnu_bctransactions.Visible = False
                            End If
                            If Mid(b, 26, 1) = 1 Then
                            mnu_eictransactions.Visible = True
                            Else
                            mnu_eictransactions.Visible = False
                            End If
                Else
                mnu_updateunitrate.Visible = False
                End If
                    If Mid(b, 27, 1) = 1 Then
                    mnu_updatedatesfornaeic.Visible = True
                    Else
                    mnu_updatedatesfornaeic.Visible = False
                    End If
                    If Mid(b, 28, 1) = 1 Then
                    mnu_updateqty.Visible = True
                    Else
                    mnu_updateqty.Visible = False
                    End If
                    If Mid(b, 29, 1) = 1 Then
                    mnu_updatejobcharge.Visible = True
                    Else
                    mnu_updatejobcharge.Visible = False
                    End If
                                        If Mid(b, 30, 1) = 1 Then
                    mnu_updatejcb.Visible = True
                    Else
                    mnu_updatejcb.Visible = False
                    End If
Else
mnu_quickupdates.Visible = False
End If



If Mid(b, 30, 1) = 1 Then
mnu_periodendupdates.Visible = True
            If Mid(b, 31, 1) = 1 Then
            mnu_revenueprojectkeylevel.Visible = True
            Else
            mnu_revenueprojectkeylevel.Visible = False
            End If
            If Mid(b, 32, 1) = 1 Then
            mnu_costjoblevel.Visible = True
            Else
            mnu_costjoblevel.Visible = False
            End If
Else
mnu_periodendupdates.Visible = False
End If


Else
mnu_Transactions.Visible = False
End If 'transactions

' Reports
If Mid(c, 1, 1) = 1 Then
mnu_reports.Visible = True

If Mid(c, 2, 1) = 1 Then
mnu_masterlists.Visible = True
    If Mid(c, 3, 1) = 1 Then
        mnu_projectrep.Visible = True
                            If Mid(c, 4, 1) = 1 Then
                            mnu_projectlistrep.Visible = True
                            Else
                            mnu_projectlistrep.Visible = False
                            
                            End If
                            If Mid(c, 5, 1) = 1 Then
                            mnu_jobnolistrep.Visible = True
                            Else
                            mnu_jobnolistrep.Visible = False
                            End If
                            If Mid(c, 6, 1) = 1 Then
                            mnu_subjobnolistrep.Visible = True
                            Else
                            mnu_subjobnolistrep.Visible = False
                            End If
                            If Mid(c, 7, 1) = 1 Then
                            mnu_jobchargenolistrep.Visible = True
                            Else
                            mnu_jobchargenolistrep.Visible = False
                            End If
                            If Mid(c, 8, 1) = 1 Then
                            mnu_obscoderep.Visible = True
                            Else
                            mnu_obscoderep.Visible = False
                            End If
                            If Mid(c, 9, 1) = 1 Then
                            mnu_costcodelistrep.Visible = True
                            Else
                            mnu_costcodelistrep.Visible = False
                            End If
        Else
        mnu_projectrep.Visible = False
        mnu_subjobnolistrep.Visible = False
        mnu_jobchargenolistrep.Visible = False
        mnu_obscoderep.Visible = False
        mnu_costcodelistrep.Visible = False
        End If
If Mid(c, 10, 1) = 1 Then
        mnu_resourcerep.Visible = True
        If Mid(c, 11, 1) = 1 Then
        mnu_resourcetypecodelistrep.Visible = True
        Else
        mnu_resourcetypecodelistrep.Visible = False
        End If
        If Mid(c, 12, 1) = 1 Then
        mnu_resourceresponsibiltycodelistrep.Visible = True
        Else
        mnu_resourceresponsibiltycodelistrep.Visible = False
        End If
        If Mid(c, 13, 1) = 1 Then
        mnu_resourcevendorcodelistrep.Visible = True
        Else
        mnu_resourcevendorcodelistrep.Visible = False
        End If
        If Mid(c, 14, 1) = 1 Then
        mnu_resourcecodelistrep.Visible = True
        Else
        mnu_resourcecodelistrep.Visible = False
        End If
        Else
        mnu_resourcerep.Visible = False
End If

    If Mid(c, 15, 1) = 1 Then
    mnu_othesrep.Visible = True
            If Mid(c, 16, 1) = 1 Then
            mnu_spreadcodelistrep.Visible = True
            Else
            mnu_spreadcodelistrep.Visible = False
            End If
            If Mid(c, 17, 1) = 1 Then
            mnu_costtypecodelistrep.Visible = True
            Else
            mnu_costtypecodelistrep.Visible = False
            End If
            If Mid(c, 18, 1) = 1 Then
            mnu_uomrep.Visible = True
            Else
            mnu_uomrep.Visible = False
            End If
            If Mid(c, 19, 1) = 1 Then
            mnu_currencycodelistrep.Visible = True
            Else
            mnu_currencycodelistrep.Visible = False
            End If
            If Mid(c, 20, 1) = 1 Then
            mnu_exchangeratelistrep.Visible = True
            Else
            mnu_exchangeratelistrep.Visible = False
            End If
            If Mid(c, 21, 1) = 1 Then
            mnu_tranxidforoverheadpitemslistrep.Visible = True
            Else
            mnu_tranxidforoverheadpitemslistrep.Visible = False
            End If
     Else
     mnu_othesrep.Visible = False
        
        mnu_spreadcodelistrep.Visible = False
        
        mnu_costtypecodelistrep.Visible = False
        
        mnu_uomrep.Visible = False
        
        mnu_currencycodelistrep.Visible = False
        
        mnu_exchangeratelistrep.Visible = False
        
        mnu_tranxidforoverheadpitemslistrep.Visible = False
        
    End If
Else
mnu_masterlists.Visible = False
 
        
        mnu_projectrep.Visible = False
        
        mnu_projectlistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_subjobnolistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_obscoderep.Visible = False
        
        mnu_costcodelistrep.Visible = False
        
        mnu_resourcerep.Visible = False
        
        mnu_resourcetypecodelistrep.Visible = False
        
        mnu_resourceresponsibiltycodelistrep.Visible = False
        
        mnu_resourcevendorcodelistrep.Visible = False
        
        mnu_resourcecodelistrep.Visible = False
        
       mnu_othesrep.Visible = False
        
        mnu_spreadcodelistrep.Visible = False
        
        mnu_costtypecodelistrep.Visible = False
        
        mnu_uomrep.Visible = False
        
        mnu_currencycodelistrep.Visible = False
        
        mnu_exchangeratelistrep.Visible = False
        
        mnu_tranxidforoverheadpitemslistrep.Visible = False
        


End If 'master
Else
mnu_reports.Visible = False
mnu_masterlists.Visible = False
 
        
        mnu_projectrep.Visible = False
        
        mnu_projectlistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_subjobnolistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_obscoderep.Visible = False
        
        mnu_costcodelistrep.Visible = False
        
        mnu_resourcerep.Visible = False
        
        mnu_resourcetypecodelistrep.Visible = False
        
        mnu_resourceresponsibiltycodelistrep.Visible = False
        
        mnu_resourcevendorcodelistrep.Visible = False
        
        mnu_resourcecodelistrep.Visible = False
        
       mnu_othesrep.Visible = False
        
        mnu_spreadcodelistrep.Visible = False
        
        mnu_costtypecodelistrep.Visible = False
        
        mnu_uomrep.Visible = False
        
        mnu_currencycodelistrep.Visible = False
        
        mnu_exchangeratelistrep.Visible = False
        
        mnu_tranxidforoverheadpitemslistrep.Visible = False
        

End If
 

If Mid(d, 1, 1) = 1 Then
mnu_budgetedreportsrep.Visible = True
    If Mid(d, 2, 1) = 1 Then
   mnu_budgetedduartionbyspreadrep.Visible = True
    End If
        If Mid(d, 3, 1) = 1 Then
            mnu_budgetedcostdetailsrep1.Visible = True
                If Mid(d, 4, 1) = 1 Then
                mnu_bcbyresourcerep.Visible = True
                Else
                mnu_bcbyresourcerep.Visible = False
                End If
                If Mid(d, 5, 1) = 1 Then
                mnu_bcbyresourcecostcoderep.Visible = True
                Else
                mnu_bcbyresourcecostcoderep.Visible = False
                End If
                If Mid(d, 6, 1) = 1 Then
                mnu_bcbyjobcharge.Visible = True
                Else
                mnu_bcbyjobcharge.Visible = False
                End If
                If Mid(d, 7, 1) = 1 Then
                mnu_bcbyobsrep.Visible = True
                Else
                mnu_bcbyobsrep.Visible = False
                End If
                Else
                                mnu_budgetedcostdetailsrep1.Visible = False
                                
                                
                                mnu_bcbyresourcerep.Visible = False
                                
                                mnu_bcbyresourcecostcoderep.Visible = False
                                
                                mnu_bcbyjobcharge.Visible = False
                                
                                mnu_bcbyobsrep.Visible = False
        End If
Else
        mnu_budgetedreportsrep.Visible = False
         
        mnu_budgetedduartionbyspreadrep.Visible = False
        
        mnu_budgetedcostdetailsrep1.Visible = False
        
        mnu_bcbyresourcerep.Visible = False
        
        mnu_bcbyresourcecostcoderep.Visible = False
        
        mnu_bcbyjobcharge.Visible = False
        
        mnu_bcbyobsrep.Visible = False
End If
        If Mid(d, 8, 1) = 1 Then
        mnu_estimatedincurredreportsrep.Visible = True
        If Mid(d, 9, 1) = 1 Then
        mnu_estimatedprogressdurationbyspreadrep.Visible = True
        End If
            If Mid(d, 10, 1) = 1 Then
            mnu_estimatedincurredcostdetailsrep.Visible = True
            If Mid(d, 11, 1) = 1 Then
            mnu_estimatedincurredcostbyresourcerep.Visible = True
            Else
            mnu_estimatedincurredcostbyresourcerep.Visible = False
            End If
            If Mid(d, 12, 1) = 1 Then
            mnu_estimatedincurredcostbyjobchargerep.Visible = True
            Else
            mnu_estimatedincurredcostbyjobchargerep.Visible = False
            End If
            If Mid(d, 13, 1) = 1 Then
            mnu_estimatedincurredcostbyobsrep.Visible = True
            Else
            mnu_estimatedincurredcostbyobsrep.Visible = False
            End If
            Else
                    mnu_estimatedincurredcostdetailsrep.Visible = False
                    
                    mnu_estimatedincurredcostbyresourcerep.Visible = False
                    
                    mnu_estimatedincurredcostbyjobchargerep.Visible = False
                    
                    mnu_estimatedincurredcostbyobsrep.Visible = False
            End If
        Else
                    mnu_estimatedincurredreportsrep.Visible = False
                    
                    
                    mnu_estimatedprogressdurationbyspreadrep.Visible = False
                    
                    mnu_estimatedincurredcostdetailsrep.Visible = False
                    
                    mnu_estimatedincurredcostbyresourcerep.Visible = False
                    
                    mnu_estimatedincurredcostbyjobchargerep.Visible = False
                    
                    mnu_estimatedincurredcostbyobsrep.Visible = False
        End If
            If Mid(d, 14, 1) = 1 Then
            mnu_managementreports.Visible = True
            If Mid(d, 15, 1) = 1 Then
            mnu_l0rep.Visible = True
            Else
            mnu_l0rep.Visible = False
            End If
            If Mid(d, 16, 1) = 1 Then
            mnu_l1rep.Visible = True
            Else
            mnu_l1rep.Visible = False
            End If
            If Mid(d, 17, 1) = 1 Then
            mnu_l2rep.Visible = True
            Else
            mnu_l2rep.Visible = False
            End If
            If Mid(d, 18, 1) = 1 Then
            mnu_l3rep.Visible = True
            Else
            mnu_l3rep.Visible = False
            End If
            Else
            mnu_managementreports.Visible = False
            mnu_l0rep.Visible = False
            
            mnu_l1rep.Visible = False
            
            mnu_l2rep.Visible = False
            
            mnu_l3rep.Visible = False
            End If

If Mid(d, 19, 1) = 1 Then
mnu_miscelleneousrep.Visible = True
                                    If Mid(d, 20, 1) = 1 Then
                                    mnu_revenuedetailsrep.Visible = True
                                        If Mid(d, 21, 1) = 1 Then
                                        mnu_budgetedrevenuevariationorderrep.Visible = True
                                        Else
                                        mnu_budgetedrevenuevariationorderrep.Visible = False
                                        End If
                                        If Mid(d, 22, 1) = 1 Then
                                        mnu_revenuebilledunbilledrep.Visible = True
                                        Else
                                        mnu_revenuebilledunbilledrep.Visible = False
                                        End If
                                    Else
                                    
                                    mnu_revenuedetailsrep.Visible = False
                                    
                                    mnu_budgetedrevenuevariationorderrep.Visible = False
                                    
                                    mnu_revenuebilledunbilledrep.Visible = False
                                    End If


                                If Mid(d, 23, 1) = 1 Then
                                mnu_costaccruallistrep.Visible = True
                                Else
                                mnu_costaccruallistrep.Visible = False
                                End If
                                If Mid(d, 24, 1) = 1 Then
                                mnu_costsummarybyresourcerep.Visible = True
                                Else
                                mnu_costsummarybyresourcerep.Visible = False
                                End If
                                If Mid(d, 25, 1) = 1 Then
                                mnu_estimatebilledcostrep.Visible = True
                                Else
                                mnu_estimatebilledcostrep.Visible = False
                                End If
                        If Mid(d, 26, 1) = 1 Then
                        mnu_tablesrep.Visible = True
                        
                                    If Mid(d, 27, 1) = 1 Then
                                    mnu_tablesbcrep.Visible = True
                                    Else
                                    mnu_tablesbcrep.Visible = False
                                    End If
                                    If Mid(d, 28, 1) = 1 Then
                                    mnu_tableseicrep.Visible = True
                                    Else
                                    mnu_tableseicrep.Visible = False
                                    End If
                        Else
                        mnu_tablesrep.Visible = False
                        
                        mnu_tablesbcrep.Visible = False
                                
                        mnu_tableseicrep.Visible = False
                        End If
Else
 
mnu_miscelleneousrep.Visible = False

mnu_revenuedetailsrep.Visible = False

mnu_budgetedrevenuevariationorderrep.Visible = False

mnu_revenuebilledunbilledrep.Visible = False

mnu_costaccruallistrep.Visible = False

mnu_costsummarybyresourcerep.Visible = False

mnu_estimatebilledcostrep.Visible = False

mnu_tablesrep.Visible = False

mnu_tablesbcrep.Visible = False
        
mnu_tableseicrep.Visible = False
End If


If Mid(f, 2, 1) = 1 Then
mnu_utilities.Visible = True
If Mid(f, 3, 1) = 1 Then
mnu_BackUp.Visible = True
Else
mnu_BackUp.Visible = False
End If
If Mid(f, 4, 1) = 1 Then
mnu_restore.Visible = True
Else
mnu_restore.Visible = False
End If
If Mid(f, 5, 1) = 1 Then
mnu_sendmessage.Visible = True
Else
mnu_sendmessage.Visible = False
End If
Else
mnu_utilities.Visible = False

mnu_BackUp.Visible = False
mnu_restore.Visible = False
mnu_sendmessage.Visible = False
End If

' Administration
If Mid(f, 6, 1) = 1 Then
mnu_administration.Visible = True
            If Mid(f, 7, 1) = 1 Then
            mnu_companyparameter.Visible = True
            Else
            mnu_companyparameter.Visible = False
            End If
            If Mid(f, 8, 1) = 1 Then
            mnu_createPassword.Visible = True
            Else
            mnu_createPassword.Visible = False
            End If
            If Mid(f, 9, 1) = 1 Then
            mnu_userrights.Visible = True
            Else
            mnu_userrights.Visible = False
            End If
            If Mid(f, 10, 1) = 1 Then
            mnu_rulesvalidations.Visible = False
            Else
            mnu_rulesvalidations.Visible = False
            End If
Else
        mnu_administration.Visible = False
        
        
        mnu_companyparameter.Visible = False
        
        mnu_createPassword.Visible = False
        
        mnu_userrights.Visible = False
        
        mnu_rulesvalidations.Visible = False
        
        
        
End If
' Help
If Mid(f, 11, 1) = 1 Then
mnu_help.Visible = True
If Mid(f, 12, 1) = 1 Then
mnu_dataflow.Visible = True
Else
mnu_dataflow.Visible = False
End If
If Mid(f, 13, 1) = 1 Then
mnu_formhelp.Visible = True
Else
mnu_formhelp.Visible = False
End If
Else
        mnu_help.Visible = False
        
        
        mnu_dataflow.Visible = False
        
        mnu_formhelp.Visible = False
End If
  
    
End Sub

Public Sub userinvisible()
   
        mnumaster.Visible = False
        
        mnu_projectcodes.Visible = False
        
        mnu_projectkey.Visible = False
        
        mnu_jobno.Visible = False
        
        mnu_subjobno.Visible = False
        
        mnu_jobchargeno.Visible = False
        
        mnu_obscode.Visible = False
        
        mnu_resourcecodes.Visible = False
        
        mnu_resourcetypecodes.Visible = False
        
        mnu_resourceresponsibilitycodes.Visible = False
        
        mnu_resourcevendorcodes.Visible = False
        
        mnu_resourcecode.Visible = False
        
        mnu_resourcemaptoprojectkey1.Visible = False
        
        
        mnu_othercodes.Visible = False
        
        mnu_spreadcode.Visible = False
        
        mnu_costtypecode.Visible = False
        
        mnu_uom.Visible = False
        
        mnu_currencycode.Visible = False
        
        mnu_exchangerate.Visible = False
        
        mnu_ohpiitemcode.Visible = False
        
        mnu_othertranxcoces.Visible = False
        
        mnu_Transactions.Visible = False
        
        mnu_budgeteddetails.Visible = False
        
        mnu_budgeteddurationbyspread.Visible = False
        
        mnu_budgetedcostdetails.Visible = False
        
        mnu_bcbyresource.Visible = False
        
        mnu_bcbyjobcharge.Visible = False
        
        mnu_generateeicdetailsbybudget.Visible = False
        
        mnu_generateeictransactions.Visible = False
        
        mnu_editposttransactions.Visible = False
        
        mnu_estimateddetails.Visible = False
        
        mnu_estimatedprogressdurationbyspread.Visible = False
        
        mnu_estimatedincurredcostdetails.Visible = False
        
        mnu_eicbyresource.Visible = False
        
        mnu_eicbyresource.Visible = False
        
        mnu_otherdetails.Visible = False
        
        mnu_revenuebdgtvoadjbilledunbilled.Visible = False
        
        mnu_otherincexpoverheadestrecovery.Visible = False
        
        mnu_variationorderunrealized.Visible = False
        
        mnu_billedcost.Visible = False
        mnu_projectdiary.Visible = False
        mnu_updatejobcharge.Visible = False
        mnu_updatejcb.Visible = False
        mnu_quickupdates.Visible = False
        
        mnu_updateworkcomplete.Visible = False
        
        mnu_updateunitrate.Visible = False
        
        mnu_bctransactions.Visible = False
        
        mnu_eictransactions.Visible = False
        
        mnu_updatedatesfornaeic.Visible = False
        
        
        mnu_periodendupdates.Visible = False
        
        mnu_revenueprojectkeylevel.Visible = False
        
        mnu_costjoblevel.Visible = False
        
        ' Reports
        
        mnu_reports.Visible = False
        
        mnu_masterlists.Visible = False
        
        mnu_projectrep.Visible = False
        
        mnu_projectlistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_subjobnolistrep.Visible = False
        
        mnu_jobchargenolistrep.Visible = False
        
        mnu_obscoderep.Visible = False
        
        mnu_costcodelistrep.Visible = False
        
        mnu_resourcerep.Visible = False
        
        mnu_resourcetypecodelistrep.Visible = False
        
        mnu_resourceresponsibiltycodelistrep.Visible = False
        
        mnu_resourcevendorcodelistrep.Visible = False
        
        mnu_resourcecodelistrep.Visible = False
        
        mnu_othesrep.Visible = False
        
        mnu_spreadcodelistrep.Visible = False
        
        mnu_costtypecodelistrep.Visible = False
        
        mnu_uomrep.Visible = False
        
        mnu_currencycodelistrep.Visible = False
        
        mnu_exchangeratelistrep.Visible = False
        
        mnu_tranxidforoverheadpitemslistrep.Visible = False
        
        
        
        
        mnu_budgetedreportsrep.Visible = False
        
        mnu_budgetedduartionbyspreadrep.Visible = False
        
        mnu_budgetedcostdetailsrep1.Visible = False
        
        mnu_bcbyresourcerep.Visible = False
        
        mnu_bcbyresourcecostcoderep.Visible = False
        
        mnu_bcbyjobcharge.Visible = False
        
        mnu_bcbyobsrep.Visible = False
        
        mnu_estimatedincurredreportsrep.Visible = False
        
        mnu_estimatedprogressdurationbyspreadrep.Visible = False
        
        mnu_estimatedincurredcostdetailsrep.Visible = False
        
        mnu_estimatedincurredcostbyresourcerep.Visible = False
        
        mnu_estimatedincurredcostbyjobchargerep.Visible = False
        
        mnu_estimatedincurredcostbyobsrep.Visible = False
        
        mnu_managementreports.Visible = False
        
        mnu_l0rep.Visible = False
        
        mnu_l1rep.Visible = False
        
        mnu_l2rep.Visible = False
        
        mnu_l3rep.Visible = False
        
        
        
        mnu_miscelleneousrep.Visible = False
        
        mnu_revenuedetailsrep.Visible = False
        
        mnu_budgetedrevenuevariationorderrep.Visible = False
        
        mnu_revenuebilledunbilledrep.Visible = False
        
        
        
        
        mnu_costaccruallistrep.Visible = False
        
        mnu_costsummarybyresourcerep.Visible = False
        
        mnu_estimatebilledcostrep.Visible = False
        
        mnu_tablesrep.Visible = False
        
        mnu_tablesbcrep.Visible = False
        
        mnu_tableseicrep.Visible = False
        
        
        
        
        mnu_utilities.Visible = False
        mnu_BackUp.Visible = False
        mnu_restore.Visible = False
        mnu_sendmessage.Visible = False
        
        
        
        mnu_administration.Visible = False
        
        mnu_companyparameter.Visible = False
        
        mnu_createPassword.Visible = False
        
        mnu_userrights.Visible = False
        
        mnu_rulesvalidations.Visible = False
        
        
        mnu_help.Visible = False
        
        mnu_dataflow.Visible = False
        
        mnu_formhelp.Visible = False

End Sub
