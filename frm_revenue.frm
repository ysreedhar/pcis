VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_revenue 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REVENUE-BDGT/VO/ADJ/BILLED/UNBILLED"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Budget Details"
      TabPicture(0)   =   "frm_revenue.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Variation Order (+)"
      TabPicture(1)   =   "frm_revenue.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Variation Order (-)"
      TabPicture(2)   =   "frm_revenue.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Billed Revenue Details"
      TabPicture(3)   =   "frm_revenue.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Budgeted VO"
      TabPicture(4)   =   "frm_revenue.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flex_grid4 
            Height          =   6615
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   11668
            _Version        =   393216
            Rows            =   3
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   14450266
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   17
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flex_grid3 
            Height          =   6495
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   11456
            _Version        =   393216
            Rows            =   3
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   14450266
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   15
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flex_grid2 
            Height          =   6495
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   11456
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   14450266
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   13
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flex_grid1 
            Height          =   6495
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   11456
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   14450266
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   11
         Top             =   360
         Width           =   11055
         Begin MSFlexGridLib.MSFlexGrid flex_grid 
            Height          =   6615
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   11668
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   14450266
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11055
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Revenue Summary"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   4440
         TabIndex        =   6
         Top             =   600
         Width           =   6495
         Begin VB.TextBox txt_bvo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1080
            TabIndex        =   36
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4320
            TabIndex        =   32
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txt_ubl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4320
            TabIndex        =   31
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txt_bld 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4320
            TabIndex        =   30
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txt_adj 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1080
            TabIndex        =   29
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txt_vos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1080
            TabIndex        =   28
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txt_bgt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1080
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BVO(RM)"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UBL(RM)"
            Height          =   195
            Left            =   3120
            TabIndex        =   33
            Top             =   600
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Revn(RM)"
            Height          =   195
            Left            =   3120
            TabIndex        =   23
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BLD(RM)"
            Height          =   195
            Left            =   3120
            TabIndex        =   22
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BGT(RM)"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VO+(RM)"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VO-(RM)"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Project Information"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4215
         Begin VB.TextBox txt_projdesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   26
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txt_status 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   25
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txt_projcode 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Project Status"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Project Description"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Project Code"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.ComboBox cbo_projcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key - Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
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
         AutoSize        =   -1  'True
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
         Left            =   8200
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   0
         Width           =   2295
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
            Picture         =   "frm_revenue.frx":008C
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":019E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":05F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":0E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":12E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":7580
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":789A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":7BB4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":814E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":86E8
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":8C82
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":921C
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":932E
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":9870
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":9E0A
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":A3A4
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":AC7E
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":AD90
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":AEA2
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":AFB4
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":B0C6
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":B1D8
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":B2EA
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":B884
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":BE1E
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":C3B8
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":C952
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":CA64
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":CB76
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":D110
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":D222
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":D334
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":D8CE
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":D9E0
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":DF7A
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":E514
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":E626
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":EBC0
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":F15A
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":F6F4
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":F806
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":FDA0
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":FEB2
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":FFC4
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":100D6
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":101E8
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":102FA
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":10894
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":109A6
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":10AB8
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":11052
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":115EC
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":11B86
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":12120
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":126BA
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":12C54
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_revenue.frx":131EE
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_revenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer


Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_projcode_Click()
txt_projcode.Text = "-"
txt_projdesc.Text = "-"
txt_status.Text = "-"
pn = Split(cbo_projcode.Text, "-", Len(cbo_projcode.Text), vbTextCompare)

txt_projcode.Text = pn(0)
txt_projdesc.Text = pn(1)
Dim rv1 As New ADODB.Recordset
If rv1.State Then rv1.Close
rv1.Open "select DISTINCT(job_proj_status) from jobcharge where job_proj_key='" & pn(0) & "'", Cn, 3, 2
If Not rv1.EOF Then
txt_status.Text = rv1(0)
End If
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Data is been Loaded......"
Call flex_data
Unload frmBusy

End Sub

Private Sub flex_grid_Click()
On Error Resume Next
'back color
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'------end---------
Unload revenue
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from revenue where rev_id=" & id, Cn, 3, 2
If Not sh.EOF Then
txt_projcode.Text = sh!rev_projcode
 
txt_projdesc.Text = sh!rev_projstatus
revenue.cbo_revtype.Text = sh!rev_type
revenue.cbo_jobno.Text = sh!rev_jobno
 
revenue.cbo_curcy.Text = sh!rev_Currency
revenue.txt_amount.Text = sh!rev_amount
revenue.txt_exchange.Text = sh!rev_exchange
revenue.txt_totalamount.Text = sh!rev_totamount
revenue.txt_notes.Text = sh!rev_tranxnotes
End If
revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775
sh.Close


vprev = flex_grid.Row
X = 1
End Sub

Private Sub flex_grid1_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid1.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid1.Row = vprev
    flex_grid1.Col = 1
    Set flex_grid1.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid1.Cols - 1
    flex_grid1.Col = i
    flex_grid1.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid1.Row = current
For i = 1 To flex_grid1.Cols - 1
flex_grid1.Col = i
flex_grid1.CellBackColor = vbYellow
Next
flex_grid1.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'------end---------
Unload revenue
Dim id As Double
id = 0
If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id = flex_grid1.TextMatrix(flex_grid1.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from revenue where rev_id=" & id, Cn, 3, 2
If Not sh.EOF Then
txt_projcode.Text = sh!rev_projcode
 
txt_projdesc.Text = sh!rev_projstatus
revenue.cbo_revtype.Text = sh!rev_type
revenue.cbo_jobno.Text = sh!rev_jobno
 
revenue.cbo_curcy.Text = sh!rev_Currency
revenue.txt_amount.Text = sh!rev_amount
revenue.txt_exchange.Text = sh!rev_exchange
revenue.txt_totalamount.Text = sh!rev_totamount
revenue.txt_notes.Text = sh!rev_tranxnotes
revenue.txt_perc.Text = sh!perc
revenue.txt_perc.Visible = True
revenue.Label2.Visible = True
End If
revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775
sh.Close


vprev = flex_grid1.Row
X = 2
End Sub

Private Sub flex_grid2_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid2.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid2.Row = vprev
    flex_grid2.Col = 1
    Set flex_grid2.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid2.Cols - 1
    flex_grid2.Col = i
    flex_grid2.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid2.Row = current
For i = 1 To flex_grid2.Cols - 1
flex_grid2.Col = i
flex_grid2.CellBackColor = vbYellow
Next
flex_grid2.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'------end---------
Unload revenue
Dim id As Double
id = 0
If flex_grid2.TextMatrix(flex_grid2.Row, 0) = "" Then Exit Sub
id = flex_grid2.TextMatrix(flex_grid2.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from revenue where rev_id=" & id, Cn, 3, 2
If Not sh.EOF Then
txt_projcode.Text = sh!rev_projcode
 
txt_projdesc.Text = sh!rev_projstatus
revenue.cbo_revtype.Text = sh!rev_type
revenue.cbo_jobno.Text = sh!rev_jobno
 
revenue.cbo_curcy.Text = sh!rev_Currency
revenue.txt_amount.Text = sh!rev_amount
revenue.txt_exchange.Text = sh!rev_exchange
revenue.txt_totalamount.Text = sh!rev_totamount
revenue.txt_notes.Text = sh!rev_tranxnotes
End If
revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775
sh.Close


vprev = flex_grid2.Row
X = 3
End Sub

Private Sub flex_grid3_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid3.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid3.Row = vprev
    flex_grid3.Col = 1
    Set flex_grid3.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid3.Cols - 1
    flex_grid3.Col = i
    flex_grid3.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid3.Row = current
For i = 1 To flex_grid3.Cols - 1
flex_grid3.Col = i
flex_grid3.CellBackColor = vbYellow
Next
flex_grid3.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'------end---------
Unload revenue
Dim id As Double
id = 0
If flex_grid3.TextMatrix(flex_grid3.Row, 0) = "" Then Exit Sub
id = flex_grid3.TextMatrix(flex_grid3.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from revenue where rev_id=" & id, Cn, 3, 2
If Not sh.EOF Then
txt_projcode.Text = sh!rev_projcode
 
txt_projdesc.Text = sh!rev_projstatus
revenue.cbo_revtype.Text = sh!rev_type
revenue.cbo_jobno.Text = sh!rev_jobno
 
revenue.txt_invoice.Text = sh!rev_invoice
revenue.DTP_inv.Value = sh!rev_invoicedate
revenue.cbo_curcy.Text = sh!rev_Currency
revenue.txt_amount.Text = sh!rev_amount
revenue.txt_exchange.Text = sh!rev_exchange
revenue.txt_totalamount.Text = sh!rev_totamount
revenue.txt_notes.Text = sh!rev_tranxnotes

End If
revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775
sh.Close


vprev = flex_grid3.Row
X = 4
End Sub

Private Sub flex_grid4_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid4.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid4.Row = vprev
    flex_grid4.Col = 1
    Set flex_grid4.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid4.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid4.Row = current
For i = 1 To flex_grid4.Cols - 1
flex_grid4.Col = i
flex_grid4.CellBackColor = vbYellow
Next
flex_grid4.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'------end---------
Unload revenue
Dim id As Double
id = 0
If flex_grid4.TextMatrix(flex_grid4.Row, 0) = "" Then Exit Sub
id = flex_grid4.TextMatrix(flex_grid4.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from revenue where rev_type <> 'UBL' and rev_id=" & id, Cn, 3, 2
If Not sh.EOF Then
txt_projcode.Text = sh!rev_projcode
 
txt_projdesc.Text = sh!rev_projstatus
revenue.cbo_revtype.Text = sh!rev_type
revenue.cbo_jobno.Text = sh!rev_jobno
'revenue.txt_invoice.Text = sh!rev_invoice
'revenue.DTP_inv.Value = sh!rev_invoicedate
revenue.cbo_curcy.Text = sh!rev_Currency
revenue.txt_amount.Text = sh!rev_amount
revenue.txt_exchange.Text = sh!rev_exchange
revenue.txt_totalamount.Text = sh!rev_totamount
revenue.txt_notes.Text = sh!rev_tranxnotes
revenue.DTP_tdate.Value = sh!t_date


revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775
End If

sh.Close


vprev = flex_grid4.Row
X = 5
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "REVENUE-BDGT/VO/ADJ/BILLED/UNBILLED"
Call flex_title
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5


Dim rv As New ADODB.Recordset
If rv.State Then rv.Close
rv.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rv.EOF
cbo_projcode.AddItem rv(0) & "  -  " & rv(1)
rv.MoveNext
Wend


If SSTab1.Caption = "Unbilled Revenue Details" Then
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Else
Toolbar1.Buttons(1).Enabled = True
End If
'revenue.txt_perc.Visible = False
'revenue.Label2.Visible = False
 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()

On Error Resume Next
   
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Revn Type"
        .TextMatrix(0, 2) = "Job No"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Curcy"
        .TextMatrix(0, 4) = "Amount"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "XRate"
        .TextMatrix(0, 6) = "Amount(RM)"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 3300
        
    End With
    
    With flex_grid1
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Revn Type"
        .TextMatrix(0, 2) = "Job No"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Curcy"
        .TextMatrix(0, 4) = "Amount"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "XRate"
        .TextMatrix(0, 6) = "Amount(RM)"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 3300
        .TextMatrix(0, 8) = "VO(+)%"
        .ColWidth(8) = 1000
    End With
    
    With flex_grid2
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Revn Type"
        .TextMatrix(0, 2) = "Job No"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Curcy"
        .TextMatrix(0, 4) = "Amount"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "XRate"
        .TextMatrix(0, 6) = "Amount(RM)"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 3300
              
    End With
    With flex_grid3
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Revn Type"
        .TextMatrix(0, 2) = "Job No"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Inv No"
        .TextMatrix(0, 4) = "Inv Date"
        
        .TextMatrix(0, 5) = "Curcy"
        .TextMatrix(0, 6) = "Amount"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "XRate"
        .TextMatrix(0, 8) = "Amount(RM)"
        .ColWidth(8) = 1500
        .TextMatrix(0, 9) = "Notes"
        .ColWidth(9) = 3300
              
    End With
    With flex_grid4
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Revn Type"
        .TextMatrix(0, 2) = "Job No"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Curcy"
        .TextMatrix(0, 4) = "Amount"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "XRate"
        .TextMatrix(0, 6) = "Amount(RM)"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 3300
        
    End With
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload revenue
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "Unbilled Revenue Details" Then
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Else
Toolbar1.Buttons(1).Enabled = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next


If Button.Caption = "New" Then
If cbo_projcode.Text = "" Then
MsgBox "select Project"
cbo_projcode.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload revenue
revenue.Show
revenue.Top = 3200
revenue.Left = 0
revenue.Height = 2295
revenue.Width = 8775 ' to save new record

ElseIf Button.Caption = "Save" Then

            If cbo_projcode.Text = "" Then
            MsgBox "Select Project"
            cbo_projcode.SetFocus
            Exit Sub
            End If
        If revenue.cbo_revtype.Text = "" Then
        MsgBox "Select Revenue Type"
        revenue.cbo_revtype.SetFocus
        Exit Sub
        End If
            If revenue.cbo_curcy.Text = "" Then
            MsgBox "Select Currency"
            revenue.cbo_curcy.SetFocus
            Exit Sub
            End If
        If revenue.txt_amount.Text = "" Then
        MsgBox "Enter Amount"
        revenue.txt_amount.SetFocus
        Exit Sub
        End If
            If revenue.txt_exchange.Text = "" Then
            MsgBox "Enter XRate"
            revenue.txt_exchange.SetFocus
            Exit Sub
            End If
        If revenue.cbo_jobno.Text = "" Then
        MsgBox "Select JobNo."
        revenue.cbo_jobno.SetFocus
        Exit Sub
        End If
 
hh = Split(revenue.cbo_jobno.Text, "  -  ", Len(revenue.cbo_jobno.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from revenue", Cn, 3, 2
sv.AddNew
sv!rev_projcode = txt_projcode.Text
sv!rev_projstatus = txt_status.Text
sv!rev_type = revenue.cbo_revtype.Text
sv!rev_jobno = hh(0)
sv!rev_invoice = revenue.txt_invoice.Text
sv!rev_invoicedate = revenue.DTP_inv.Value
sv!rev_Currency = revenue.cbo_curcy.Text
sv!rev_amount = revenue.txt_amount.Text
sv!rev_exchange = revenue.txt_exchange.Text
sv!rev_totamount = revenue.txt_totalamount.Text
sv!rev_tranxnotes = revenue.txt_notes.Text
sv!perc = revenue.txt_perc.Text
sv!t_date = revenue.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv.Update
sv.Close
MsgBox "New Record Added Succesfully"
Unload revenue
Call flex_data
Call flex_title
'to modify existing record
ElseIf Button.Caption = "Modify" Then
     If cbo_projcode.Text = "" Then
            MsgBox "Select Project"
            cbo_projcode.SetFocus
            Exit Sub
            End If
        If revenue.cbo_revtype.Text = "" Then
        MsgBox "Select Revenue Type"
        revenue.cbo_revtype.SetFocus
        Exit Sub
        End If
            If revenue.cbo_curcy.Text = "" Then
            MsgBox "Select Currency"
            revenue.cbo_curcy.SetFocus
            Exit Sub
            End If
        If revenue.txt_amount.Text = "" Then
        MsgBox "Enter Amount"
        revenue.txt_amount.SetFocus
        Exit Sub
        End If
            If revenue.txt_exchange.Text = "" Then
            MsgBox "Enter XRate"
            revenue.txt_exchange.SetFocus
            Exit Sub
            End If
            If revenue.cbo_jobno.Text = "" Then
            MsgBox "Select JobNo."
            revenue.cbo_jobno.SetFocus
            Exit Sub
            End If
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If X = 1 Then
'If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
End If
If X = 2 Then
'If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id1 = 0
id1 = flex_grid1.TextMatrix(flex_grid1.Row, 0)
End If
If X = 3 Then
id1 = 0
'If flex_grid2.TextMatrix(flex_grid2.Row, 0) = "" Then Exit Sub
id1 = flex_grid2.TextMatrix(flex_grid2.Row, 0)
End If
If X = 4 Then
id1 = 0
'If flex_grid3.TextMatrix(flex_grid3.Row, 0) = "" Then Exit Sub
id1 = flex_grid3.TextMatrix(flex_grid3.Row, 0)
End If
If X = 5 Then
id1 = 0
'If flex_grid3.TextMatrix(flex_grid3.Row, 0) = "" Then Exit Sub
id1 = flex_grid4.TextMatrix(flex_grid4.Row, 0)
End If
hh1 = Split(revenue.cbo_jobno.Text, "  -  ", Len(revenue.cbo_jobno.Text), vbTextCompare)

Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from revenue where rev_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!rev_projcode = txt_projcode.Text
md!rev_projstatus = txt_status.Text
md!rev_type = revenue.cbo_revtype.Text
md!rev_jobno = hh1(0)
md!rev_invoice = revenue.txt_invoice.Text
md!rev_invoicedate = revenue.DTP_inv.Value
md!rev_Currency = revenue.cbo_curcy.Text
md!rev_amount = revenue.txt_amount.Text
md!rev_exchange = revenue.txt_exchange.Text
md!rev_totamount = revenue.txt_totalamount.Text
md!rev_tranxnotes = revenue.txt_notes.Text
md!t_date = revenue.DTP_tdate.Value
md!perc = revenue.txt_perc.Text
md!u_date = Now
md!t_user = main.Label2.Caption
md.Update
md.Close
MsgBox "Record Modified Successfully"
End If

Unload revenue
Call flex_data
Call flex_title

'to delete
ElseIf Button.Caption = "Delete" Then
Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If X = 1 Then
'If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
End If
If X = 2 Then
'If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id2 = 0
id2 = flex_grid1.TextMatrix(flex_grid1.Row, 0)
End If
If X = 3 Then
id2 = 0
'If flex_grid2.TextMatrix(flex_grid2.Row, 0) = "" Then Exit Sub
id2 = flex_grid2.TextMatrix(flex_grid2.Row, 0)
End If
If X = 4 Then
id2 = 0
'If flex_grid3.TextMatrix(flex_grid3.Row, 0) = "" Then Exit Sub
id2 = flex_grid3.TextMatrix(flex_grid3.Row, 0)
End If
If X = 5 Then
id2 = 0
'If flex_grid3.TextMatrix(flex_grid3.Row, 0) = "" Then Exit Sub
id2 = flex_grid4.TextMatrix(flex_grid4.Row, 0)
End If
Cn.Execute "delete from revenue where rev_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload revenue
Call flex_data
Call flex_title
Else
Unload revenue
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload revenue
End If
End Sub
Public Sub flex_data()
On Error Resume Next
Dim bd As Double
Dim co As Double
Dim ad As Double
Dim bl As Double
Dim ubl As Double
Dim bdvo As Double
bd = 0: co = 0: ad = 0: bl = 0: ubl = 0: bdvo = 0

Dim pnh As String
pn = Split(cbo_projcode.Text, "  -  ", Len(cbo_projcode.Text), vbTextCompare)
pnh = Mid(pn(0), 1, 3)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from revenue where rev_type='BGT' and rev_projcode='" & pn(0) & "' order by rev_jobno,rev_Currency ", Cn, 3, 2
With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!rev_type
        Dim flj1 As New ADODB.Recordset
        If flj1.State Then flj1.Close
        flj1.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldata!rev_jobno & "'", Cn, 3, 2
        If Not flj1.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata!rev_jobno & "  -  " & flj1(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata!rev_jobno
        End If
        .TextMatrix(.Rows - 1, 3) = fldata!rev_Currency
        .TextMatrix(.Rows - 1, 4) = Format(fldata!rev_amount, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = Format(fldata!rev_exchange, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata!rev_totamount, "###,###,##0.00")
        bd = bd + fldata!rev_totamount
        .TextMatrix(.Rows - 1, 7) = fldata!rev_tranxnotes
      fldata.MoveNext
    Wend
    txt_bgt.Text = Format(bd, "###,###,##0.00")
End With
Dim fldataa As New ADODB.Recordset
If fldataa.State Then fldataa.Close
fldataa.Open "select * from revenue where rev_type='VO(+)' and rev_projcode='" & pn(0) & "' order by rev_jobno,rev_Currency ", Cn, 3, 2

With flex_grid1
    .Rows = 1
    While Not fldataa.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldataa(0)
        .TextMatrix(.Rows - 1, 1) = fldataa!rev_type
         Dim flj2 As New ADODB.Recordset
        If flj2.State Then flj2.Close
        flj2.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldataa!rev_jobno & "'", Cn, 3, 2
        If Not flj2.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldataa!rev_jobno & "  -  " & flj2(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldataa!rev_jobno
        End If
         
        .TextMatrix(.Rows - 1, 3) = fldataa!rev_Currency
        .TextMatrix(.Rows - 1, 4) = Format(fldataa!rev_amount, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = fldataa!rev_exchange
        .TextMatrix(.Rows - 1, 6) = Format(fldataa!rev_totamount, "###,###,##0.00")
        co = co + fldataa!rev_totamount
        .TextMatrix(.Rows - 1, 7) = fldataa!rev_tranxnotes
        .TextMatrix(.Rows - 1, 8) = fldataa!perc
      fldataa.MoveNext
    Wend
    txt_vos.Text = Format(co, "###,###,##0.00")
End With
Dim fldatab As New ADODB.Recordset
If fldatab.State Then fldatab.Close
fldatab.Open "select * from revenue where rev_type='VO(-)' and rev_projcode='" & pn(0) & "' order by rev_jobno,rev_Currency", Cn, 3, 2

With flex_grid2
    .Rows = 1
    While Not fldatab.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldatab(0)
        .TextMatrix(.Rows - 1, 1) = fldatab!rev_type
         Dim flj3 As New ADODB.Recordset
        If flj3.State Then flj3.Close
        flj3.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldatab!rev_jobno & "'", Cn, 3, 2
        If Not flj3.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldatab!rev_jobno & "  -  " & flj3(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldatab!rev_jobno
        End If
         
        .TextMatrix(.Rows - 1, 3) = fldatab!rev_Currency
        .TextMatrix(.Rows - 1, 4) = Format(fldatab!rev_amount, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = fldatab!rev_exchange
        .TextMatrix(.Rows - 1, 6) = Format(fldatab!rev_totamount, "###,###,##0.00")
        ad = ad + fldatab!rev_totamount
        .TextMatrix(.Rows - 1, 7) = fldatab!rev_tranxnotes
      fldatab.MoveNext
    Wend
    txt_adj.Text = Format(ad, "###,###,##0.00")
End With
Dim fldatac As New ADODB.Recordset
If fldatac.State Then fldatac.Close
fldatac.Open "select * from revenue where rev_type='BLD' and rev_projcode='" & pn(0) & "' order by rev_jobno ,rev_Currency", Cn, 3, 2

With flex_grid3
    .Rows = 1
    While Not fldatac.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldatac(0)
        .TextMatrix(.Rows - 1, 1) = fldatac!rev_type
         Dim flj4 As New ADODB.Recordset
        If flj4.State Then flj4.Close
        flj4.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldatac!rev_jobno & "'", Cn, 3, 2
        If Not flj4.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldatac!rev_jobno & "  -  " & flj4(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldatac!rev_jobno
        End If
        
        .TextMatrix(.Rows - 1, 3) = fldatac!rev_invoice
        .TextMatrix(.Rows - 1, 4) = fldatac!rev_invoicedate
        .TextMatrix(.Rows - 1, 5) = fldatac!rev_Currency
        .TextMatrix(.Rows - 1, 6) = Format(fldatac!rev_amount, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = fldatac!rev_exchange
        .TextMatrix(.Rows - 1, 8) = Format(fldatac!rev_totamount, "###,###,##0.00")
        bl = bl + fldatac!rev_totamount
        .TextMatrix(.Rows - 1, 9) = fldatac!rev_tranxnotes
      fldatac.MoveNext
    Wend
   txt_bld.Text = Format(bl, "###,###,##0.00")
End With

'bgt vo
Dim fldatabg As New ADODB.Recordset
If fldatabg.State Then fldatabg.Close
fldatabg.Open "select * from revenue where rev_type='BGT VO' and rev_projcode='" & pn(0) & "' order by rev_jobno,rev_Currency ", Cn, 3, 2
With flex_grid4
    .Rows = 1
    While Not fldatabg.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldatabg(0)
        .TextMatrix(.Rows - 1, 1) = fldatabg!rev_type
        Dim fljbg As New ADODB.Recordset
        If fljbg.State Then fljbg.Close
        fljbg.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldatabg!rev_jobno & "'", Cn, 3, 2
        If Not fljbg.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldatabg!rev_jobno & "  -  " & flj1(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldatabg!rev_jobno
        End If
        .TextMatrix(.Rows - 1, 3) = fldatabg!rev_Currency
        .TextMatrix(.Rows - 1, 4) = Format(fldatabg!rev_amount, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = Format(fldatabg!rev_exchange, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldatabg!rev_totamount, "###,###,##0.00")
        bdvo = bdvo + fldatabg!rev_totamount
        .TextMatrix(.Rows - 1, 7) = fldatabg!rev_tranxnotes
      fldatabg.MoveNext
    Wend
    txt_bvo.Text = Format(CDbl(bdvo), "###,###,##0.00")
End With





'-----------





Dim b1 As Double
Dim b2 As Double

b1 = 0: b2 = 0
Cn.Execute "delete from revenue where rev_type='UBL' and rev_projcode='" & pn(0) & "'"
Dim jnc As New ADODB.Recordset
If jnc.State Then jnc.Close
jnc.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & pn(0) & "' order by jobno_code", Cn, 3, 2
While Not jnc.EOF
 
Dim a1 As Double
Dim a2 As Double
Dim a3 As Double
a1 = 0: a2 = 0: a3 = 0


Dim b1g As Double
Dim b2g As Double
b1g = 0: b2g = 0

Dim jng As New ADODB.Recordset
If jng.State Then jng.Close
jng.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j,cost c where j.job_code=c.bd_jobcharge and j.jobno='" & jnc(0) & "' and j.job_proj_key = '" & pn(0) & "' and c.bd_costtype='E'", Cn, 3, 2
If Not jng.EOF Then
b1g = b1g + jng(0)
b2g = b2g + jng(1)
End If
b1 = b1 + b1g
b2 = b2 + b2g



jnc.MoveNext
Wend
Dim rvn As New ADODB.Recordset
If rvn.State Then rvn.Close
rvn.Open "select * from revenue  ", Cn, 3, 2
rvn.AddNew
rvn!rev_projcode = pn(0)
rvn!rev_type = "UBL"
rvn!rev_totamount = Format((b1 / (b1 + b2)) * (bd + co + ad), "###,###,###,##0.00")
rvn!rev_jobno = "-"
rvn.Update
'''Dim rmj As New ADODB.Recordset
'''If rmj.State Then rmj.Close
'''rmj.Open "select SUM(rev_totamount) from revenue where rev_type='UBL' and rev_projcode='" & pn(0) & "' ", Cn, 3, 2
''' If Not rmj.EOF Then
txt_ubl.Text = ""
txt_ubl.Text = Format((b1 / (b1 + b2)) * (bd + co + ad), "###,###,###,##0.00")
'''End If
Text1.Text = ""
Text1.Text = Format(txt_ubl.Text - bl, "###,###,##0.00")
End Sub


Private Sub txt_bld_Change()
On Error Resume Next
Text1.Text = Format(txt_ubl.Text - txt_bld.Text, "###,###,##0.00")
End Sub

Private Sub txt_ubl_Change()
On Error Resume Next
Text1.Text = Format(txt_ubl.Text - txt_bld.Text, "###,###,##0.00")
End Sub
