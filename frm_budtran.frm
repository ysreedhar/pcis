VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_budtran 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Generate EIC from BC Transactions"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   41
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8295
      Left            =   0
      TabIndex        =   40
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   14631
      _Version        =   393216
      Rows            =   3
      Cols            =   22
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin VB.ComboBox cbo_resc 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6720
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox cbo_year 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_pproj 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   16
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":070E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":0E1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":152A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":1C38
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":230E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":27F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":2EFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":360C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":3D1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":4428
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":4B36
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":5018
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":53B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":5AC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_budtran.frx":61D2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Project"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   120
         Width           =   2535
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Jobcharge"
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
         Left            =   6720
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16761024
      TabCaption(0)   =   "Budgeted Cost"
      TabPicture(0)   =   "frm_budtran.frx":68E0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   11175
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Transaction Information"
            Enabled         =   0   'False
            Height          =   735
            Left            =   4920
            TabIndex        =   34
            Top             =   960
            Visible         =   0   'False
            Width           =   6135
            Begin VB.TextBox txt_respcode 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox txt_respname 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Resc  Resp Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Resc Resp Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   480
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Project Information"
            Enabled         =   0   'False
            Height          =   855
            Left            =   4920
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   6135
            Begin VB.TextBox txt_projdesc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   5400
               TabIndex        =   29
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox textprojkey 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   5160
               TabIndex        =   28
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   285
               Left            =   720
               TabIndex        =   27
               Text            =   "BCWP- RM"
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   285
               Left            =   720
               TabIndex        =   26
               Text            =   "BDGT- RM"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox Txt_gtotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   1800
               TabIndex        =   25
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_btotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   1800
               TabIndex        =   24
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox textcosttype 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   4560
               TabIndex        =   23
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Project Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   33
               Top             =   480
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Project Key"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   32
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cost Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3720
               TabIndex        =   30
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resource Information"
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   4695
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   1440
               TabIndex        =   16
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   1440
               TabIndex        =   15
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox textresccode 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1920
               TabIndex        =   14
               Top             =   240
               Width           =   2655
            End
            Begin VB.TextBox textrescname 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1920
               TabIndex        =   13
               Top             =   480
               Width           =   2655
            End
            Begin VB.TextBox txt_brate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1920
               TabIndex        =   12
               Top             =   720
               Width           =   2655
            End
            Begin VB.TextBox txt_crate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1920
               TabIndex        =   11
               Top             =   960
               Width           =   2655
            End
            Begin VB.TextBox txt_vendor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Left            =   1920
               TabIndex        =   10
               Top             =   1200
               Width           =   2655
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Vendor Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   1200
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Rate(Current)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Rate(Budget)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Resource Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Resource Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   840
               Width           =   1695
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   635
      ButtonWidth     =   6906
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Generate EIC Transactions from BC Transactions"
            Key             =   "ar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComCtl2.DTPicker dtpdefault 
         Height          =   375
         Left            =   9000
         TabIndex        =   42
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy H:mm:ss"
         Format          =   48889859
         CurrentDate     =   38140
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
            Picture         =   "frm_budtran.frx":68FC
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":6A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":6E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":72B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":7704
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":7B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":DDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":E10A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":E424
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":E9BE
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":EF58
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":F4F2
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":FA8C
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":FB9E
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":100E0
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":1067A
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":10C14
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":114EE
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11600
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11712
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11824
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11936
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11A48
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":11B5A
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":120F4
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":1268E
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":12C28
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":131C2
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":132D4
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":133E6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":13980
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":13A92
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":13BA4
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":1413E
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":14250
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":147EA
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":14D84
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":14E96
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":15430
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":159CA
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":15F64
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16076
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16610
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16722
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16834
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16946
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16A58
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":16B6A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":17104
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":17216
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":17328
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":178C2
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":17E5C
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":183F6
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":18990
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":18F2A
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":194C4
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budtran.frx":19A5E
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_budtran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As Double
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_pproj_Change()
 
cbo_resc.Clear
 
 

gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
 
        
   Dim rs As New ADODB.Recordset
        rs.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & gg(0) & "' order by job_code", Cn, 3, 2
        While Not rs.EOF
        cbo_resc.AddItem rs(0) & "  -  " & rs(1)
        rs.MoveNext
        Wend
        rs.Close
        
        
kl1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
If cbo_resc.Text = "" Then
'MsgBox "Select Resource Code"
cbo_resc.SetFocus
Exit Sub
End If
 
End Sub

Private Sub cbo_pproj_Click()
 
cbo_resc.Clear
 

gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
 
        
        Dim rs As New ADODB.Recordset
        rs.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & gg(0) & "' order by job_code", Cn, 3, 2
        While Not rs.EOF
        cbo_resc.AddItem rs(0) & "  -  " & rs(1)
        rs.MoveNext
        Wend
        rs.Close
        
kl1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
If cbo_resc.Text = "" Then
'MsgBox "Select Resource Code"
cbo_resc.SetFocus
Exit Sub
End If

 
 
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
 

Call flex_data1

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
gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)


Call flex_data1

End Sub

Private Sub cbo_resc_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_year_Change()

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
 
 



'''
End Sub

Private Sub cbo_year_Click()

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

Private Sub cbo_year_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub
Private Sub flex_grid_Click()
On Error Resume Next
 
Static vprev As Integer
current = flex_grid.Row
 
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
' back color
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


Unload budgetedcoste
Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
budgetedcoste.Show
 
  Dim rsd As New ADODB.Recordset
  If rsd.State Then rsd.Close
  rsd.Open "select * from cost where bd_id=" & ID, Cn, 3, 2
  If Not rsd.EOF Then
        cbo_year.Text = rsd!bd_year
        textresccode.Text = rsd!bd_resccode
        textrescname.Text = rsd!bd_rescname
        txt_vendor.Text = rsd!bd_vendor
        textprojkey.Text = rsd!bd_projectkey
        txt_projdesc.Text = rsd!bd_projectdesc
        textcosttype.Text = rsd!bd_costtype
        txt_respcode.Text = rsd!bd_respcode

        budgetedcoste.txt_qty.Text = rsd!bd_qty
        budgetedcoste.txt_days.Text = rsd!bd_days
        budgetedcoste.txt_totdays.Text = rsd!bd_tqty
        budgetedcoste.txt_unitrate.Text = Format(rsd!bd_unitrate, "###,###,##0.00")
        budgetedcoste.txt_Xrate.Text = rsd!bd_xchg
        budgetedcoste.txt_downtime.Text = rsd!bd_downtime
        budgetedcoste.txt_esclfactor.Text = rsd!bd_escl
        budgetedcoste.txt_Extdamt.Text = Format(rsd!bd_extdamt, "###,###,##0.00")
        budgetedcoste.txt_wrkcomp.Text = rsd!bd_wrkcomp
        budgetedcoste.txt_bcwpamt.Text = Format(rsd!bd_bcwpamt, "###,###,##0.00")


        Dim rr1 As New ADODB.Recordset
        If rr1.State Then rr1.Close
        rr1.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & rsd!bd_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        txt_respname.Text = rr(0)
        End If

        txt_brate.Text = Format(rsd!bd_brate, "###,###,##0.00")
        txt_crate.Text = Format(rsd!bd_crate, "###,###,##0.00")


 End If
        budgetedcoste.cbo_spread.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
        budgetedcoste.cbo_tranx.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
        budgetedcoste.cbo_jobcharge.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
        budgetedcoste.cbo_obs.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
        budgetedcoste.cbo_costcode.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
        
        budgetedcoste.cbo_uom.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
        budgetedcoste.cbo_curr.Text = flex_grid.TextMatrix(flex_grid.Row, 11)
        budgetedcoste.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 20)
        budgetedcoste.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 19)

budgetedcoste.Show
budgetedcoste.Top = 3200
budgetedcoste.Left = 0
budgetedcoste.Height = 3540
budgetedcoste.Width = 8850
rsd.Close



vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "GENERATE EIC TRANSACTIONS"
Call flex_title
'Call flex_data1

Me.Top = 5
Me.Left = 5
DTP_cod.Value = Format(Date, "dd/mm/yyyy")
dtpdefault.Value = Format(Date, "dd/MM/yyyy H:mm:ss")
Dim i As Integer
i = 0
For i = 2000 To 2050
cbo_year.AddItem i
Next i
 Me.Width = 11415
 Me.Height = 9750

 
End Sub
Public Sub flex_title()
On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 300
        .TextMatrix(0, 2) = "TranX"
        .ColWidth(2) = 0
        .ColAlignment(2) = Left
        .TextMatrix(0, 3) = "Spread"
        .ColWidth(3) = 1100
        .ColAlignment(3) = 0
         
        .TextMatrix(0, 4) = "Resource"
        .ColWidth(4) = 3500
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "OBS"
        .ColWidth(5) = 0
        .TextMatrix(0, 6) = "Costcode"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Qty"
        .ColWidth(7) = 520
        
        .TextMatrix(0, 8) = "Days"
        .ColWidth(8) = 620
        
        .TextMatrix(0, 9) = "TotalQty"
        .ColWidth(9) = 900
        
        .TextMatrix(0, 10) = "UOM"
        .ColWidth(10) = 600
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "Curcy"
        .ColWidth(11) = 600
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "Xrate"
        .ColWidth(12) = 600
        
        .TextMatrix(0, 13) = "UnitRate"
        .ColWidth(13) = 1000
        
        .TextMatrix(0, 14) = "D/Time %"
        .ColWidth(14) = 600
        
        .TextMatrix(0, 15) = "Escl %"
        .ColWidth(15) = 600
        
        .TextMatrix(0, 16) = "BDGT(RM)"
        .ColWidth(16) = 1100
        
        .TextMatrix(0, 17) = "% Wrk Compltd"
        .ColWidth(17) = 600
        .TextMatrix(0, 18) = "BCWP(RM)"
        .ColWidth(18) = 1100
        
        .TextMatrix(0, 19) = "Notes"
        .ColWidth(19) = 4000
        .ColAlignment(19) = 0
        .ColWidth(20) = 0
        
        .TextMatrix(0, 21) = "Cg.Type"
        .ColWidth(21) = 700
        
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
 
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
 
Dim uid As Double
uid = 0
Dim l As Double

' to save new record
 If Button.Caption = "Generate EIC Transactions from BC Transactions" Then
  
Y = MsgBox("Have you selected the Default date", vbYesNo)
If Y = vbYes Then
 
Dim f As Double
Dim dh As Double
dh = 0
f = 0
f = flex_grid.Rows

For i = 0 To f - 1
If Check2(i).Value = vbChecked Then
 '*************************************
  
ng = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
nh = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
ab = Split(flex_grid.TextMatrix(i, 4), "  -  ", Len(flex_grid.TextMatrix(i, 4)), vbTextCompare)
ac = Split(flex_grid.TextMatrix(i, 3), "  -  ", Len(flex_grid.TextMatrix(i, 3)), vbTextCompare)
ad = Split(flex_grid.TextMatrix(i, 6), "  -  ", Len(flex_grid.TextMatrix(i, 6)), vbTextCompare)


 
Dim btra As New ADODB.Recordset
If btra.State Then btra.Close
btra.Open "select * from cost ", Cn, 3, 2

btra.AddNew
        btra!estid = flex_grid.TextMatrix(i, 0)
        btra!bd_year = cbo_year.Text
        btra!bd_resccode = ab(0)
        btra!bd_rescname = ab(1)
        Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & ab(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & nh(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
            btra!bd_brate = Format(fl!dresc_rate, "###,###,##0.00")
            btra!bd_vendor = fl!resc_vendorcode
            btra!bd_respcode = fl!resc_respcode
            Dim rr As New ADODB.Recordset
            If rr.State Then rr.Close
            rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
            If Not rr.EOF Then
            btra!bd_respname = rr(0)
            End If
            btra!bd_crate = 0
            End If
fl.Close
         '''- (((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 14))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))) + ((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))))+ ((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13)))
        btra!bd_projectkey = nh(0)
        btra!bd_projectdesc = nh(1)
        btra!bd_costtype = "E"
        btra!bd_cuttdate = main.DTPcutdate1.Value
        
        btra!bd_spread = ac(0)
        btra!bd_tranx = flex_grid.TextMatrix(i, 2)
        btra!bd_JobCharge = ng(0)
        btra!bd_costcode = ad(0)
        btra!bd_qty = flex_grid.TextMatrix(i, 7)
        btra!bd_days = flex_grid.TextMatrix(i, 8)
        
        btra!bd_uom = flex_grid.TextMatrix(i, 10)
        btra!bd_curr = flex_grid.TextMatrix(i, 11)
        btra!bd_unitrate = (((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))))
        btra!bd_xchg = flex_grid.TextMatrix(i, 12)
        btra!bd_downtime = flex_grid.TextMatrix(i, 14)
        btra!bd_escl = flex_grid.TextMatrix(i, 15)
        btra!bd_extdamt = CDbl(flex_grid.TextMatrix(i, 12)) * CDbl((((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))))) * CDbl(flex_grid.TextMatrix(i, 9))
        btra!bd_wrkcomp = flex_grid.TextMatrix(i, 17)
        btra!bd_bcwpamt = flex_grid.TextMatrix(i, 18)
        
        If ac(0) = "NA" Then
        If flex_grid.TextMatrix(i, 8) = "" Then
        btra!bd_tqty = flex_grid.TextMatrix(i, 7)
        btra!bd_chk = 0
        End If
        If IsNull(flex_grid.TextMatrix(i, 8)) Then
        btra!bd_tqty = flex_grid.TextMatrix(i, 7)
        btra!bd_chk = 0
        End If
        If flex_grid.TextMatrix(i, 8) >= 1 Then
        btra!bd_tqty = flex_grid.TextMatrix(i, 9)
        btra!bd_chk = 1
        End If
        
                
        
        
        btra!bd_sdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
        If CDbl(flex_grid.TextMatrix(i, 8)) = 0 Then
        btra!bd_edate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
        Else
        btra!bd_edate = Format(DateAdd("d", CDbl(flex_grid.TextMatrix(i, 8)), dtpdefault.Value), "dd/MM/yyyy H:mm:ss")
        End If
        btra!bd_type = "-"
        Else
        btra!bd_sdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
        btra!bd_edate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
        btra!bd_type = "A"
        End If
        btra!bd_inv = "-"
        btra!bd_invdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
        
        
        btra!bd_notes = "-"
        
        
        Dim dtt As Double
        Dim dtt1 As Double
        dtt = 0: dtt1 = 0
        dtt = btra!bd_sdate
        dtt1 = (CDbl(btra!bd_days) + CDbl(btra!bd_e_days))
        If btra!bd_chk = 1 Then
        btra!bd_edate = Format(dtt + dtt1, "dd/MM/yyyy H:mm:ss")
        End If
        
        
        
        btra!t_date = flex_grid.TextMatrix(i, 20)
        btra!u_date = Now
        btra!t_user = main.Label2.Caption
        btra!bd_obs = flex_grid.TextMatrix(i, 5)
        btra!bd_idd = flex_grid.TextMatrix(i, 0)
        btra!bd_ChargeType = flex_grid.TextMatrix(i, 21)
        Cn.Execute "update cost set estid= '" & flex_grid.TextMatrix(i, 0) & "' where bd_costtype='B' and bd_id= '" & flex_grid.TextMatrix(i, 0) & "'"
btra.Update
End If
'**************************



Next
 MsgBox "Generated Successfully"
Call flex_data1

Else
dtpdefault.SetFocus
End If
'to modify existing record
ElseIf Button.Caption = "Close" Then
Unload Me
Unload budgetedcoste
End If

End Sub

Public Sub flex_data1()
On Error Resume Next
 
Dim ji As Integer
ji = 0
With flex_grid
For ji = 1 To kk - 1
Check2(ji).Visible = False
Next
End With


rscc = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
jk = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim btotal As Double
btotal = 0

Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where  bd_jobcharge='" & rscc(0) & "'  and bd_projectkey='" & jk(0) & "' and bd_year= '" & cbo_year.Text & "' and   estid <  1 and  bd_costtype='B'  order by bd_costcode,bd_spread ,bd_jobcharge,bd_resccode", Cn, 3, 2
With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        
        'loading checkbox
        On Error Resume Next
         Load Check2(.Rows - 1)
        .Col = 1
        .Row = .Rows - 1
        Check2(.Rows - 1).Left = .Left + .CellLeft
        Check2(.Rows - 1).Top = .Top + .CellTop
        Check2(.Rows - 1).Height = .CellHeight
        Check2(.Rows - 1).Width = .CellWidth
        Check2(.Rows - 1).ZOrder 0
        Check2(.Rows - 1).Visible = True
        Check2(.Rows - 1).Value = 1
        .TextMatrix(.Rows - 1, 2) = fldata!bd_tranx
        Dim spd As New ADODB.Recordset
        If spd.State Then spd.Close
        spd.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spd.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_spread & "  -  " & spd(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_spread
        End If
        spd.Close
        Dim ki As New ADODB.Recordset
If ki.State Then ki.Close
ki.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & fldata!bd_resccode & "' ", Cn, 3, 2
If Not ki.EOF Then
.TextMatrix(.Rows - 1, 4) = fldata!bd_resccode & "  -  " & ki(0)
Else
.TextMatrix(.Rows - 1, 4) = fldata!bd_resccode
End If
        .TextMatrix(.Rows - 1, 5) = fldata!bd_obs
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 6) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 6) = fldata!bd_costcode
        End If
        cs.Close
        If IsNull(fldata!bd_qty) Then
        .TextMatrix(.Rows - 1, 7) = ""
        Else
        .TextMatrix(.Rows - 1, 7) = fldata!bd_qty
        End If
        If IsNull(fldata!bd_days) Then
        .TextMatrix(.Rows - 1, 8) = ""
        Else
        .TextMatrix(.Rows - 1, 8) = fldata!bd_days
        End If
        .TextMatrix(.Rows - 1, 9) = fldata!bd_tqty
        .TextMatrix(.Rows - 1, 10) = fldata!bd_uom
        .TextMatrix(.Rows - 1, 11) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 12) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 13) = fldata!bd_unitrate
        .TextMatrix(.Rows - 1, 14) = fldata!bd_downtime
        .TextMatrix(.Rows - 1, 15) = fldata!bd_escl
        .TextMatrix(.Rows - 1, 16) = Format(fldata!bd_extdamt, "###,###,###,###,##0.00")
         gtotal = gtotal + fldata!bd_extdamt
        .TextMatrix(.Rows - 1, 17) = Format(fldata!bd_wrkcomp, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 18) = Format(fldata!bd_bcwpamt, "###,###,###,###,##0.00")
        .TextMatrix(.Rows - 1, 19) = fldata!bd_notes
        .TextMatrix(.Rows - 1, 20) = fldata!t_date
        .TextMatrix(.Rows - 1, 21) = fldata!bd_ChargeType
         btotal = btotal + fldata!bd_bcwpamt
        
        fldata.MoveNext
        Wend
        kk = 0
        kk = flex_grid.Rows
End With
Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
txt_btotal.Text = Format(btotal, "###,###,##0.00")
End Sub
 


