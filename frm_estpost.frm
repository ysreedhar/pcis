VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_estpost 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EIC Post"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11535
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Rectify"
         Height          =   375
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cbo_pproj 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   5295
      End
      Begin VB.ComboBox cbo_year 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_resc 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6960
         TabIndex        =   2
         Top             =   360
         Width           =   4335
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
         Left            =   6960
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   120
         Width           =   855
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
         TabIndex        =   5
         Top             =   120
         Width           =   2535
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
               Picture         =   "frm_estpost.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":070E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":0E1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":152A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":1C38
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":230E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":27F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":2EFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":360C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":3D1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":4428
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":4B36
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":5018
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":53B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":5AC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_estpost.frx":61D2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14631
      _Version        =   393216
      Rows            =   3
      Cols            =   21
      RowHeightMin    =   300
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   2055
      Left            =   0
      TabIndex        =   8
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
      TabPicture(0)   =   "frm_estpost.frx":68E0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   11175
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resource Information"
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Visible         =   0   'False
            Width           =   4695
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
               TabIndex        =   34
               Top             =   1200
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
               TabIndex        =   33
               Top             =   960
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
               TabIndex        =   32
               Top             =   720
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
               TabIndex        =   31
               Top             =   480
               Width           =   2655
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
               TabIndex        =   30
               Top             =   240
               Width           =   2655
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
               TabIndex        =   29
               Top             =   960
               Width           =   375
            End
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
               TabIndex        =   28
               Top             =   720
               Width           =   375
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
               TabIndex        =   39
               Top             =   240
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
               TabIndex        =   38
               Top             =   480
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
               TabIndex        =   37
               Top             =   720
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
               TabIndex        =   36
               Top             =   960
               Width           =   1695
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
               TabIndex        =   35
               Top             =   1200
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Project Information"
            Enabled         =   0   'False
            Height          =   855
            Left            =   4920
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   6135
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
               TabIndex        =   22
               Top             =   240
               Width           =   735
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
               TabIndex        =   21
               Top             =   480
               Width           =   1455
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
               TabIndex        =   20
               Top             =   240
               Width           =   1455
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
               TabIndex        =   19
               Text            =   "BDGT- RM"
               Top             =   240
               Width           =   735
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
               TabIndex        =   18
               Text            =   "BCWP- RM"
               Top             =   480
               Width           =   855
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
               TabIndex        =   17
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
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
               TabIndex        =   16
               Top             =   240
               Visible         =   0   'False
               Width           =   615
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
               TabIndex        =   26
               Top             =   240
               Width           =   735
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
               TabIndex        =   25
               Top             =   240
               Width           =   375
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
               TabIndex        =   24
               Top             =   240
               Visible         =   0   'False
               Width           =   975
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
               TabIndex        =   23
               Top             =   480
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Transaction Information"
            Enabled         =   0   'False
            Height          =   735
            Left            =   4920
            TabIndex        =   10
            Top             =   960
            Visible         =   0   'False
            Width           =   6135
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
               TabIndex        =   12
               Top             =   480
               Width           =   2415
            End
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
               TabIndex        =   11
               Top             =   240
               Width           =   2415
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
               TabIndex        =   14
               Top             =   480
               Width           =   1455
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
               TabIndex        =   13
               Top             =   240
               Width           =   1455
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   635
      ButtonWidth     =   2540
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit / Post EIC"
            Key             =   "ar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComCtl2.DTPicker dtpdefault 
         Height          =   375
         Left            =   3720
         TabIndex        =   41
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy H:mm:ss"
         Format          =   67174403
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
            Picture         =   "frm_estpost.frx":68FC
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":6A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":6E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":72B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":7704
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":7B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":DDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":E10A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":E424
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":E9BE
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":EF58
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":F4F2
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":FA8C
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":FB9E
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":100E0
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":1067A
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":10C14
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":114EE
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11600
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11712
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11824
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11936
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11A48
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":11B5A
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":120F4
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":1268E
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":12C28
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":131C2
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":132D4
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":133E6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":13980
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":13A92
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":13BA4
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":1413E
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":14250
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":147EA
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":14D84
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":14E96
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":15430
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":159CA
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":15F64
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16076
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16610
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16722
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16834
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16946
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16A58
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":16B6A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":17104
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":17216
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":17328
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":178C2
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":17E5C
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":183F6
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":18990
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":18F2A
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":194C4
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_estpost.frx":19A5E
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_estpost"
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
 
On Error Resume Next
 
cbo_resc.Clear
 
 

gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
 
        
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
 
cbo_resc.Clear
 

gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
 
        
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
 Call flex_data1

End Sub

Private Sub cbo_resc_Click()
On Error Resume Next
  
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

Private Sub Command1_Click()
'kl = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim rsn As String
Dim rsc As String
Dim pk As String
Dim pd As String
Dim vn As String
Dim rsp As String
Dim lj As String
lj = Mid(gg(0), 1, 3)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from cost where bd_costtype='E' and bd_jobcharge like '" & lj & "%' and bd_year='" & cbo_year.Text & "' ", Cn, 3, 2
While Not rs.EOF


Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code  and resc_code='" & rs!bd_resccode & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        rsn = fl!resc_desc
        rsc = fl!resc_code
        pk = gg(0)
        pd = gg(1)
        vn = fl!resc_vendorcode
        rsp = fl!resc_respcode
        End If

 rs!bd_rescname = rsn
 rs!bd_resccode = rsc
 rs!bd_projectkey = pk
 rs!bd_projectdesc = pd
 rs!bd_vendor = vn
 rs!bd_respcode = rsp
 rs.Update
rs.MoveNext
Wend

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
Unload esttran
esttran.Top = 3200
esttran.Left = 0
esttran.Height = 3915
esttran.Width = 9645
esttran.Show

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
 
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)
'esttran.Show

Dim rsdd  As New ADODB.Recordset
If rsdd.State Then rsd.Close
rsdd.Open "select * from cost1  where bd_id=" & id, Cn, 3, 2
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
        esttran.cbo_spread.Text = rsdd!bd_spread & "  -  " & spd1(0)
        Else
        esttran.cbo_spread.Text = rsdd!bd_spread
        End If
        spd1.Close
        
        esttran.cbo_tranx.Text = rsdd!bd_tranx
      
        
        Dim fl1 As New ADODB.Recordset
            If fl1.State Then fl1.Close
            fl1.Open "select DISTINCT(resc_desc) from resourcemaster rm, resourcedetails rd where  resc_code='" & rsdd!bd_resccode & "'  ", Cn, 3, 2
            If Not fl1.EOF Then
             esttran.cbo_jobcharge.Text = rsdd!bd_resccode & "  -  " & fl1(0)
            End If
         Dim cs1 As New ADODB.Recordset
        If cs1.State Then cs1.Close
        cs1.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & rsdd!bd_costcode & "' ", Cn, 3, 2
        If Not cs1.EOF Then
        esttran.cbo_costcode.Text = rsdd!bd_costcode & "  -  " & cs1(0)
        Else
        esttran.cbo_costcode.Text = rsdd!bd_costcode
        End If
        cs1.Close
         
        esttran.txt_qty.Text = rsdd!bd_qty
        esttran.txt_days.Text = rsdd!bd_days
        esttran.txt_tqty.Text = rsdd!bd_tqty
        esttran.cbo_uom.Text = rsdd!bd_uom
        esttran.cbo_curr.Text = rsdd!bd_curr
        esttran.txt_Xrate.Text = rsdd!bd_xchg
        esttran.txt_unitrate.Text = Format(rsdd!bd_unitrate, "###,###,##0.00")
        esttran.txt_Extdamt.Text = Format(rsdd!bd_extdamt, "###,###,##0.00")
        esttran.txt_note.Text = rsdd!bd_notes
        esttran.cbo_obs.Text = rsdd!bd_obs
                                If IsNull(rsdd!bd_e_days) = True Then
                                esttran.txt_edays.Text = ""
                                Else
                                esttran.txt_edays.Text = rsdd!bd_e_days
                                End If
        esttran.txt_etqty.Text = rsdd!bd_e_tqty
        esttran.txt_ectcamt.Text = Format(rsdd!bd_e_extdamt, "###,###,##0.00")
                        If IsNull(rsdd!bd_sdate) = False Then
                        esttran.DTP_ed.Value = flex_grid.TextMatrix(flex_grid.Row, 15)
                        Else
                         
                        esttran.DTP_ed.Value = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
                        End If
                    If IsNull(rsdd!bd_edate) = False Then
                    esttran.DTP_sd.Value = flex_grid.TextMatrix(flex_grid.Row, 14)
                    Else
                    esttran.DTP_sd.Value = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
                    End If
        
                If rsdd!bd_days >= 1 Then
                esttran.Check1.Value = 1
                ElseIf rsdd!bd_e_days >= 1 Then
                esttran.Check1.Value = 1
                Else
                esttran.Check1.Value = 0
                End If
                
               esttran.cbo_type.Text = rsdd!bd_type

End If
If esttran.cbo_spread.Text = "NA  -  Not Applicable" Then
                If esttran.Check1.Value = 0 Then
                esttran.DTP_ed.Enabled = 0
                Else
                        esttran.DTP_sd.Enabled = True
                        esttran.DTP_ed.Enabled = True
                        esttran.Check1.Visible = True
                        esttran.lbl.Visible = True
                End If
                esttran.cbo_type.Text = "-"
Else
esttran.DTP_sd.Enabled = False
esttran.DTP_ed.Enabled = False
esttran.Check1.Visible = False
esttran.lbl.Visible = False
esttran.cbo_type.Text = "A"
End If



vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "EDIT/POST TRANSACTIONS"
Call flex_title
Call flex_data1
dtpdefault.Value = Format(Date, "dd/MM/yyyy H:mm:ss")
Me.Top = 5
Me.Left = 5
DTP_cod.Value = Format(Date, "dd/mm/yyyy")
Dim i As Integer
i = 0
For i = 2000 To 2050
cbo_year.AddItem i
Next i


Dim kj As String
kj = MsgBox("Select Default date", vbOK)

dtpdefault.SetFocus

 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()
On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Select"
        .ColWidth(1) = 0
        .TextMatrix(0, 2) = "TranX"
        .ColWidth(2) = 600
        .ColAlignment(2) = Left
        .TextMatrix(0, 3) = "Spread"
        .ColWidth(3) = 2500
        .ColAlignment(3) = 0
         
        .TextMatrix(0, 4) = "Resource"
        .ColWidth(4) = 3500
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "OBS"
        .ColWidth(5) = 0
        .TextMatrix(0, 6) = "Costcode"
        .ColWidth(6) = 3000
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
        
        .TextMatrix(0, 14) = "SDate"
        .ColWidth(14) = 2000
        
        .TextMatrix(0, 15) = "EDate"
        .ColWidth(15) = 2000
        
        .TextMatrix(0, 16) = "ACWP"
        .ColWidth(16) = 1100
        
        .TextMatrix(0, 17) = "%WC"
        .ColWidth(17) = 0
        .TextMatrix(0, 18) = "ECTC"
        .ColWidth(18) = 1100
        
        .TextMatrix(0, 19) = "Notes"
        .ColWidth(19) = 4000
        .ColAlignment(19) = 0
        .ColWidth(20) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = " "
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
 
Dim uid As Double
uid = 0

' to save new record
 If Button.Caption = "Edit / Post EIC" Then
If esttran.cbo_spread.Text = "" Then
MsgBox "Select Spread"
esttran.cbo_spread.SetFocus
Exit Sub
End If
If esttran.cbo_tranx.Text = "" Then
MsgBox "Select Tranx"
esttran.cbo_tranx.SetFocus
Exit Sub
End If
If esttran.cbo_type.Text = "" Then
MsgBox "Select SUB-JC"
esttran.cbo_type.SetFocus
Exit Sub
End If
If esttran.cbo_jobcharge.Text = "" Then
MsgBox "Select Jobcharge"
esttran.cbo_jobcharge.SetFocus
Exit Sub
End If
If esttran.cbo_obs.Text = "" Then
MsgBox "Select OBS Code"
esttran.cbo_obs.SetFocus
Exit Sub
End If

If esttran.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
esttran.cbo_costcode.SetFocus
Exit Sub
End If
If esttran.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
esttran.txt_qty.SetFocus
Exit Sub
End If
If esttran.cbo_uom.Text = "" Then
MsgBox "Select UOM"
esttran.cbo_uom.SetFocus
Exit Sub
End If
If esttran.cbo_curr.Text = "" Then
MsgBox "Select Currency"
esttran.cbo_curr.SetFocus
Exit Sub
End If
If esttran.txt_unitrate.Text = "" Then
MsgBox "Enter Quantity"
esttran.txt_unitrate.SetFocus
Exit Sub
End If
On Error Resume Next
es = Split(esttran.cbo_spread.Text, "  -  ", Len(esttran.cbo_spread.Text), vbTextCompare)
es1 = Split(esttran.cbo_jobcharge.Text, "  -  ", Len(esttran.cbo_jobcharge.Text), vbTextCompare)
es2 = Split(esttran.cbo_costcode.Text, "  -  ", Len(esttran.cbo_costcode.Text), vbTextCompare)

 
Dim id1 As Double
id1 = 0
kn1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
If flex_grid.TextMatrix(flex_grid.Row, 1) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 1)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from cost", Cn, 3, 2
md.AddNew
        md!bd_year = cbo_year.Text
        md!bd_resccode = textresccode.Text
        md!bd_rescname = textrescname.Text
        md!bd_vendor = txt_vendor.Text
        md!bd_brate = txt_stdrate.Text
        md!bd_crate = txt_currate.Text
        md!bd_projectkey = textprojkey.Text
        md!bd_projectdesc = txt_projdesc.Text
        md!bd_costtype = "E"
        md!bd_respcode = txt_respcode.Text
        md!bd_respname = txt_respname.Text
        md!bd_cuttdate = main.DTPcutdate1.Value
        md!bd_spread = es(0)
        md!bd_tranx = esttran.cbo_tranx.Text
        md!bd_jobcharge = kn1(0)
        md!bd_costcode = es2(0)
        md!bd_qty = esttran.txt_qty.Text
        md!bd_days = esttran.txt_days.Text
        md!bd_tqty = esttran.txt_tqty.Text
        md!bd_uom = esttran.cbo_uom.Text
        md!bd_curr = esttran.cbo_curr.Text
        md!bd_xchg = esttran.txt_Xrate.Text
        md!bd_unitrate = esttran.txt_unitrate.Text
        md!bd_extdamt = esttran.txt_Extdamt.Text
        md!bd_e_days = esttran.txt_edays.Text
        md!bd_e_tqty = esttran.txt_etqty.Text
        md!bd_e_extdamt = esttran.txt_ectcamt.Text
        md!bd_edate = esttran.DTP_ed.Value
        md!bd_sdate = esttran.DTP_sd.Value
        md!bd_notes = esttran.txt_note.Text
        If md!bd_days >= 1 Then
        md!bd_chk = 1
        
        ElseIf md!bd_e_days >= 1 Then
        md!bd_chk = 1
        Else
        md!bd_chk = 0
        End If
        Dim dtt As Double
        Dim dtt1 As Double
        dtt = 0: dtt1 = 0
        dtt = md!bd_sdate
        dtt1 = (CDbl(md!bd_days) + CDbl(md!bd_e_days))
        If md!bd_chk = 1 Then
        md!bd_edate = Format(dtt + dtt1, "dd/MM/yyyy H:mm:ss")
        End If
               
        md!t_date = esttran.DTP_tdate.Value
        md!u_date = Now
        md!t_user = main.Label2.Caption
        md!bd_type = esttran.cbo_type.Text
        md!bd_obs = esttran.cbo_obs.Text
        md!estid = flex_grid.TextMatrix(flex_grid.Row, 1)
        md.Update
md.Close
Cn.Execute "delete from cost1 where bd_idd=" & id1


MsgBox "Selected Record Posted TO EIC"
Call flex_data1
'to modify existing record
ElseIf Button.Caption = "Close" Then
Unload Me
Unload esttran
End If

End Sub

Public Sub flex_data1()


 

Dim idddd As Double
idddd = 0
jgh = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
Dim ass As New ADODB.Recordset
If ass.State Then ass.Close
ass.Open "select * from cost1  where bd_jobcharge='" & jgh(0) & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='B'", Cn, 3, 2
While Not ass.EOF
If ass!bd_spread <> "NA" Then
idddd = ass!bd_id
Dim dys As Double
Dim perw As Double
dys = 0: perw = 0
nh = Split(ass!bd_jobcharge, "  -  ", Len(ass!bd_jobcharge), vbTextCompare)
ng = Split(ass!bd_spread, "  -  ", Len(ass!bd_spread), vbTextCompare)
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & ass!bd_jobcharge & "' and bdgt_spread_code='" & ass!bd_spread & "'", Cn, 3, 2
If Not bd.EOF Then
dys = bd!bdgt_days
'perw = bd!bdgt_per_workcomplete
End If
jk = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from cost1 where   bd_jobcharge='" & ass!bd_jobcharge & "' and bd_projectkey='" & jk(0) & "' and bd_spread='" & ass!bd_spread & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='B' and bd_id=" & idddd, Cn, 3, 2
If Not fl.EOF Then
        If fl!bd_spread <> "NA" Then
        fl!bd_days = dys
        fl!bd_tqty = (fl!bd_qty) * (dys)

        End If

fl!bd_extdamt = (fl!bd_xchg) * (fl!bd_unitrate) * (fl!bd_tqty) * ((100 + fl!bd_downtime) / 100) * ((100 + fl!bd_escl) / 100)
'fl!bd_wrkcomp = perw
fl!bd_bcwpamt = (fl!bd_wrkcomp / 100) * (fl!bd_extdamt)
fl.Update
End If
End If
ass.MoveNext
Wend

rscc = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
jk = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim btotal As Double
btotal = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost1  where  bd_jobcharge='" & rscc(0) & "'  and bd_projectkey='" & jk(0) & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='E' order by bd_costcode,bd_spread ,bd_jobcharge,bd_resccode", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata("bd_id")
         .TextMatrix(.Rows - 1, 1) = fldata("bd_idd")
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
        'load costcode
        
         
        
        
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
        If IsNull(fldata!bd_tqty) Then
        .TextMatrix(.Rows - 1, 9) = ""
        Else
        .TextMatrix(.Rows - 1, 9) = fldata!bd_tqty
        End If
        .TextMatrix(.Rows - 1, 10) = fldata!bd_uom
        .TextMatrix(.Rows - 1, 11) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 12) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 13) = fldata!bd_unitrate
        
         
        
        Dim pg As New ADODB.Recordset
        If pg.State Then pg.Close
        pg.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not pg.EOF Then
            .TextMatrix(.Rows - 1, 14) = Format(pg!prgs_startdate, "dd/MM/yyyy H:mm:ss")
            .TextMatrix(.Rows - 1, 15) = Format(pg!prgs_enddate, "dd/MM/yyyy H:mm:ss")
            
            Else
            .TextMatrix(.Rows - 1, 14) = Format(fldata!bd_sdate, "dd/MM/yyyy H:mm:ss")
            .TextMatrix(.Rows - 1, 15) = Format(fldata!bd_edate, "dd/MM/yyyy H:mm:ss")
            
        End If
        .TextMatrix(.Rows - 1, 16) = Format(fldata!bd_extdamt, "###,###,###,###,##0.00")
         gtotal = gtotal + fldata!bd_extdamt
        .TextMatrix(.Rows - 1, 17) = Format(fldata!bd_wrkcomp, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 18) = Format(fldata!bd_bcwpamt, "###,###,###,###,##0.00")
        .TextMatrix(.Rows - 1, 19) = fldata!bd_notes
        .TextMatrix(.Rows - 1, 20) = fldata!t_date
        btotal = btotal + fldata!bd_bcwpamt
        
        fldata.MoveNext
        Wend
        kk = 0
        kk = flex_grid.Rows
End With
Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
txt_btotal.Text = Format(btotal, "###,###,##0.00")
End Sub



