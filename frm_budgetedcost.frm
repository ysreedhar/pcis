VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_budgetedcost 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Budgeted Cost by Jobcharge"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   14880
   LinkTopic       =   "Form2"
   ScaleHeight     =   10260
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   11175
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transaction Information"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4920
         TabIndex        =   36
         Top             =   960
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Project Information"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4920
         TabIndex        =   24
         Top             =   120
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   29
            Text            =   "BCWP- RM"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resource Information"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   120
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11175
      Begin VB.ComboBox cbo_pproj 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
      Begin VB.ComboBox cbo_year 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbo_resc 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6720
         TabIndex        =   1
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
         Left            =   6720
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
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Copy To Excel"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   8
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
            Picture         =   "frm_budgetedcost.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedcost.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   7095
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   12515
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
End
Attribute VB_Name = "frm_budgetedcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_pproj_Change()
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

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        textrescname.Text = fl!resc_desc
        textresccode.Text = fl!resc_code
        textprojkey.Text = gg(0)
        txt_projdesc.Text = gg(1)
        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        textcosttype.Text = "B"
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
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
            If Not fl1.EOF Then
            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close

 
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
gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        textrescname.Text = fl!resc_desc
        textresccode.Text = fl!resc_code
        textprojkey.Text = gg(0)
        txt_projdesc.Text = gg(1)
        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        textcosttype.Text = "B"
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
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close


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

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        textrescname.Text = fl!resc_desc
        textresccode.Text = fl!resc_code
        textprojkey.Text = gg(0)
        txt_projdesc.Text = gg(1)
        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        textcosttype.Text = "B"
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
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close


Call flex_data1

End Sub

Private Sub cbo_resc_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'KeyAscii = 0
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
 

'''gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
'''Dim rs As New ADODB.Recordset
'''If rs.State Then rs.Close
'''rs.Open "select DISTINCT(rd.dresc_code),rm.resc_desc from resourcedetails rd ,resourcemaster rm  where  rm.resc_id=rd.resc_id and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' order by rd.dresc_code", Cn, 3, 2
'''While Not rs.EOF
'''cbo_resc.AddItem rs(0) & "  -  " & rs(1)
'''rs.MoveNext
'''Wend
'''rs.Close
'''
'''kl1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
'''If cbo_resc.Text = "" Then
''''MsgBox "Select Resource Code"
'''cbo_resc.SetFocus
'''Exit Sub
'''End If
'''
'''Dim fl As New ADODB.Recordset
'''If fl.State Then fl.Close
'''fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
'''        If Not fl.EOF Then
'''        textrescname.Text = fl!resc_desc
'''        textresccode.Text = fl!resc_code
'''        textprojkey.Text = gg(0)
'''        txt_projdesc.Text = gg(1)
'''        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
'''        textcosttype.Text = "B"
'''        txt_vendor.Text = fl!resc_vendorcode
'''        txt_respcode.Text = fl!resc_respcode
'''        Dim rr As New ADODB.Recordset
'''        If rr.State Then rr.Close
'''        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
'''        If Not rr.EOF Then
'''        txt_respname.Text = rr(0)
'''        End If
'''        Text3.Text = fl!dresc_curcy
'''        End If
'''fl.Close
'''
'''Dim fl1 As New ADODB.Recordset
'''        If fl1.State Then fl1.Close
'''        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
'''            If Not fl1.EOF Then
'''            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
'''            Text4.Text = fl1!dresc_curcy
'''            End If
'''        fl1.Close
'''
'''
'''Call flex_data1
'''
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
 
'On Error Resume Next
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
 

'''gg = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
'''
'''Dim rs As New ADODB.Recordset
'''If rs.State Then rs.Close
'''rs.Open "select DISTINCT(rd.dresc_code),rm.resc_desc from resourcedetails rd ,resourcemaster rm  where  rm.resc_id=rd.resc_id and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' order by rd.dresc_code", Cn, 3, 2
'''While Not rs.EOF
'''cbo_resc.AddItem rs(0) & "  -  " & rs(1)
'''rs.MoveNext
'''Wend
'''rs.Close
'''
'''kl1 = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
'''If cbo_resc.Text = "" Then
''''MsgBox "Select Resource Code"
'''cbo_resc.SetFocus
'''Exit Sub
'''End If
'''
'''Dim fl As New ADODB.Recordset
'''If fl.State Then fl.Close
'''fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
'''        If Not fl.EOF Then
'''        textrescname.Text = fl!resc_desc
'''        textresccode.Text = fl!resc_code
'''        textprojkey.Text = gg(0)
'''        txt_projdesc.Text = gg(1)
'''        txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
'''        textcosttype.Text = "B"
'''        txt_vendor.Text = fl!resc_vendorcode
'''        txt_respcode.Text = fl!resc_respcode
'''       Dim rr As New ADODB.Recordset
'''        If rr.State Then rr.Close
'''        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
'''        If Not rr.EOF Then
'''        txt_respname.Text = rr(0)
'''        End If
'''        Text3.Text = fl!dresc_curcy
'''        End If
'''fl.Close
'''
'''Dim fl1 As New ADODB.Recordset
'''        If fl1.State Then fl1.Close
'''        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl1(0) & "' and rd.dresc_year='" & cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
'''            If Not fl1.EOF Then
'''            txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
'''            Text4.Text = fl1!dresc_curcy
'''            End If
'''        fl1.Close
'''
'''
'''Call flex_data1

End Sub

Private Sub cbo_year_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub flex_grid_Click()
On Error Resume Next
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture



vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture


Unload budgetedcost1
Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
budgetedcost1.Show
  sv!bd_year = cbo_year.Text
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
        
        budgetedcost1.txt_qty.Text = rsd!bd_qty
        budgetedcost1.txt_days.Text = rsd!bd_days
        budgetedcost1.txt_totdays.Text = rsd!bd_tqty
        budgetedcost1.txt_unitrate.Text = Format(rsd!bd_unitrate, "###,###,##0.00")
        budgetedcost1.txt_Xrate.Text = rsd!bd_xchg
        budgetedcost1.txt_downtime.Text = rsd!bd_downtime
        budgetedcost1.txt_esclfactor.Text = rsd!bd_escl
        budgetedcost1.txt_Extdamt.Text = Format(rsd!bd_extdamt, "###,###,##0.00")
        budgetedcost1.txt_wrkcomp.Text = rsd!bd_wrkcomp
        budgetedcost1.txt_bcwpamt.Text = Format(rsd!bd_bcwpamt, "###,###,##0.00")
        
        
        Dim rr1 As New ADODB.Recordset
        If rr1.State Then rr1.Close
        rr1.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & rsd!bd_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        txt_respname.Text = rr(0)
        End If
        txt_brate.Text = Format(rsd!bd_brate, "###,###,##0.00")
        txt_crate.Text = Format(rsd!bd_crate, "###,###,##0.00")
 End If
        budgetedcost1.cbo_spread.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
        budgetedcost1.cbo_tranx.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
        budgetedcost1.cbo_jobcharge.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
        budgetedcost1.cbo_obs.Text = flex_grid.TextMatrix(flex_grid.Row, 18)
        budgetedcost1.cbo_costcode.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
'''        budgetedcost1.txt_qty.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 6), "###,###,##0.00")
'''        budgetedcost1.txt_days.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 7), "###,###,##0.00")
'''        budgetedcost1.txt_totdays.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 8), "###,###,##0.00")
        budgetedcost1.cbo_uom.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
        budgetedcost1.cbo_curr.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
'''        budgetedcost1.txt_unitrate.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 12), "###,###,##0.00")
'''        budgetedcost1.txt_Xrate.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 11), "###,###,##0.00")
'''        budgetedcost1.txt_downtime.Text = flex_grid.TextMatrix(flex_grid.Row, 13)
'''        budgetedcost1.txt_esclfactor.Text = flex_grid.TextMatrix(flex_grid.Row, 14)
'''        budgetedcost1.txt_Extdamt.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 15), "###,###,##0.00")
'''        budgetedcost1.txt_wrkcomp.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 16), "###,##0.00")
'''        budgetedcost1.txt_bcwpamt.Text = Format(flex_grid.TextMatrix(flex_grid.Row, 17), "###,###,##0.00")
        budgetedcost1.cboChargeType.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
        budgetedcost1.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 19)
        budgetedcost1.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 20)
budgetedcost1.Show
budgetedcost1.Top = 3200
budgetedcost1.Left = 0
budgetedcost1.Height = 3540
budgetedcost1.Width = 8850
rsd.Close
vprev = flex_grid.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BC BY JOBCHARGE"
Call flex_title
Call flex_data1
 
Me.Top = 5
Me.Left = 5
DTP_cod.Value = Format(Date, "dd/mm/yyyy")
Dim i As Integer
i = 0
For i = 2000 To 2050
cbo_year.AddItem i
Next i

 Me.Width = 11415
 Me.Height = 9750
 
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
End Sub
Public Sub flex_title()
On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "TranX"
        .ColWidth(1) = 600
        .ColAlignment(1) = Left
        .TextMatrix(0, 2) = "Spread"
        .ColWidth(2) = 1100
        .ColAlignment(2) = 0
         
        .TextMatrix(0, 3) = "Resource"
        .ColWidth(3) = 3500
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Cg. Type"
        .ColWidth(4) = 600
        .TextMatrix(0, 5) = "Costcode"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Qty"
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 7) = "Days"
        .ColWidth(7) = 620
        
        .TextMatrix(0, 8) = "TotalQty"
        .ColWidth(8) = 1000
        
        .TextMatrix(0, 9) = "UOM"
        .ColWidth(9) = 600
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Curcy"
        .ColWidth(10) = 600
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "Xrate"
        .ColWidth(11) = 600
        
        .TextMatrix(0, 12) = "UnitRate"
        .ColWidth(12) = 1000
        
        .TextMatrix(0, 13) = "D/Time %"
        .ColWidth(13) = 600
        
        .TextMatrix(0, 14) = "Escl %"
        .ColWidth(14) = 600
        
        .TextMatrix(0, 15) = "BDGT(RM)"
        .ColWidth(15) = 1100
        
        .TextMatrix(0, 16) = "% Wrk Compltd"
        .ColWidth(16) = 600
        .TextMatrix(0, 17) = "BCWP(RM)"
        .ColWidth(17) = 1100
        .TextMatrix(0, 18) = "OBS"
        .ColWidth(18) = 600
        .TextMatrix(0, 19) = "Notes"
        .ColWidth(19) = 4000
        .ColAlignment(19) = 0
        .ColWidth(20) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload budgetedcost1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then

If cbo_year.Text = "" Then
MsgBox "select Year"
cbo_year.SetFocus
Exit Sub
End If
If cbo_pproj.Text = "" Then
MsgBox "select Project"
cbo_pproj.SetFocus
Exit Sub
End If
If cbo_resc.Text = "" Then
MsgBox "select Resource"
cbo_resc.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload budgetedcost1
budgetedcost1.Show
budgetedcost1.Top = 3200
budgetedcost1.Left = 0
budgetedcost1.Height = 3540
budgetedcost1.Width = 8850
Dim uid As Double
uid = 0
' to save new record
ElseIf Button.Caption = "Save" Then

'validate
 If budgetedcost1.cbo_spread.Text = "" Then
 MsgBox "select Spread"
 budgetedcost1.cbo_spread.SetFocus
 Exit Sub
 End If
 If budgetedcost1.cbo_tranx.Text = "" Then
 MsgBox "select Tranx"
 budgetedcost1.cbo_tranx.SetFocus
 Exit Sub
 End If
 If budgetedcost1.cbo_jobcharge.Text = "" Then
 MsgBox "select Resource"
 budgetedcost1.cbo_jobcharge.SetFocus
 Exit Sub
 End If
    If budgetedcost1.cbo_obs.Text = "" Then
    MsgBox "select OBS"
    budgetedcost1.cbo_obs.SetFocus
    Exit Sub
    End If
If budgetedcost1.cbo_costcode.Text = "" Then
MsgBox "select CostCode"
budgetedcost1.cbo_costcode.SetFocus
Exit Sub
End If
If budgetedcost1.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
budgetedcost1.txt_qty.SetFocus
Exit Sub
End If
If budgetedcost1.cbo_uom.Text = "" Then
MsgBox "select UOM"
budgetedcost1.cbo_uom.SetFocus
Exit Sub
End If
If budgetedcost1.cbo_curr.Text = "" Then
MsgBox "select Currency"
budgetedcost1.cbo_curr.SetFocus
Exit Sub
End If
If budgetedcost1.txt_unitrate.Text = "" Then
MsgBox "Enter UnitRate"
budgetedcost1.txt_unitrate.SetFocus
Exit Sub
End If
If budgetedcost1.txt_unitrate.Text = 0 Then
MsgBox "Enter UnitRate"
budgetedcost1.txt_unitrate.SetFocus
Exit Sub
End If
ng = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
ng1 = Split(budgetedcost1.cbo_spread.Text, "  -  ", Len(budgetedcost1.cbo_spread.Text), vbTextCompare)
ng2 = Split(budgetedcost1.cbo_costcode.Text, "  -  ", Len(budgetedcost1.cbo_costcode.Text), vbTextCompare)
 
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from cost ", Cn, 3, 2
sv.AddNew
        
        sv!bd_year = cbo_year.Text
        sv!bd_resccode = textresccode.Text
        sv!bd_rescname = textrescname.Text
        sv!bd_vendor = txt_vendor.Text
        sv!bd_projectkey = textprojkey.Text
        sv!bd_projectdesc = txt_projdesc.Text
        sv!bd_costtype = textcosttype.Text
        sv!bd_respcode = txt_respcode.Text
        sv!bd_respname = txt_respname.Text
        sv!bd_brate = txt_brate.Text
        sv!bd_crate = txt_crate.Text
        sv!bd_spread = ng1(0)
        sv!bd_tranx = budgetedcost1.cbo_tranx.Text
        sv!bd_JobCharge = ng(0)
        sv!bd_costcode = ng2(0)
        sv!bd_qty = budgetedcost1.txt_qty.Text
        sv!bd_days = budgetedcost1.txt_days.Text
        sv!bd_tqty = budgetedcost1.txt_totdays.Text
        sv!bd_uom = budgetedcost1.cbo_uom.Text
        sv!bd_curr = budgetedcost1.cbo_curr.Text
        sv!bd_unitrate = budgetedcost1.txt_unitrate.Text
        sv!bd_xchg = budgetedcost1.txt_Xrate.Text
        sv!bd_downtime = budgetedcost1.txt_downtime.Text
        sv!bd_escl = budgetedcost1.txt_esclfactor.Text
        sv!bd_extdamt = budgetedcost1.txt_Extdamt.Text
        sv!bd_wrkcomp = budgetedcost1.txt_wrkcomp.Text
        sv!bd_bcwpamt = budgetedcost1.txt_bcwpamt.Text
        sv!bd_ChargeType = budgetedcost1.cboChargeType.Text
        sv!bd_notes = budgetedcost1.txt_notes.Text
        sv!t_date = budgetedcost1.DTP_tdate.Value
        sv!u_date = Now
        sv!t_user = main.Label2.Caption
        sv!bd_obs = budgetedcost1.cbo_obs.Text
sv.Update
sv.Close

Call flex_data1
'Call flex_title

MsgBox "New Record Added Succesfully"
Unload budgetedcost1



 
 '----------------END--------------


'to modify existing record
ElseIf Button.Caption = "Modify" Then

 If budgetedcost1.cbo_spread.Text = "" Then
 MsgBox "select Spread"
 budgetedcost1.cbo_spread.SetFocus
 Exit Sub
 End If
 If budgetedcost1.cbo_tranx.Text = "" Then
 MsgBox "select Tranx"
 budgetedcost1.cbo_tranx.SetFocus
 Exit Sub
 End If
 If budgetedcost1.cbo_jobcharge.Text = "" Then
 MsgBox "select Resource"
 budgetedcost1.cbo_jobcharge.SetFocus
 Exit Sub
 End If
     If budgetedcost1.cbo_obs.Text = "" Then
    MsgBox "select OBS"
    budgetedcost1.cbo_obs.SetFocus
    Exit Sub
    End If
If budgetedcost1.cbo_costcode.Text = "" Then
MsgBox "select CostCode"
budgetedcost1.cbo_costcode.SetFocus
Exit Sub
End If
If budgetedcost1.txt_qty.Text = "" Then
MsgBox "Enter Quantity"
budgetedcost1.txt_qty.SetFocus
Exit Sub
End If
If budgetedcost1.cbo_uom.Text = "" Then
MsgBox "select UOM"
budgetedcost1.cbo_uom.SetFocus
Exit Sub
End If
If budgetedcost1.cbo_curr.Text = "" Then
MsgBox "select Currency"
budgetedcost1.cbo_curr.SetFocus
Exit Sub
End If
If budgetedcost1.txt_unitrate.Text = "" Then
MsgBox "Enter UnitRate"
budgetedcost1.txt_unitrate.SetFocus
Exit Sub
End If
If budgetedcost1.txt_unitrate.Text = 0 Then
MsgBox "Enter UnitRate"
budgetedcost1.txt_unitrate.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0

ng = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
ng1 = Split(budgetedcost1.cbo_spread.Text, "  -  ", Len(budgetedcost1.cbo_spread.Text), vbTextCompare)
ng2 = Split(budgetedcost1.cbo_costcode.Text, "  -  ", Len(budgetedcost1.cbo_costcode.Text), vbTextCompare)
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
                            md!bd_projectkey = textprojkey.Text
                            md!bd_projectdesc = txt_projdesc.Text
                            md!bd_costtype = textcosttype.Text
                            md!bd_respcode = txt_respcode.Text
                            md!bd_respname = txt_respname.Text
                            md!bd_brate = txt_brate.Text
                            md!bd_crate = txt_crate.Text
                            md!bd_spread = ng1(0)
                            md!bd_tranx = budgetedcost1.cbo_tranx.Text
                            md!bd_JobCharge = ng(0)
                            md!bd_costcode = ng2(0)
                            md!bd_qty = budgetedcost1.txt_qty.Text
                            md!bd_days = budgetedcost1.txt_days.Text
                            md!bd_tqty = budgetedcost1.txt_totdays.Text
                            md!bd_uom = budgetedcost1.cbo_uom.Text
                            md!bd_curr = budgetedcost1.cbo_curr.Text
                            md!bd_unitrate = budgetedcost1.txt_unitrate.Text
                            md!bd_xchg = budgetedcost1.txt_Xrate.Text
                            md!bd_downtime = budgetedcost1.txt_downtime.Text
                            md!bd_escl = budgetedcost1.txt_esclfactor.Text
                            md!bd_extdamt = budgetedcost1.txt_Extdamt.Text
                            md!bd_wrkcomp = budgetedcost1.txt_wrkcomp.Text
                            md!bd_bcwpamt = budgetedcost1.txt_bcwpamt.Text
                            md!bd_ChargeType = budgetedcost1.cboChargeType.Text
                            md!bd_notes = budgetedcost1.txt_notes.Text
                            md!t_date = budgetedcost1.DTP_tdate.Value
                            md!u_date = Now
                            md!t_user = main.Label2.Caption
                            md!bd_obs = budgetedcost1.cbo_obs.Text
                    md.Update
                    md.Close
                    End If
Call flex_data1
'Call flex_title
MsgBox "Selected Record Modified"
Unload budgetedcost1
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
Unload budgetedcost1
Call flex_data1
'Call flex_title


Else
Unload budgetedcost1
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload budgetedcost1
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

Public Sub flex_data1()
On Error Resume Next
If cbo_resc.Text = "" Then Exit Sub
Dim fi As Double
fi = 0
Dim dys As Double
Dim perw As Double
jgh = Split(cbo_resc.Text, "  -  ", Len(cbo_resc.Text), vbTextCompare)
Dim ass As New ADODB.Recordset
If ass.State Then ass.Close
ass.Open "select * from cost  where bd_jobcharge='" & jgh(0) & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='B'", Cn, 3, 2
While Not ass.EOF
If ass!bd_spread <> "NA" Then
fi = ass!bd_id

dys = 0: perw = 0
nh = Split(ass!bd_JobCharge, "  -  ", Len(ass!bd_JobCharge), vbTextCompare)
ng = Split(ass!bd_spread, "  -  ", Len(ass!bd_spread), vbTextCompare)
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & ass!bd_JobCharge & "' and bdgt_spread_code='" & ass!bd_spread & "'", Cn, 3, 2
If Not bd.EOF Then
dys = bd!bdgt_days
'perw = bd!bdgt_per_workcomplete
End If
jk = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from cost where   bd_jobcharge='" & ass!bd_JobCharge & "' and bd_projectkey='" & jk(0) & "' and bd_spread='" & ass!bd_spread & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='B' and bd_id=" & fi, Cn, 3, 2
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
fldata.Open "select * from cost  where  bd_jobcharge='" & rscc(0) & "'  and bd_projectkey='" & jk(0) & "' and bd_year= '" & cbo_year.Text & "' and  bd_costtype='B' order by bd_costcode,bd_spread ,bd_jobcharge,bd_resccode", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!bd_tranx
        Dim spd As New ADODB.Recordset
        If spd.State Then spd.Close
        spd.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spd.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread & "  -  " & spd(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread
        End If
        spd.Close
Dim ki As New ADODB.Recordset
If ki.State Then ki.Close
ki.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & fldata!bd_resccode & "' ", Cn, 3, 2
If Not ki.EOF Then
.TextMatrix(.Rows - 1, 3) = fldata!bd_resccode & "  -  " & ki(0)
Else
.TextMatrix(.Rows - 1, 3) = fldata!bd_resccode
End If
        .TextMatrix(.Rows - 1, 4) = fldata!bd_ChargeType
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 5) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 5) = fldata!bd_costcode
        End If
        cs.Close
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_qty, "###,###,##0.000")
        If IsNull(fldata!bd_days) Then
        .TextMatrix(.Rows - 1, 7) = ""
        Else
        .TextMatrix(.Rows - 1, 7) = fldata!bd_days
        End If
        .TextMatrix(.Rows - 1, 8) = fldata!bd_tqty
        .TextMatrix(.Rows - 1, 9) = fldata!bd_uom
        .TextMatrix(.Rows - 1, 10) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 11) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 12) = fldata!bd_unitrate
        .TextMatrix(.Rows - 1, 13) = Format(fldata!bd_downtime, "###,###,##0.000")
        .TextMatrix(.Rows - 1, 14) = Format(fldata!bd_escl, "###,###,##0.000")
        .TextMatrix(.Rows - 1, 15) = Format(fldata!bd_extdamt, "###,###,###,###,##0.00")
         gtotal = gtotal + fldata!bd_extdamt
        .TextMatrix(.Rows - 1, 16) = Format(fldata!bd_wrkcomp, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 17) = Format(fldata!bd_bcwpamt, "###,###,###,###,##0.00")
        .TextMatrix(.Rows - 1, 18) = fldata!bd_obs
        .TextMatrix(.Rows - 1, 19) = fldata!bd_notes
        .TextMatrix(.Rows - 1, 20) = fldata!t_date
        btotal = btotal + fldata!bd_bcwpamt
           
        fldata.MoveNext
        Wend
End With
Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
txt_btotal.Text = Format(btotal, "###,###,##0.00")
End Sub
