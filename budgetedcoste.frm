VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form budgetedcoste 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fr3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "BDGT   /  BCWP  Details"
         TabPicture(0)   =   "budgetedcoste.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Notes"
         TabPicture(1)   =   "budgetedcoste.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame7"
         Tab(1).Control(1)=   "Label12"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   -75000
            TabIndex        =   41
            Top             =   300
            Width           =   9015
            Begin VB.TextBox txt_notes 
               Height          =   2415
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   42
               Top             =   240
               Width           =   8295
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   0
            TabIndex        =   2
            Top             =   300
            Width           =   9015
            Begin VB.ComboBox cbo_costcode 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5520
               TabIndex        =   33
               Top             =   960
               Width           =   3135
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "BDGT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   120
               TabIndex        =   12
               Top             =   1320
               Width           =   6375
               Begin VB.TextBox txt_esclfactor 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   22
                  Text            =   "0"
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox txt_downtime 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2820
                  TabIndex        =   21
                  Text            =   "0"
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox txt_Xrate 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   20
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.TextBox txt_Extdamt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000E&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "###,###,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
                  Height          =   285
                  Left            =   4560
                  TabIndex        =   19
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.ComboBox cbo_uom 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3720
                  TabIndex        =   18
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.ComboBox cbo_curr 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5040
                  TabIndex        =   17
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox txt_unitrate 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "###,###,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
                  Height          =   285
                  Left            =   1110
                  TabIndex        =   16
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.TextBox txt_qty 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   15
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.TextBox txt_days 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   14
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox txt_totdays 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2340
                  TabIndex        =   13
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.Label Label31 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Escl %"
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
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   32
                  Top             =   840
                  Width           =   735
               End
               Begin VB.Label Label30 
                  BackStyle       =   0  'Transparent
                  Caption         =   "D/ Time %"
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
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   31
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "BDGT Amount"
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
                  Left            =   4560
                  TabIndex        =   30
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.Label Label10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Xchg Rate"
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
                  TabIndex        =   29
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Unit Rate"
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
                  Left            =   1110
                  TabIndex        =   28
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Currency"
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
                  Left            =   5040
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "UOM"
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
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Total Quantity"
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
                  Left            =   2340
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Days"
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
                  Left            =   1320
                  TabIndex        =   24
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Quantity"
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
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.ComboBox cbo_tranx 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4560
               TabIndex        =   11
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox cbo_jobcharge 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   4095
            End
            Begin VB.ComboBox cbo_spread 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   4335
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "BCWP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   6600
               TabIndex        =   4
               Top             =   1320
               Width           =   2055
               Begin VB.TextBox txt_bcwpamt 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "###,###,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
                  Height          =   285
                  Left            =   120
                  TabIndex        =   6
                  Text            =   "0"
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.TextBox txt_wrkcomp 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   5
                  Text            =   "0"
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label32 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "% Complete"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   8
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label33 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "BCWP Amount"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   7
                  Top             =   900
                  Width           =   1215
               End
            End
            Begin VB.ComboBox cbo_obs 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4320
               TabIndex        =   3
               Text            =   "XX"
               Top             =   960
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker DTP_tdate 
               Height          =   315
               Left            =   6000
               TabIndex        =   34
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   64487425
               CurrentDate     =   38733
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Transaction Date"
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
               Left            =   6000
               TabIndex        =   40
               Top             =   120
               Width           =   1230
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "CostCode  -  Description"
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
               Left            =   5520
               TabIndex        =   39
               Top             =   720
               Width           =   1755
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "RescCode  -  Description"
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
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label26 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Transaction Type"
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
               Left            =   4560
               TabIndex        =   37
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Spread Code - Description"
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
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label13 
               BackColor       =   &H00FFFFFF&
               Caption         =   "OBS Code"
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
               Left            =   4320
               TabIndex        =   35
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BDGT / BCWP Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   -72840
            TabIndex        =   43
            Top             =   0
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "budgetedcoste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
