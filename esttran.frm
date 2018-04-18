VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form esttran 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6588
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
         ForeColor       =   12582912
         TabCaption(0)   =   "ACWP / ECTC Details"
         TabPicture(0)   =   "esttran.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Notes"
         TabPicture(1)   =   "esttran.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(1)=   "Label12"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3400
            Left            =   0
            TabIndex        =   6
            Top             =   300
            Width           =   9855
            Begin VB.ComboBox cbo_spread 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   43
               Top             =   480
               Width           =   4335
            End
            Begin VB.ComboBox cbo_tranx 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4560
               TabIndex        =   42
               Top             =   480
               Width           =   1335
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFFFFF&
               Height          =   1695
               Left            =   120
               TabIndex        =   35
               Top             =   1440
               Width           =   2295
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   360
                  Width           =   375
               End
               Begin MSComCtl2.DTPicker DTP_ed 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1200
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "dd-MM-yyyy H:mm:ss"
                  Format          =   64290819
                  CurrentDate     =   37987
               End
               Begin MSComCtl2.DTPicker DTP_sd 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   38
                  Top             =   600
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   609
                  _Version        =   393216
                  CustomFormat    =   "dd-MM-yyyy H:mm:ss"
                  Format          =   64290819
                  CurrentDate     =   37987
               End
               Begin VB.Label lbl 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Apply  Date"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800080&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.Label Label27 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Start Date"
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
                  Left            =   1440
                  TabIndex        =   40
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label Label26 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "End Date"
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
                  Left            =   1440
                  TabIndex        =   39
                  Top             =   960
                  Width           =   855
               End
            End
            Begin VB.ComboBox cbo_jobcharge 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   34
               Top             =   1080
               Width           =   4335
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ACWP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   2400
               TabIndex        =   17
               Top             =   1440
               Width           =   5265
               Begin VB.ComboBox cbo_uom 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   25
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.ComboBox cbo_curr 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   24
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.TextBox txt_tqty 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   23
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox txt_days 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   22
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox txt_qty 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   21
                  Top             =   480
                  Width           =   1095
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
                  Left            =   2400
                  TabIndex        =   20
                  Top             =   1200
                  Width           =   1215
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
                  Left            =   3720
                  TabIndex        =   19
                  Top             =   1200
                  Width           =   1455
               End
               Begin VB.TextBox txt_Xrate 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   18
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Total Quantity"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800080&
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   33
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Days"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800080&
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   32
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Quantity"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Unit Rate"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   30
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Currency"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "UOM"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ACWP Amount(RM)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   3720
                  TabIndex        =   27
                  Top             =   960
                  Width           =   1410
               End
               Begin VB.Label Label10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Xchg Rate"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   26
                  Top             =   960
                  Width           =   855
               End
            End
            Begin VB.ComboBox cbo_costcode 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5760
               TabIndex        =   16
               Top             =   1080
               Width           =   3495
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ECTC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   7680
               TabIndex        =   9
               Top             =   1440
               Width           =   1785
               Begin VB.TextBox txt_edays 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   12
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txt_ectcamt 
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
                  TabIndex        =   11
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txt_etqty 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   840
                  TabIndex        =   10
                  Top             =   480
                  Width           =   855
               End
               Begin VB.Label Label23 
                  BackStyle       =   0  'Transparent
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
                  ForeColor       =   &H00800080&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label Label24 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Total Qty"
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
                  Left            =   840
                  TabIndex        =   14
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ECTC Amount(RM)"
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
                  Left            =   120
                  TabIndex        =   13
                  Top             =   960
                  Width           =   1335
               End
            End
            Begin VB.ComboBox cbo_type 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6000
               TabIndex        =   8
               Text            =   "A"
               Top             =   480
               Width           =   975
            End
            Begin VB.ComboBox cbo_obs 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4560
               TabIndex        =   7
               Text            =   "XX"
               Top             =   1080
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker DTP_tdate 
               Height          =   315
               Left            =   7080
               TabIndex        =   44
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   64290817
               CurrentDate     =   38733
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Spread Code"
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
               TabIndex        =   51
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label30 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               TabIndex        =   50
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFC0C0&
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
               TabIndex        =   49
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label32 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Cost Code"
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
               Left            =   5760
               TabIndex        =   48
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               Left            =   7080
               TabIndex        =   47
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "SUB-JC"
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
               TabIndex        =   46
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label14 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               Left            =   4560
               TabIndex        =   45
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2655
            Left            =   -74880
            TabIndex        =   4
            Top             =   360
            Width           =   9975
            Begin VB.TextBox txt_notes 
               Height          =   2295
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   5
               Top             =   240
               Width           =   9615
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3400
            Left            =   -75000
            TabIndex        =   2
            Top             =   300
            Width           =   9855
            Begin VB.TextBox txt_note 
               Height          =   2775
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   3
               Top             =   240
               Width           =   9255
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ACWP / ECTC Details"
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
            TabIndex        =   53
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "  Notes"
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
            Left            =   -73200
            TabIndex        =   52
            Top             =   0
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "esttran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_curr_Change()
Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
If Not cr1.EOF Then
txt_Xrate.Text = cr1!cur_xchgrate
End If
cr1.Close
Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from resourcedetails where dresc_code='" & frm_estpost.textresccode.Text & "' and dresc_proj='" & frm_estpost.textprojkey.Text & "' and dresc_curcy='" & cbo_curr.Text & "' and dresc_year='" & frm_estpost.cbo_year.Text & "'", Cn, 3, 2
If Not cr2.EOF Then
txt_unitrate.Text = cr2!dresc_rate
End If
txt_Xrate_Change
End Sub

Private Sub cbo_curr_Click()
Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
If Not cr1.EOF Then
txt_Xrate.Text = cr1!cur_xchgrate
End If
cr1.Close
Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from resourcedetails where dresc_code='" & frm_estpost.textresccode.Text & "' and dresc_proj='" & frm_estpost.textprojkey.Text & "' and dresc_curcy='" & cbo_curr.Text & "' and dresc_year='" & frm_estpost.cbo_year.Text & "'", Cn, 3, 2
If Not cr2.EOF Then
txt_unitrate.Text = cr2!dresc_rate
End If
txt_Xrate_Change
End Sub

Private Sub cbo_jobcharge_Change()

On Error Resume Next


kl = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
nj = Split(frm_estpost.cbo_pproj.Text, "  -  ", Len(frm_estpost.cbo_pproj.Text), vbTextCompare)

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & frm_estpost.cbo_year.Text & "' and dresc_proj='" & nj(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        frm_estpost.textrescname.Text = fl!resc_desc
        frm_estpost.textresccode.Text = fl!resc_code
        frm_estpost.textprojkey.Text = nj(0)
        frm_estpost.txt_projdesc.Text = nj(1)
        frm_estpost.txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        frm_estpost.textcosttype.Text = "E"
        frm_estpost.txt_vendor.Text = fl!resc_vendorcode
        frm_estpost.txt_respcode.Text = fl!resc_respcode
        Dim rr As New ADODB.Recordset
        If rr.State Then rr.Close
        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        frm_estpost.txt_respname.Text = rr(0)
        End If
        frm_estpost.Text3.Text = fl!dresc_curcy
        
         
        End If
fl.Close

Dim fl1 As New ADODB.Recordset
        If fl1.State Then fl1.Close
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & frm_estpost.cbo_year.Text & "' and dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            frm_estpost.txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            frm_estpost.Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close



'''''''''''''''
spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
spl1 = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)
cbo_type.Clear
cbo_type.Text = "A"
Dim ty As New ADODB.Recordset
If ty.State Then ty.Close
ty.Open "select Distinct(prgs_type) from progressdurationdetails where prgs_spread_code='" & spl(0) & "' and prgs_job_key ='" & spl1(0) & "' order by prgs_type ", Cn, 3, 2
While Not ty.EOF
cbo_type.AddItem ty(0)
ty.MoveNext
Wend
ty.Close
'cbo_costcode.ListIndex = 0
cbo_type.ListIndex = 0


nn = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)
nnm = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)

 If cbo_type.Text = "" Then cbo_type.Text = "A"
 
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from progressdurationdetails where prgs_spread_code='" & nnm(0) & "' and prgs_job_key='" & nn(0) & "' and prgs_type='" & cbo_type.Text & "'", Cn, 3, 2
If Not bd.EOF Then
DTP_sd.Value = Format(bd!prgs_startdate, "dd-MM-yyyy H:mm:ss")
DTP_ed.Value = Format(bd!prgs_enddate, "dd-MM-yyyy H:mm:ss")
Else
DTP_sd.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
DTP_ed.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
End If


Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then
lbl.Visible = True
Check1.Visible = True
If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
   txt_days.Enabled = True
   txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                    c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
                   
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_days.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_edays.Text
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                    
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If

Dim cc As New ADODB.Recordset
If cc.State Then cc.Close
cc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & frm_estpost.textresccode.Text & "' ", Cn, 3, 2
While Not cc.EOF
cbo_costcode.AddItem cc(0) & "  -  " & cc(1)
cc.MoveNext
Wend
cc.Close

Dim cr As New ADODB.Recordset
            If cr.State Then cr.Close
            cr.Open "select DISTINCT(dresc_curcy) from resourcedetails where  dresc_code='" & frm_estpost.textresccode.Text & "' and dresc_proj='" & nj(0) & "' and  dresc_year='" & frm_estpost.cbo_year.Text & "' order by  dresc_curcy", Cn, 3, 2
            If Not cr.EOF Then
            cbo_curr.Text = cr(0)
            End If
cr.Close


Dim um As New ADODB.Recordset
            If um.State Then um.Close
            um.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_id=rd.resc_id and rm.resc_code='" & frm_estpost.textresccode.Text & "' and rd.dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not um.EOF Then
            cbo_uom.Text = um!resc_uom
            End If
um.Close

End Sub

Private Sub cbo_jobcharge_Click()
On Error Resume Next



kl = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
nj = Split(frm_estpost.cbo_pproj.Text, "  -  ", Len(frm_estpost.cbo_pproj.Text), vbTextCompare)

Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl(0) & "' and rd.dresc_year='" & frm_estpost.cbo_year.Text & "' and dresc_proj='" & nj(0) & "' ", Cn, 3, 2
        If Not fl.EOF Then
        frm_estpost.textrescname.Text = fl!resc_desc
        frm_estpost.textresccode.Text = fl!resc_code
        frm_estpost.textprojkey.Text = nj(0)
        frm_estpost.txt_projdesc.Text = nj(1)
        frm_estpost.txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
        frm_estpost.textcosttype.Text = "E"
        frm_estpost.txt_vendor.Text = fl!resc_vendorcode
        frm_estpost.txt_respcode.Text = fl!resc_respcode
        Dim rr As New ADODB.Recordset
        If rr.State Then rr.Close
        rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        frm_estpost.txt_respname.Text = rr(0)
        End If
        frm_estpost.Text3.Text = fl!dresc_curcy
        
         
        End If
fl.Close

Dim fl1 As New ADODB.Recordset
        If fl1.State Then fl1.Close
        fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and rd.dresc_ratetype='CR' and rm.resc_code='" & kl(0) & "' and rd.dresc_year='" & frm_estpost.cbo_year.Text & "' and dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not fl1.EOF Then
            frm_estpost.txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
            frm_estpost.Text4.Text = fl1!dresc_curcy
            End If
        fl1.Close
''''''''''''
spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
spl1 = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)
cbo_type.Clear
cbo_type.Text = "A"
Dim ty As New ADODB.Recordset
If ty.State Then ty.Close
ty.Open "select Distinct(prgs_type) from progressdurationdetails where prgs_spread_code='" & spl(0) & "' and prgs_job_key ='" & spl1(0) & "' order by prgs_type ", Cn, 3, 2
While Not ty.EOF
cbo_type.AddItem ty(0)
ty.MoveNext
Wend
ty.Close
'cbo_costcode.ListIndex = 0
cbo_type.ListIndex = 0


nn = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)
nnm = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)

 If cbo_type.Text = "" Then cbo_type.Text = "A"
 
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from progressdurationdetails where prgs_spread_code='" & nnm(0) & "' and prgs_job_key='" & nn(0) & "' and prgs_type='" & cbo_type.Text & "'", Cn, 3, 2
If Not bd.EOF Then
DTP_sd.Value = Format(bd!prgs_startdate, "dd-MM-yyyy H:mm:ss")
DTP_ed.Value = Format(bd!prgs_enddate, "dd-MM-yyyy H:mm:ss")
Else
DTP_sd.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
DTP_ed.Value = Format(Date, "dd-MM-yyyy H:mm:ss")
End If


Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then
lbl.Visible = True
Check1.Visible = True
If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
   txt_days.Enabled = True
   txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                   c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
                   
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_days.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_edays.Text
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                    
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
Dim cc As New ADODB.Recordset
If cc.State Then cc.Close
cc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & frm_estpost.textresccode.Text & "' ", Cn, 3, 2
While Not cc.EOF
cbo_costcode.AddItem cc(0) & "  -  " & cc(1)
cc.MoveNext
Wend
cc.Close

Dim cr As New ADODB.Recordset
            If cr.State Then cr.Close
            cr.Open "select DISTINCT(dresc_curcy) from resourcedetails where  dresc_code='" & frm_estpost.textresccode.Text & "' and dresc_proj='" & nj(0) & "' and  dresc_year='" & frm_estpost.cbo_year.Text & "' order by  dresc_curcy", Cn, 3, 2
            If Not cr.EOF Then
            cbo_curr.Text = cr(0)
            End If
cr.Close


Dim um As New ADODB.Recordset
            If um.State Then um.Close
            um.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_id=rd.resc_id and rm.resc_code='" & frm_estpost.textresccode.Text & "' and rd.dresc_proj='" & nj(0) & "'", Cn, 3, 2
            If Not um.EOF Then
            cbo_uom.Text = um!resc_uom
            End If
um.Close
End Sub
Private Sub cbo_spread_Change()
On Error Resume Next
cbo_type.Clear
spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
If cbo_spread.Text = "NA  -  Not Applicable" Then

          
                                
            cbo_type.Text = "-"
            cbo_tranx.Enabled = True
            cbo_tranx.Clear
            cbo_tranx.Text = "ME"
            cbo_tranx.AddItem "ME"
            cbo_tranx.AddItem "AJ"
            txt_days.Text = ""
            txt_days.Enabled = False
            txt_edays.Enabled = False
            cbo_jobcharge.Enabled = True
Else
            cbo_tranx.Text = "SD"
            cbo_type.Text = "A"
            txt_days.Enabled = False
             txt_edays.Enabled = False
             
            cbo_tranx.Enabled = False
End If
On Error Resume Next

'spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
spl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
If spl(0) = "" Then Exit Sub
If spl1(0) = "" Then Exit Sub
 

 
Dim ty As New ADODB.Recordset
If ty.State Then ty.Close
ty.Open "select Distinct(prgs_type) from progressdurationdetails where prgs_spread_code='" & spl(0) & "' and prgs_job_key ='" & spl1(0) & "' order by prgs_type ", Cn, 3, 2
While Not ty.EOF
cbo_type.AddItem ty(0)
ty.MoveNext
Wend
ty.Close
End Sub

Private Sub cbo_spread_Click()
On Error Resume Next
cbo_type.Clear
spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
If cbo_spread.Text = "NA  -  Not Applicable" Then

           
                                
              cbo_type.Text = "-"
            cbo_tranx.Enabled = True
            cbo_tranx.Clear
            cbo_tranx.Text = "ME"
            cbo_tranx.AddItem "ME"
            cbo_tranx.AddItem "AJ"
            txt_days.Text = ""
            txt_days.Enabled = False
            txt_edays.Enabled = False
            cbo_jobcharge.Enabled = True
Else
            cbo_tranx.Text = "SD"
              cbo_type.Text = "A"
            txt_days.Enabled = False
             txt_edays.Enabled = False
             
            cbo_tranx.Enabled = False
End If
On Error Resume Next

'spl = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
spl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
If spl(0) = "" Then Exit Sub
If spl1(0) = "" Then Exit Sub
 

Dim ty As New ADODB.Recordset
If ty.State Then ty.Close
ty.Open "select Distinct(prgs_type) from progressdurationdetails where prgs_spread_code='" & spl(0) & "' and prgs_job_key ='" & spl1(0) & "' order by prgs_type ", Cn, 3, 2
While Not ty.EOF
cbo_type.AddItem ty(0)
ty.MoveNext
Wend
ty.Close
End Sub

Private Sub cbo_type_Click()
On Error Resume Next
nn = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
nnm = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
 
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from progressdurationdetails where prgs_spread_code='" & nnm(0) & "' and prgs_job_key='" & nn(0) & "' and prgs_type='" & cbo_type.Text & "'", Cn, 3, 2
If Not bd.EOF Then
DTP_sd.Value = bd!prgs_startdate
DTP_ed.Value = bd!prgs_enddate
End If


Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then
lbl.Visible = True
Check1.Visible = True
If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
            
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                   c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
                   
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_days.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_edays.Text
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                    
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If

End Sub

Private Sub Check1_Click()
Exit Sub

Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then


lbl.Visible = True
Check1.Visible = True


If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
 txt_days.Enabled = True
 txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                    c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 txt_days.Enabled = False
 txt_edays.Enabled = False
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
                   
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_qty.Text
                                    txt_edays.Text = 0
                                    txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_qty.Text
                                    txt_edays.Text = ""
                                    txt_days.Text = 0
                                    txt_tqty.Text = 0
                    End If
                    
                    
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
txt_days.Enabled = True
txt_edays.Enabled = True
End If


End Sub

Private Sub DTP_ed_Change()
Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then


lbl.Visible = True
Check1.Visible = True


If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
  txt_days.Enabled = True
  txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                   c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
  txt_days.Enabled = False
  txt_edays.Enabled = False
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_qty.Text
                                    txt_edays.Text = 0
                                    txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_qty.Text
                                    txt_edays.Text = ""
                                    txt_days.Text = 0
                                    txt_tqty.Text = 0
                    End If
                    
                 
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
End Sub

Private Sub DTP_ed_Click()
Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then


lbl.Visible = True
Check1.Visible = True


If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
    txt_days.Enabled = True
    txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                   c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
  txt_days.Enabled = False
  txt_edays.Enabled = False
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_qty.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_qty.Text
                                    txt_edays.Text = ""
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                 
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
End Sub



Private Sub DTP_sd_Change()
Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then


lbl.Visible = True
Check1.Visible = True


If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
  txt_days.Enabled = True
  txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    Else
                    a = 0
                    c = DTP_ed.Value - DTP_sd.Value
                    End If
                    
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
     txt_days.Enabled = False
     txt_edays.Enabled = False
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_qty.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_qty.Text
                                    txt_edays.Text = ""
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                 
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
End Sub

Private Sub DTP_sd_Click()
Dim a As Double
a = 0
Dim c As Double
c = 0
If cbo_spread.Text = "NA  -  Not Applicable" Then


lbl.Visible = True
Check1.Visible = True


If Check1.Value = 1 Then
DTP_ed.Enabled = True
DTP_sd.Enabled = True
       txt_days.Enabled = True
       txt_edays.Enabled = True
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                    a = DTP_ed.Value - DTP_sd.Value
                    c = 0
                    ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
                    a = main.DTPcutdate1 - DTP_sd.Value
                    c = DTP_ed.Value - main.DTPcutdate1
                    
                    Else
                    a = 0
                   c = DTP_ed.Value - DTP_sd.Value
                    End If
                    txt_days.Text = a
                    txt_edays.Text = c
            
 ElseIf Check1.Value = 0 Then
 DTP_ed.Enabled = False
 DTP_sd.Enabled = True
 DTP_ed.Value = DTP_sd.Value
                   
                    If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                    txt_tqty.Text = txt_qty.Text
                                    txt_edays.Text = 0
                                     txt_etqty.Text = 0
                       Else
                       
                                    txt_etqty.Text = txt_qty.Text
                                    txt_edays.Text = ""
                                    txt_days.Text = 0
                                     txt_tqty.Text = 0
                    End If
                    
                 
 End If

Else
DTP_ed.Enabled = False
DTP_sd.Enabled = False
lbl.Visible = False
Check1.Visible = False
        If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
        a = DTP_ed.Value - DTP_sd.Value
        c = 0
        ElseIf DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value >= main.DTPcutdate1 Then
        a = main.DTPcutdate1 - DTP_sd.Value
        c = DTP_ed.Value - main.DTPcutdate1
        
        Else
        a = 0
       c = DTP_ed.Value - DTP_sd.Value
        End If
txt_days.Text = a
txt_edays.Text = c
End If
End Sub


Private Sub Form_Load()
'Unload frm_progresstranx
On Error Resume Next
Call connect
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
txt_unitrate.Text = frm_estpost.txt_brate.Text
yy = Split(frm_estpost.cbo_pproj.Text, "  -  ", Len(frm_estpost.cbo_pproj.Text), vbTextCompare)
 ass = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)
cbo_spread.AddItem "NA  -  Not Applicable"
Dim tr As New ADODB.Recordset
            If tr.State Then tr.Close
            tr.Open "select DISTINCT(p.prgs_spread_code),s.spread_desc   from progressdurationdetails p,spreadmaster s where p.prgs_spread_code=s.spread_code and p.prgs_job_key='" & ass(0) & "' order by prgs_spread_code", Cn, 3, 2
            While Not tr.EOF
            cbo_spread.AddItem tr(0) & "  -  " & tr(1)
            
            tr.MoveNext
            Wend
tr.Close
aas = Split(frm_estpost.cbo_resc.Text, "  -  ", Len(frm_estpost.cbo_resc.Text), vbTextCompare)

 



Dim cr As New ADODB.Recordset
            If cr.State Then cr.Close
            cr.Open "select DISTINCT(dresc_curcy) from resourcedetails where  dresc_code='" & frm_estpost.textresccode.Text & "' and dresc_proj='" & yy(0) & "' and  dresc_year='" & frm_estpost.cbo_year.Text & "' order by  dresc_curcy", Cn, 3, 2
            If Not cr.EOF Then
            cbo_curr.Text = cr(0)
            End If
cr.Close


Dim um As New ADODB.Recordset
            If um.State Then um.Close
            um.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_id=rd.resc_id and rm.resc_code='" & frm_estpost.textresccode.Text & "' and rd.dresc_proj='" & yy(0) & "'", Cn, 3, 2
            If Not um.EOF Then
            cbo_uom.Text = um!resc_uom
            End If
um.Close

Dim cc As New ADODB.Recordset
            If cc.State Then cc.Close
            cc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & frm_estpost.textresccode.Text & "' ", Cn, 3, 2
            While Not cc.EOF
            cbo_costcode.AddItem cc(0) & "  -  " & cc(1)
            cc.MoveNext
            Wend
cc.Close


 Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from currencymaster order by cur_currency", Cn, 3, 2
While Not cr2.EOF
cbo_curr.AddItem cr2!cur_currency
cr2.MoveNext
Wend
cr2.Close

Dim um1 As New ADODB.Recordset
If um1.State Then um1.Close
um1.Open "select * from uom order by uom_uom", Cn, 3, 2
While Not um1.EOF
cbo_uom.AddItem um1!uom_uom
um1.MoveNext
Wend
um1.Close

cbo_tranx.AddItem "SD"
cbo_tranx.AddItem "ME"
cbo_tranx.AddItem "AJ"
Call jobch
End Sub

Private Sub txt_days_Change()
If txt_days.Text = "" Then Exit Sub
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_sd.Value)

dt1 = CDbl(txt_days.Text)
dt2 = dt + dt1
DTP_ed.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")

txt_tqty.Text = txt_qty.Text * txt_days.Text

txt_Xrate_Change
End Sub

Private Sub txt_days_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_sd.Value)
dt1 = CDbl(txt_days.Text)
dt2 = dt + dt1
DTP_ed.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")
End Sub

Private Sub txt_edays_Change()
If txt_days.Text = "" Then Exit Sub
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_sd.Value)
dt1 = CDbl(txt_edays.Text)
dt2 = dt + dt1
If txt_edays.Text = 0 Then
Else
DTP_ed.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")
End If
txt_etqty.Text = txt_qty.Text * txt_edays.Text

End Sub

Private Sub txt_edays_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(DTP_sd.Value)
dt1 = CDbl(txt_edays.Text)
dt2 = dt + dt1
DTP_ed.Value = Format(dt2, "dd-MM-yyyy H:mm:ss")
End Sub

Private Sub txt_etqty_Change()
On Error Resume Next
txt_ectcamt.Text = (txt_etqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
End Sub

Private Sub txt_qty_Change()
On Error Resume Next
If cbo_spread.Text = "NA  -  Not Applicable" Then
        If Check1.Value = False Then
        
                         If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                      txt_tqty.Text = txt_qty.Text
                                      txt_edays.Text = 0
                                      txt_etqty.Text = 0
                         Else
                         
                                      txt_etqty.Text = txt_qty.Text
                                      txt_edays.Text = ""
                                      txt_days.Text = 0
                                      txt_tqty.Text = 0
                         End If
          Else
          txt_tqty.Text = txt_qty * txt_days.Text
          txt_etqty.Text = txt_qty.Text * txt_edays.Text
          End If
   Else
   txt_tqty.Text = txt_qty * txt_days.Text
   txt_etqty.Text = txt_qty.Text * txt_edays.Text
   End If
txt_Xrate_Change
End Sub

Private Sub txt_qty_KeyPress(KeyAscii As Integer)
On Error Resume Next
If cbo_spread.Text = "NA  -  Not Applicable" Then
        If Check1.Value = False Then
        
                         If DTP_sd.Value <= main.DTPcutdate1 And DTP_ed.Value <= main.DTPcutdate1 Then
                                      txt_tqty.Text = txt_qty.Text
                                      txt_edays.Text = 0
                                       txt_etqty.Text = 0
                         Else
                         
                                      txt_etqty.Text = txt_qty.Text
                                      txt_edays.Text = ""
                                      txt_days.Text = 0
                                       txt_tqty.Text = 0
                         End If
          Else
          txt_tqty.Text = txt_qty * txt_days.Text
          
          End If
   Else
   txt_tqty.Text = txt_qty * txt_days.Text
   txt_etqty.Text = txt_qty.Text * txt_edays.Text
   End If
txt_Xrate_Change
End Sub

Private Sub txt_unitrate_Change()
On Error Resume Next
txt_Extdamt.Text = (txt_tqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
'txt_etqty.Text = (txt_qty.Text) * (txt_edays.Text)
txt_ectcamt.Text = (txt_etqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
End Sub

Private Sub txt_unitrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = (txt_tqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
'txt_etqty.Text = (txt_qty.Text) * (txt_edays.Text)
txt_ectcamt.Text = (txt_etqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
End Sub

Private Sub txt_Xrate_Change()
On Error Resume Next
txt_Extdamt.Text = (txt_tqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)
'txt_etqty.Text = (txt_qty.Text) * (txt_edays.Text)
txt_ectcamt.Text = (txt_etqty.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)


End Sub


Public Sub jobch()
jh = Split(frm_estpost.cbo_pproj.Text, "  -  ", Len(frm_estpost.cbo_pproj.Text), vbTextCompare)
cbo_jobcharge.Clear
                    Dim fl1 As New ADODB.Recordset
                        If fl1.State Then fl1.Close
                        fl1.Open "select DISTINCT(resc_code),resc_desc from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code   and rd.dresc_year='" & frm_estpost.cbo_year.Text & "' and dresc_proj='" & jh(0) & "' ", Cn, 3, 2
                        While Not fl1.EOF
                         cbo_jobcharge.AddItem fl1(0) & "  -  " & fl1(1)
                        fl1.MoveNext
                        Wend
                                
                                
                                
                                Dim bs As New ADODB.Recordset
                                If bs.State Then bs.Close
                                bs.Open "select Distinct(resp_code) from responsibledetails order by resp_code", Cn, 3, 2
                                While Not bs.EOF
                                cbo_obs.AddItem bs(0)
                                bs.MoveNext
                                Wend
                                bs.Close
End Sub

