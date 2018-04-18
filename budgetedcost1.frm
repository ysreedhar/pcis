VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form budgetedcost1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Budgeted Cost"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fr3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9255
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   0
         TabIndex        =   17
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
         TabPicture(0)   =   "budgetedcost1.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Notes"
         TabPicture(1)   =   "budgetedcost1.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label12"
         Tab(1).Control(1)=   "Frame7"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   0
            TabIndex        =   21
            Top             =   300
            Width           =   9015
            Begin VB.ComboBox cboChargeType 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   7440
               TabIndex        =   45
               Text            =   "XX"
               Top             =   360
               Width           =   1095
            End
            Begin VB.ComboBox cbo_obs 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4320
               TabIndex        =   43
               Text            =   "XX"
               Top             =   960
               Width           =   1095
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
               TabIndex        =   38
               Top             =   1320
               Width           =   2055
               Begin VB.TextBox txt_wrkcomp 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   15
                  Text            =   "0"
                  Top             =   600
                  Width           =   855
               End
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
                  TabIndex        =   16
                  Text            =   "0"
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.Label Label33 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   40
                  Top             =   900
                  Width           =   1215
               End
               Begin VB.Label Label32 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   39
                  Top             =   360
                  Width           =   1215
               End
            End
            Begin VB.ComboBox cbo_spread 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   360
               Width           =   4335
            End
            Begin VB.ComboBox cbo_jobcharge 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   960
               Width           =   4095
            End
            Begin VB.ComboBox cbo_tranx 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   4560
               TabIndex        =   1
               Top             =   360
               Width           =   1335
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
               TabIndex        =   22
               Top             =   1320
               Width           =   6375
               Begin VB.TextBox txt_totdays 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2340
                  TabIndex        =   7
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox txt_days 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   6
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox txt_qty 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   285
                  Left            =   120
                  TabIndex        =   5
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
                  Left            =   1110
                  TabIndex        =   10
                  Top             =   1080
                  Width           =   1575
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
                  TabIndex        =   9
                  Top             =   480
                  Width           =   1215
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
                  TabIndex        =   8
                  Top             =   480
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
                  Left            =   4560
                  TabIndex        =   14
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.TextBox txt_Xrate 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   855
               End
               Begin VB.TextBox txt_downtime 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2820
                  TabIndex        =   12
                  Text            =   "0"
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox txt_esclfactor 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   13
                  Text            =   "0"
                  Top             =   1080
                  Width           =   735
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
                  TabIndex        =   32
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00FFC0C0&
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
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   31
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   30
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label8 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   27
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label10 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   26
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   25
                  Top             =   840
                  Width           =   1335
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
                  TabIndex        =   24
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label Label31 
                  BackColor       =   &H00FFC0C0&
                  BackStyle       =   0  'Transparent
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
                  TabIndex        =   23
                  Top             =   840
                  Width           =   735
               End
            End
            Begin VB.ComboBox cbo_costcode 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   5520
               TabIndex        =   3
               Top             =   960
               Width           =   3135
            End
            Begin MSComCtl2.DTPicker DTP_tdate 
               Height          =   315
               Left            =   6000
               TabIndex        =   2
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   48955393
               CurrentDate     =   38733
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Charge Type"
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
               Left            =   7440
               TabIndex        =   46
               Top             =   120
               Width           =   930
            End
            Begin VB.Label Label13 
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
               Left            =   4320
               TabIndex        =   44
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               TabIndex        =   37
               Top             =   120
               Width           =   1935
            End
            Begin VB.Label Label26 
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
               TabIndex        =   36
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               TabIndex        =   35
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
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
               TabIndex        =   34
               Top             =   720
               Width           =   1755
            End
            Begin VB.Label Label29 
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
               Left            =   6000
               TabIndex        =   33
               Top             =   120
               Width           =   1230
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   -75000
            TabIndex        =   20
            Top             =   300
            Width           =   9015
            Begin VB.TextBox txt_notes 
               Height          =   2415
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   18
               Top             =   240
               Width           =   8295
            End
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
            TabIndex        =   42
            Top             =   0
            Width           =   855
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
            TabIndex        =   41
            Top             =   0
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "budgetedcost1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

Private Sub cbo_costcode_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

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
cr2.Open "select * from resourcedetails where dresc_code='" & frm_budgetedcost.textresccode.Text & "' and dresc_proj='" & frm_budgetedcost.textprojkey.Text & "' and dresc_curcy='" & cbo_curr.Text & "' and dresc_year='" & frm_budgetedcost.cbo_year.Text & "'", Cn, 3, 2
If Not cr2.EOF Then
'txt_unitrate.Text = cr2!dresc_rate
End If
End Sub

Private Sub cbo_curr_Click()
Dim crr As New ADODB.Recordset
If crr.State Then crr.Close
crr.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not crr.EOF Then
    txt_Xrate.Text = crr!cur_xchgrate
    End If
crr.Close
Dim crr2 As New ADODB.Recordset
If crr2.State Then crr2.Close
crr2.Open "select * from resourcedetails where dresc_code='" & frm_budgetedcost.textresccode.Text & "' and dresc_proj='" & frm_budgetedcost.textprojkey.Text & "' and dresc_curcy='" & cbo_curr.Text & "' and dresc_year='" & frm_budgetedcost.cbo_year.Text & "'", Cn, 3, 2
If Not crr2.EOF Then
'txt_unitrate.Text = crr2!dresc_rate
End If
End Sub

Private Sub cbo_curr_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub cbo_jobcharge_Change()
On Error Resume Next
                    gg = Split(frm_budgetedcost.cbo_pproj, "  -  ", Len(frm_budgetedcost.cbo_pproj), vbTextCompare)
                    jch = Split(frm_budgetedcost.cbo_resc.Text, "  -  ", Len(frm_budgetedcost.cbo_resc.Text), vbTextCompare)
                    nn = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
                    kl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
                    If cbo_spread.Text <> "NA  -  Not Applicable" Then
                    
                    Dim bd As New ADODB.Recordset
                    If bd.State Then bd.Close
                    bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & jch(0) & "' and bdgt_spread_code ='" & nn(0) & "' ", Cn, 3, 2
                    If Not bd.EOF Then
                       txt_days.Text = bd!bdgt_days
                       txt_wrkcomp.Text = bd!bdgt_per_workcomplete
                    End If
                    bd.Close
                    Else
                    
'''                    txt_days.Text = ""
                    End If
           'cbo_costcode.ListIndex = 0
            
            
                    Dim fl As New ADODB.Recordset
                    If fl.State Then fl.Close
                    fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & frm_budgetedcost.cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
                    If Not fl.EOF Then
                    frm_budgetedcost.textrescname.Text = fl!resc_desc
                    frm_budgetedcost.textresccode.Text = fl!resc_code
                    frm_budgetedcost.textprojkey.Text = gg(0)
                    frm_budgetedcost.txt_projdesc.Text = gg(1)
                    frm_budgetedcost.txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
                    frm_budgetedcost.textcosttype.Text = "B"
                    frm_budgetedcost.txt_vendor.Text = fl!resc_vendorcode
                    frm_budgetedcost.txt_respcode.Text = fl!resc_respcode
                    Dim rr As New ADODB.Recordset
                    If rr.State Then rr.Close
                    rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
                    If Not rr.EOF Then
                    frm_budgetedcost.txt_respname.Text = rr(0)
                    End If
                    frm_budgetedcost.Text3.Text = fl!dresc_curcy
                    End If
                    fl.Close
                    
                    Dim fl1 As New ADODB.Recordset
                    If fl1.State Then fl1.Close
                    fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='CR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & frm_budgetedcost.cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
                    If Not fl1.EOF Then
                    frm_budgetedcost.txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
                    frm_budgetedcost.Text4.Text = fl1!dresc_curcy
                    End If
                    fl1.Close
                    
                    
cbo_costcode.Clear
Dim cc As New ADODB.Recordset
If cc.State Then cc.Close
cc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & frm_budgetedcost.textresccode.Text & "' ", Cn, 3, 2
While Not cc.EOF
cbo_costcode.AddItem cc(0) & "  -  " & cc(1)
cc.MoveNext
Wend
cc.Close










Dim cr As New ADODB.Recordset
            If cr.State Then cr.Close
            cr.Open "select DISTINCT(dresc_curcy) from resourcedetails where  dresc_code='" & frm_budgetedcost.textresccode.Text & "' and dresc_proj='" & gg(0) & "' and  dresc_year='" & frm_budgetedcost.cbo_year.Text & "' order by  dresc_curcy", Cn, 3, 2
            If Not cr.EOF Then
            cbo_curr.Text = cr(0)
            End If
cr.Close


Dim um As New ADODB.Recordset
            If um.State Then um.Close
            um.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_id=rd.resc_id and rm.resc_code='" & frm_budgetedcost.textresccode.Text & "' and rd.dresc_proj='" & gg(0) & "'", Cn, 3, 2
            If Not um.EOF Then
            cbo_uom.Text = um!resc_uom
            End If
um.Close


                    End Sub
                    
                    Private Sub cbo_jobcharge_Click()
                    On Error Resume Next
                    gg = Split(frm_budgetedcost.cbo_pproj, "  -  ", Len(frm_budgetedcost.cbo_pproj), vbTextCompare)
                      jch = Split(frm_budgetedcost.cbo_resc.Text, "  -  ", Len(frm_budgetedcost.cbo_resc.Text), vbTextCompare)
                    nn = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
                    kl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
                    
                    If cbo_spread.Text <> "NA  -  Not Applicable" Then
                  
                    Dim bd As New ADODB.Recordset
                    If bd.State Then bd.Close
                    bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & jch(0) & "' and bdgt_spread_code ='" & nn(0) & "' ", Cn, 3, 2
                    If Not bd.EOF Then
                       txt_days.Text = bd!bdgt_days
                       txt_wrkcomp.Text = bd!bdgt_per_workcomplete
                    End If
                    bd.Close
                    Else
                    
                    txt_days.Text = ""
                    End If
            'cbo_costcode.ListIndex = 0
            
            
                    Dim fl As New ADODB.Recordset
                    If fl.State Then fl.Close
                    fl.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='BR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & frm_budgetedcost.cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
                    If Not fl.EOF Then
                    frm_budgetedcost.textrescname.Text = fl!resc_desc
                    frm_budgetedcost.textresccode.Text = fl!resc_code
                    frm_budgetedcost.textprojkey.Text = gg(0)
                    frm_budgetedcost.txt_projdesc.Text = gg(1)
                    frm_budgetedcost.txt_brate.Text = Format(fl!dresc_rate, "###,###,##0.00")
                    frm_budgetedcost.textcosttype.Text = "B"
                    frm_budgetedcost.txt_vendor.Text = fl!resc_vendorcode
                    frm_budgetedcost.txt_respcode.Text = fl!resc_respcode
                    Dim rr As New ADODB.Recordset
                    If rr.State Then rr.Close
                    rr.Open "select DISTINCT(resp_desc)  from responsiblemaster where resp_code='" & fl!resc_respcode & "' ", Cn, 3, 2
                    If Not rr.EOF Then
                    frm_budgetedcost.txt_respname.Text = rr(0)
                    End If
                    frm_budgetedcost.Text3.Text = fl!dresc_curcy
                    End If
                    fl.Close
                    
                    Dim fl1 As New ADODB.Recordset
                    If fl1.State Then fl1.Close
                    fl1.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code and dresc_ratetype='CR' and resc_code='" & kl1(0) & "' and rd.dresc_year='" & frm_budgetedcost.cbo_year.Text & "' and dresc_proj='" & gg(0) & "' ", Cn, 3, 2
                    If Not fl1.EOF Then
                    frm_budgetedcost.txt_crate.Text = Format(fl1!dresc_rate, "###,###,##0.00")
                    frm_budgetedcost.Text4.Text = fl1!dresc_curcy
                    End If
                    fl1.Close
                    
                    cbo_costcode.Clear
                    Dim ccc As New ADODB.Recordset
            If ccc.State Then ccc.Close
            ccc.Open "select DISTINCT(cc.cc_code),cc_desc from costcode cc,resourcecostcode rcc where cc.cc_id=rcc.rcc_id and rcc.rcc_resource='" & frm_budgetedcost.textresccode.Text & "' ", Cn, 3, 2
            While Not ccc.EOF
            cbo_costcode.AddItem ccc(0) & "  -  " & ccc(1)
            ccc.MoveNext
            Wend
ccc.Close



Dim crc As New ADODB.Recordset
            If crc.State Then crc.Close
            crc.Open "select DISTINCT(dresc_curcy) from resourcedetails where  dresc_code='" & frm_budgetedcost.textresccode.Text & "' and dresc_proj='" & gg(0) & "' and  dresc_year='" & frm_budgetedcost.cbo_year.Text & "' order by  dresc_curcy", Cn, 3, 2
            If Not crc.EOF Then
            cbo_curr.Text = crc(0)
            End If
crc.Close


Dim umc As New ADODB.Recordset
            If umc.State Then umc.Close
            umc.Open "select * from resourcemaster rm, resourcedetails rd where rm.resc_id=rd.resc_id and rm.resc_code='" & frm_budgetedcost.textresccode.Text & "' and rd.dresc_proj='" & gg(0) & "'", Cn, 3, 2
            If Not umc.EOF Then
            cbo_uom.Text = umc!resc_uom
            End If
umc.Close


End Sub


Private Sub cbo_jobcharge_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_spread_Change()
If cbo_spread.Text <> "NA  -  Not Applicable" Then
            cbo_tranx.Clear
            cbo_tranx.AddItem "SD"
            cbo_tranx.AddItem "ME"
            cbo_tranx.AddItem "AJ"
Else
            cbo_tranx.Clear
            cbo_tranx.AddItem "ME"
            cbo_tranx.AddItem "AJ"
            cbo_tranx.Text = "ME"
End If
                gg = Split(frm_budgetedcost.cbo_pproj, "  -  ", Len(frm_budgetedcost.cbo_pproj), vbTextCompare)
                jch = Split(frm_budgetedcost.cbo_resc.Text, "  -  ", Len(frm_budgetedcost.cbo_resc.Text), vbTextCompare)
                nn = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
                kl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
Dim bddd As New ADODB.Recordset
If bddd.State Then bddd.Close
                    bddd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & jch(0) & "' and bdgt_spread_code ='" & nn(0) & "' ", Cn, 3, 2
                    If Not bddd.EOF Then
                       txt_days.Text = bddd!bdgt_days
                       txt_wrkcomp.Text = bddd!bdgt_per_workcomplete
                 
                    
                    Else
                    
''                    txt_days.Text = ""
                    End If
                    bddd.Close
End Sub

Private Sub cbo_spread_Click()

                If cbo_spread.Text <> "NA  -  Not Applicable" Then
                cbo_tranx.Clear
                cbo_tranx.AddItem "SD"
                cbo_tranx.AddItem "ME"
                cbo_tranx.AddItem "AJ"
                cbo_tranx.Text = "SD"
                Else
                cbo_tranx.Clear
                cbo_tranx.AddItem "ME"
                cbo_tranx.AddItem "AJ"
                cbo_tranx.Text = "ME"
                End If

          
                        gg = Split(frm_budgetedcost.cbo_pproj, "  -  ", Len(frm_budgetedcost.cbo_pproj), vbTextCompare)
                jch = Split(frm_budgetedcost.cbo_resc.Text, "  -  ", Len(frm_budgetedcost.cbo_resc.Text), vbTextCompare)
                nn = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
                kl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
Dim bdd As New ADODB.Recordset
If bdd.State Then bdd.Close
                    bdd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & jch(0) & "' and bdgt_spread_code ='" & nn(0) & "' ", Cn, 3, 2
                    If Not bdd.EOF Then
                       txt_days.Text = bdd!bdgt_days
                       txt_wrkcomp.Text = bdd!bdgt_per_workcomplete
                   
                    
                    Else
                    
                    txt_days.Text = ""
                    End If
bdd.Close
End Sub


Private Sub cbo_spread_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub cbo_tranx_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub cbo_uom_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
Call connect
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
txt_unitrate.Text = frm_budgetedcost.txt_brate.Text
nf = Split(frm_budgetedcost.cbo_pproj.Text, "  -  ", Len(frm_budgetedcost.cbo_pproj.Text), vbTextCompare)
aas = Split(frm_budgetedcost.cbo_resc.Text, "  -  ", Len(frm_budgetedcost.cbo_resc.Text), vbTextCompare)

cbo_spread.AddItem "NA  -  Not Applicable"

'populate Charge Type
cboChargeType.Clear
    Dim rsChargeType As New ADODB.Recordset
    If rsChargeType.State Then rsChargeType.Close
    rsChargeType.Open "select chargeType from tblChargeType", Cn, 3, 2
    While Not rsChargeType.EOF
    cboChargeType.AddItem rsChargeType(0)
    rsChargeType.MoveNext
    Wend
rsChargeType.Close

Dim tr As New ADODB.Recordset
            If tr.State Then tr.Close
            tr.Open "select DISTINCT(b.bdgt_spread_code),s.spread_desc  from budgeteddurationdetails b , spreadmaster s where b.bdgt_spread_code=s.spread_code and b.bdgt_job_key= '" & aas(0) & "'order by bdgt_spread_code", Cn, 3, 2
            While Not tr.EOF
            cbo_spread.AddItem tr(0) & "  -  " & tr(1)
            tr.MoveNext
            Wend
tr.Close





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
jh = Split(frm_budgetedcost.cbo_pproj.Text, "  -  ", Len(frm_budgetedcost.cbo_pproj.Text), vbTextCompare)
cbo_jobcharge.Clear
                    Dim fl1 As New ADODB.Recordset
                        If fl1.State Then fl1.Close
                        fl1.Open "select DISTINCT(resc_code),resc_desc from resourcemaster rm, resourcedetails rd where rm.resc_code=rd.dresc_code   and rd.dresc_year='" & frm_budgetedcost.cbo_year.Text & "' and dresc_proj='" & jh(0) & "' ", Cn, 3, 2
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

Private Sub txt_bcwpamt_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub txt_days_Change()
On Error Resume Next
If txt_days.Text <> "" Then
   txt_totdays.Text = txt_qty.Text * txt_days.Text
Else
   txt_totdays.Text = txt_qty.Text
End If
txt_Xrate_Change
End Sub

Private Sub txt_days_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = validatechk(KeyAscii, 11)
If txt_days.Text <> "" Then
   txt_totdays.Text = txt_qty.Text * txt_days.Text
Else
   txt_totdays.Text = txt_qty.Text
End If
txt_Xrate_Change
End Sub


Private Sub txt_downtime_Change()
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)
End Sub

Private Sub txt_downtime_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)
End Sub

Private Sub txt_esclfactor_Change()
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)
End Sub

Private Sub txt_esclfactor_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)
End Sub

Private Sub txt_Extdamt_Change()
On Error Resume Next
txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)
End Sub

Private Sub txt_Extdamt_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub txt_qty_Change()
On Error Resume Next

If txt_days.Text = "" Then
   txt_totdays.Text = txt_qty.Text
Else
   txt_totdays.Text = txt_qty * txt_days.Text
End If
txt_Xrate_Change
End Sub
Private Sub txt_qty_KeyPress(KeyAscii As Integer)
On Error Resume Next
If txt_days.Text = "" Then
   txt_totdays.Text = txt_qty.Text
Else
End If
txt_Xrate_Change
End Sub

Private Sub txt_totdays_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub txt_unitrate_Change()
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)
txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)
End Sub

Private Sub txt_unitrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)

txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)
End Sub

Private Sub txt_wrkcomp_Change()
On Error Resume Next
txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)

txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)
End Sub

Private Sub txt_wrkcomp_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub

Private Sub txt_Xrate_Change()
On Error Resume Next
txt_Extdamt.Text = (((txt_totdays.Text) * (txt_Xrate.Text) * (txt_unitrate.Text)) * ((100 + txt_downtime.Text)) / 100) * ((100 + txt_esclfactor.Text) / 100)

txt_bcwpamt.Text = (txt_Extdamt.Text) * (txt_wrkcomp.Text / 100)
End Sub

Private Sub txt_Xrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = 0
End Sub
