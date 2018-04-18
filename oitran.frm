VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form oitran 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Other Transactions"
      TabPicture(0)   =   "oitran.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "oitran.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   7695
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MISC"
            Height          =   855
            Left            =   1920
            TabIndex        =   37
            Top             =   4080
            Width           =   5295
            Begin VB.TextBox txt_ytd 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   40
               Text            =   "0"
               Top             =   435
               Width           =   1455
            End
            Begin VB.TextBox txt_ctd 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   2040
               TabIndex        =   39
               Text            =   "0"
               Top             =   435
               Width           =   1455
            End
            Begin VB.TextBox txt_chg 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   3720
               TabIndex        =   38
               Text            =   "0"
               Top             =   435
               Width           =   1455
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "YTD-Last Mth End"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "YTD-Current Month"
               Height          =   195
               Left            =   2040
               TabIndex        =   42
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Changes Current Mth"
               Height          =   195
               Left            =   3720
               TabIndex        =   41
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "EAC"
            Height          =   855
            Left            =   120
            TabIndex        =   34
            Top             =   4080
            Width           =   1695
            Begin VB.TextBox txt_eac 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   50
               TabIndex        =   35
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Est. At Completion"
               Height          =   195
               Left            =   45
               TabIndex        =   36
               Top             =   240
               Width           =   1290
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BDGT"
            Height          =   1455
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   3615
            Begin VB.TextBox txt_rateaft 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   50
               Text            =   "0"
               Top             =   1035
               Width           =   1575
            End
            Begin VB.TextBox txt_adjbl 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   50
               TabIndex        =   48
               Text            =   "0"
               Top             =   1035
               Width           =   1575
            End
            Begin VB.TextBox txt_rateb4 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   46
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.TextBox txt_bdgt 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   50
               TabIndex        =   32
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate/Day Aft BdgtAdj"
               Height          =   195
               Left            =   1800
               TabIndex        =   51
               Top             =   840
               Width           =   1545
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Adj on B/L Budget"
               Height          =   195
               Left            =   45
               TabIndex        =   49
               Top             =   840
               Width           =   1320
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate/Day B4 BdgtAdj"
               Height          =   195
               Left            =   1800
               TabIndex        =   47
               Top             =   240
               Width           =   1545
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BaseLine Budget"
               Height          =   195
               Left            =   45
               TabIndex        =   33
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BCWP"
            Height          =   1455
            Left            =   3840
            TabIndex        =   20
            Top             =   720
            Width           =   3375
            Begin VB.TextBox txt_bcwp 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   25
               Text            =   "0"
               Top             =   1035
               Width           =   1455
            End
            Begin VB.TextBox txt_bcwpdays 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   23
               Text            =   "0"
               Top             =   1035
               Width           =   1575
            End
            Begin VB.TextBox txt_bcwpbl 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   21
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BCWP"
               Height          =   195
               Left            =   1800
               TabIndex        =   26
               Top             =   840
               Width           =   480
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ActualDays@CutOff"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   840
               Width           =   1425
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate/Day B4 BdgtAdj"
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   1545
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ECTC"
            Height          =   1935
            Left            =   3840
            TabIndex        =   15
            Top             =   2160
            Width           =   3375
            Begin VB.TextBox txt_adjustment 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   52
               Text            =   "0"
               Top             =   915
               Width           =   1455
            End
            Begin VB.TextBox txt_etc 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               TabIndex        =   27
               Text            =   "0"
               Top             =   1515
               Width           =   1455
            End
            Begin VB.TextBox txt_etcbl 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   50
               TabIndex        =   17
               Text            =   "0"
               Top             =   915
               Width           =   1575
            End
            Begin VB.TextBox txt_etcdays 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   50
               TabIndex        =   16
               Text            =   "0"
               Top             =   1515
               Width           =   1575
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Adjustment"
               Height          =   195
               Left            =   1800
               TabIndex        =   53
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ECTC"
               Height          =   195
               Left            =   1800
               TabIndex        =   28
               Top             =   1320
               Width           =   420
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate/Day Aft BdgtAdj"
               Height          =   195
               Left            =   45
               TabIndex        =   19
               Top             =   720
               Width           =   1545
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EstDays To Complete"
               Height          =   195
               Left            =   45
               TabIndex        =   18
               Top             =   1320
               Width           =   1530
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ACWP"
            Height          =   1935
            Left            =   120
            TabIndex        =   8
            Top             =   2160
            Width           =   3615
            Begin MSComCtl2.DTPicker dtp_asat 
               Height          =   330
               Left            =   1845
               TabIndex        =   44
               Top             =   435
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               Format          =   50462721
               CurrentDate     =   38140
            End
            Begin VB.TextBox txt_acwp 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   1800
               TabIndex        =   29
               Text            =   "0"
               Top             =   1515
               Width           =   1575
            End
            Begin VB.TextBox txt_acwpadj 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   45
               TabIndex        =   13
               Text            =   "0"
               Top             =   1515
               Width           =   1575
            End
            Begin VB.TextBox txt_acwpbl 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Left            =   50
               TabIndex        =   11
               Top             =   1035
               Width           =   1575
            End
            Begin VB.TextBox txt_acwpacc 
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "###,###,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   50
               TabIndex        =   9
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "As At:"
               Height          =   195
               Left            =   1845
               TabIndex        =   45
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ACWP"
               Height          =   195
               Left            =   1800
               TabIndex        =   30
               Top             =   1320
               Width           =   480
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Days Btw Rpt/CutOff"
               Height          =   195
               Left            =   45
               TabIndex        =   14
               Top             =   1320
               Width           =   1500
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate/Day Aft BdgtAdj"
               Height          =   195
               Left            =   50
               TabIndex        =   12
               Top             =   840
               Width           =   1545
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Per Accounts"
               Height          =   195
               Left            =   45
               TabIndex        =   10
               Top             =   240
               Width           =   960
            End
         End
         Begin VB.ComboBox txt_tranx 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4695
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   5160
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   50462721
            CurrentDate     =   37987
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TranX Desc"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   5160
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   1230
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   7695
         Begin VB.TextBox txt_notes 
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   240
            Width           =   6375
         End
      End
   End
End
Attribute VB_Name = "oitran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ppr As Double
Public ppsd As Date
Public pped As Date
Public ppcd As Date
Public dys As Double
 
Private Sub dtp_asat_Change()
On Error Resume Next
dys = 0
dys = main.DTPcutdate1.Value - dtp_asat.Value
 txt_acwpadj.Text = dys
txt_acwpbl.Text = txt_rateaft.Text
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
ppcd = main.DTPcutdate1.Value
 
End Sub

Private Sub dtp_asat_Click()
On Error Resume Next
dys = 0
dys = main.DTPcutdate1.Value - dtp_asat.Value
  txt_acwpadj.Text = dys
txt_acwpbl.Text = txt_rateaft.Text
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
ppcd = main.DTPcutdate1.Value
 
 
End Sub

Private Sub Form_Load()
On Error Resume Next
 dys = 0
 fab = 1
Dim ot As New ADODB.Recordset
If ot.State Then ot.Close
ot.Open "select * from othertransaction order by ot_tranx", Cn, 3, 2
While Not ot.EOF
txt_tranx.AddItem ot!ot_tranx & "  -  " & ot!ot_desc
ot.MoveNext
Wend

Unload parameters
ppr = 0
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from parameters", Cn, 3, 2
If Not pr.EOF Then
ppr = pr!p_ydays
ppsd = pr!p_sdate
pped = pr!p_edate
'ppcd = pr!p_cdate
End If
ppcd = main.DTPcutdate1.Value
 
   dys = 0
            dys = main.DTPcutdate1.Value - dtp_asat.Value
           
 
 
dtp_asat.Value = Format(Date, "dd/MM/yyyy")
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
fab = 0
End Sub

Private Sub txt_acwp_Change()
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
txt_ctd.Text = Format(Round(CDbl(txt_acwp.Text), 2), "###,###,##0")
End Sub

Private Sub txt_acwp_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
End Sub

Private Sub txt_acwpacc_Change()
On Error Resume Next
 

txt_acwpadj.Text = dys
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

Unload parameters
End Sub

Private Sub txt_acwpacc_KeyPress(KeyAscii As Integer)
On Error Resume Next
 txt_acwpadj.Text = dys
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

Unload parameters
End Sub

Private Sub txt_acwpadj_Change()
On Error Resume Next
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub

Private Sub txt_acwpadj_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")

 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

End Sub

Private Sub txt_acwpbl_Change()
On Error Resume Next
On Error Resume Next
         dys = 0
        dys = main.DTPcutdate1.Value - dtp_asat.Value
        
'txt_acwpbl.Text = Round((CDbl(txt_bdgt.Text) / ppr) * dys, 2)
txt_acwp.Text = Format(Round(CDbl(txt_acwpacc.Text) + CDbl(txt_acwpbl.Text * txt_acwpadj.Text), 2), "###,###,##0")
 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

End Sub

Private Sub txt_adjbl_Change()
On Error Resume Next
txt_rateaft.Text = Format(Round((CDbl(txt_bdgt.Text) + CDbl(txt_adjbl.Text)) / ppr, 2), "###,###,##0.00")
txt_acwpbl.Text = txt_rateaft.Text
txt_etcbl.Text = txt_rateaft.Text
End Sub

Private Sub txt_adjbl_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_rateaft.Text = Format(Round((CDbl(txt_bdgt.Text) + CDbl(txt_adjbl.Text)) / ppr, 2), "###,###,##0.00")
txt_acwpbl.Text = txt_rateaft.Text
txt_etcbl.Text = txt_rateaft.Text
End Sub
 

Private Sub txt_adjustment_Change()
On Error Resume Next
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub

Private Sub txt_adjustment_KeyPress(KeyAscii As Integer)
On Error Resume Next
If IsNumeric(txt_adjustment.Text) And txt_adjustment.Text <> "" Then
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End If
End Sub
Private Sub txt_bcwpbl_Change()
On Error Resume Next
txt_bcwp.Text = Format(Round(CDbl(txt_bcwpbl.Text) * CDbl(txt_bcwpdays.Text), 2), "###,###,##0")
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub
Private Sub txt_bcwpdays_Change()
On Error Resume Next
txt_bcwp.Text = Format(Round(CDbl(txt_bcwpbl.Text) * CDbl(txt_bcwpdays.Text), 2), "###,###,##0")
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub
Private Sub txt_bdgt_Change()
On Error Resume Next
                dys = 0
                dys = main.DTPcutdate1.Value - dtp_asat.Value
txt_rateb4.Text = Format(Round(CDbl(txt_bdgt.Text) / ppr, 2), "###,###,##0.00")
txt_rateaft.Text = Format(Round((CDbl(txt_bdgt.Text) + CDbl(txt_adjbl.Text)) / ppr, 2), "###,###,##0.00")
txt_bcwpbl.Text = txt_rateb4.Text
        If (ppcd - ppsd) < 0 Then
        txt_bcwpdays.Text = 0
        Else
        txt_bcwpdays.Text = Format(CDbl(ppcd - ppsd), "###,###,##0.00")
        End If
        txt_bcwpbl.Text = txt_rateb4.Text
txt_bcwp.Text = Format(Round(CDbl(txt_bcwpbl.Text) * CDbl(txt_bcwpdays.Text), 2), "###,###,##0")
txt_etcbl = Format(Round(CDbl(txt_bdgt.Text) / ppr, 2), "###,###,##0.00")

txt_acwpbl.Text = txt_rateaft.Text
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
Unload parameters
End Sub

Private Sub txt_bdgt_KeyPress(KeyAscii As Integer)
On Error Resume Next
                dys = 0
                dys = main.DTPcutdate1.Value - dtp_asat.Value
               

 
txt_rateb4.Text = Format(Round(CDbl(txt_bdgt.Text) / ppr, 2), "###,###,##0.00")
txt_rateaft.Text = Format(Round((CDbl(txt_bdgt.Text) + CDbl(txt_adjbl.Text)) / ppr, 2), "###,###,##0.00")
txt_bcwpbl.Text = txt_rateb4.Text
        If (ppcd - ppsd) < 0 Then
        txt_bcwpdays.Text = 0
        Else
        txt_bcwpdays.Text = Format(CDbl(ppcd - ppsd), "###,###,##0.00")
        End If
txt_bcwp.Text = Format(Round(CDbl(txt_bcwpbl.Text) * CDbl(txt_bcwpdays.Text), 2), "###,###,##0")
txt_etcbl = Format(Round(CDbl(txt_bdgt.Text) / ppr, 2), "###,###,##0.00")

txt_acwpbl.Text = txt_rateaft.Text
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
Unload parameters
End Sub

Private Sub txt_ctd_Change()
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
End Sub

Private Sub txt_ctd_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
End Sub

Private Sub txt_etc_Change()
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub

Private Sub txt_etc_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub

Private Sub txt_etcadj_Change()
On Error Resume Next
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub

Private Sub txt_etcadj_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")
End Sub

Private Sub txt_etcbl_Change()
On Error Resume Next
If (pped - ppcd) < 0 Then
txt_etcdays.Text = 0
Else
txt_etcdays.Text = Format(Round(CDbl(pped - ppcd), 2), "###,###,##0.00")
End If
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

Unload parameters
End Sub

Private Sub txt_etcdays_Change()
On Error Resume Next
 
txt_etc.Text = Format(Round((CDbl(txt_etcbl.Text) + CDbl(txt_adjustment.Text)) * (CDbl(txt_etcdays.Text)), 2), "###,###,##0")

End Sub

Private Sub txt_rateaft_Change()
On Error Resume Next
txt_acwpbl.Text = txt_rateaft.Text
txt_etcbl.Text = txt_rateaft.Text
End Sub

Private Sub txt_rateb4_Change()
On Error Resume Next
txt_bcwpbl.Text = txt_rateb4.Text
End Sub

Private Sub txt_ytd_Change()
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
End Sub

Private Sub txt_ytd_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
End Sub

