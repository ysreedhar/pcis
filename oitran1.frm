VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form oitran1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recovery Costs"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7575
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
      TabPicture(0)   =   "oitran1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "oitran1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00DC7E5A&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   -75000
         TabIndex        =   19
         Top             =   300
         Width           =   7695
         Begin VB.TextBox txt_notes 
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   20
            Top             =   240
            Width           =   6375
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   7695
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MISC"
            Height          =   1095
            Left            =   2160
            TabIndex        =   23
            Top             =   3960
            Width           =   4815
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
               Left            =   1680
               TabIndex        =   28
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
               Left            =   3240
               TabIndex        =   25
               Text            =   "0"
               Top             =   435
               Width           =   1455
            End
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
               TabIndex        =   24
               Text            =   "0"
               Top             =   435
               Width           =   1455
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "YTD-Current Month"
               Height          =   195
               Left            =   1680
               TabIndex        =   29
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Changes Current Mth"
               Height          =   195
               Left            =   3240
               TabIndex        =   27
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "YTD-Last Mth End"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1320
            End
         End
         Begin VB.TextBox txt_costcode 
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
            TabIndex        =   21
            Top             =   270
            Width           =   5775
         End
         Begin VB.ComboBox txt_tranx 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   5775
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ACWP"
            Height          =   1095
            Left            =   120
            TabIndex        =   14
            Top             =   2640
            Width           =   2775
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
               Left            =   120
               TabIndex        =   15
               Text            =   "0"
               Top             =   555
               Width           =   1575
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ACWP"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   480
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ECTC"
            Height          =   1095
            Left            =   3120
            TabIndex        =   11
            Top             =   2640
            Width           =   2775
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
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Text            =   "0"
               Top             =   555
               Width           =   1695
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ECTC"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   660
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BCWP"
            Height          =   1095
            Left            =   3120
            TabIndex        =   8
            Top             =   1320
            Width           =   2775
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
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Text            =   "0"
               Top             =   435
               Width           =   1695
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BCWP"
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BDGT"
            Height          =   1095
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   2775
            Begin VB.TextBox txt_bdgt 
               BackColor       =   &H00FFFFFF&
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
               Left            =   165
               TabIndex        =   6
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BaseLine Budget"
               Height          =   195
               Left            =   165
               TabIndex        =   7
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "EAC"
            Height          =   1095
            Left            =   120
            TabIndex        =   2
            Top             =   3960
            Width           =   1935
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
               Height          =   285
               Left            =   165
               TabIndex        =   3
               Text            =   "0"
               Top             =   435
               Width           =   1575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Est. At Completion"
               Height          =   195
               Left            =   165
               TabIndex        =   4
               Top             =   240
               Width           =   1290
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CostCode"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TranX Desc"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "oitran1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim ot As New ADODB.Recordset
If ot.State Then ot.Close
ot.Open "select * from othertransaction order by ot_tranx", Cn, 3, 2
While Not ot.EOF
txt_tranx.AddItem ot!ot_tranx & "  -  " & ot!ot_desc
ot.MoveNext
Wend
End Sub

Private Sub txt_acwp_Change()
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub

Private Sub txt_acwp_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub



Private Sub txt_etc_Change()
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub

Private Sub txt_etc_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_eac.Text = Format(Round(CDbl(txt_acwp.Text) + CDbl(txt_etc.Text), 2), "###,###,##0")
End Sub

Private Sub txt_tranx_Click()
On Error Resume Next
nm = Split(txt_costcode.Text, ",", Len(txt_costcode.Text), vbTextCompare)
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select SUM(bd_extdamt),SUM(bd_bcwpamt) from cost where bd_costtype='B' and bd_costcode='" & nm(0) & "' and bd_year='" & frm_l0.cbo_year.Text & "'", Cn, 3, 2
If Not bd.EOF Then
txt_bdgt.Text = Format(bd(0), "###,###,###,###,##0")
txt_bcwp.Text = Format(bd(1), "###,###,###,###,##0")
End If

Dim es As New ADODB.Recordset
If es.State Then es.Close
es.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from cost where bd_costtype='E' and bd_costcode='" & nm(0) & "' and bd_year='" & frm_l0.cbo_year.Text & "'", Cn, 3, 2
If Not es.EOF Then
txt_acwp.Text = Format(es(0), "###,###,###,###,##0")
txt_etc.Text = Format(es(1), "###,###,###,###,##0")
txt_eac.Text = CDbl(Format(es(0), "###,###,###,###,##0")) + CDbl(Format(es(1), "###,###,###,###,##0"))

End If


End Sub

Private Sub txt_ytd_Change()
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")

End Sub

Private Sub txt_ytd_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_chg.Text = Format(Round(CDbl(txt_acwp.Text) - CDbl(txt_ytd.Text), 2), "###,###,##0")

End Sub
