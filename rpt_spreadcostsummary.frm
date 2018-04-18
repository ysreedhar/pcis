VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_spreadcostsummary 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   11100
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   11055
      ExtentX         =   19500
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.ComboBox cbo_type 
         Height          =   315
         Left            =   9600
         TabIndex        =   11
         Text            =   "A"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Calculate"
         Height          =   255
         Left            =   9600
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   6960
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1205
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ComboBox cbo_spread 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   5175
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   7680
         Picture         =   "rpt_spreadcostsummary.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to Print"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   6840
         Picture         =   "rpt_spreadcostsummary.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Click to View"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8520
         Picture         =   "rpt_spreadcostsummary.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to Exit"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   6840
         Top             =   120
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1695
         Left            =   75
         Top             =   120
         Width           =   6615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Spread"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jobcharge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   915
      End
   End
End
Attribute VB_Name = "rpt_spreadcostsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbo_spread_Click()
spp = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
    Dim rc As New ADODB.Recordset
    If rc.State Then rc.Close
    rc.Open "select DISTINCT(c.bd_jobcharge),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and c.bd_spread = '" & spp(0) & "'   order by c.bd_jobcharge", Cn, 3, 2
    While Not rc.EOF
    List1.AddItem rc(0) & "  -  " & rc(1)
    rc.MoveNext
    Wend
    rc.Close
End Sub
Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmd_print_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT:
End Sub
Private Sub cmd_show_Click()
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
If cbo_spread.Text = "" Then
MsgBox "Select Spread"
End If
 If Check1.Value = 1 Then
 Call cuttoffdatechange
 End If
Call nocolor
Unload frmBusy
End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "SPREAD COST SUMMARY"
Me.Top = 10
Me.Left = 10
Me.Height = 9720
Me.Width = 11220
WebBrowser.Navigate "About:Blank"
    Dim tr As New ADODB.Recordset
    If tr.State Then tr.Close
    tr.Open "select DISTINCT(p.prgs_spread_code),s.spread_desc   from progressdurationdetails p,spreadmaster s where p.prgs_spread_code=s.spread_code order by prgs_spread_code", Cn, 3, 2
    While Not tr.EOF
    cbo_spread.AddItem tr(0) & "  -  " & tr(1)
    tr.MoveNext
    Wend
    tr.Close
    Dim ty As New ADODB.Recordset
    If ty.State Then ty.Close
    ty.Open "select DISTINCT(prgs_type) from progressdurationdetails order by prgs_type", Cn, 3, 2
    While Not ty.EOF
    cbo_type.AddItem ty(0)
    ty.MoveNext
    Wend
End Sub
Public Sub nocolor()
'On Error Resume Next
Dim fs As Object
Dim fso As New FileSystemObject
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
   'fs.WriteLine "      BORDER-BOTTOM: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-LEFT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-RIGHT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-TOP: Wheat 1px solid"
   fs.WriteLine "    }"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bolder;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "    }"
   fs.WriteLine "    .TrFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "   }"
   fs.WriteLine "</style>"
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
Dim cnt As Integer
 RPTHEADING fs
 cnt = 0
Dim acwp As Double
Dim ectc As Double
Dim eac As Double
Dim bdgt As Double
Dim ntotacwp As Double
Dim ntotectc As Double
Dim ntoteac As Double
Dim ntotbdgt As Double
Dim dblBudgetTotalRM As Double
Dim dblBudgetTotalUSD As Double
Dim dblACWPTotalRM As Double
Dim dblACWPTotalUSD As Double

ntotacwp = 0: ntotectc = 0: ntoteac = 0: ntotbdgt = 0
sp = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
Dim rsXChgRate As New ADODB.Recordset
Dim dblXRate As Double
If rsXChgRate.State Then rsXChgRate.Close
rsXChgRate.Open "select Cur_xchgrate from currencymaster where cur_currency = 'USD' order by u_date desc", Cn, 3, 2
If rsXChgRate.EOF Then dblXRate = 1 Else dblXRate = CDbl(rsXChgRate(0))
rsXChgRate.Close
        Dim l As Integer
        l = 0
        For l = 0 To List1.ListCount - 1
        If List1.Selected(l) = True Then
        nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=6 align=left><b>" & List1.List(l) & "</b></td></tr>"
Dim totacwp As Double
Dim totectc As Double
Dim toteac As Double
Dim totbdgt As Double
totacwp = 0: totectc = 0: toteac = 0: totbdgt = 0
        Dim k As Integer
        k = 0
        For k = 0 To List2.ListCount - 1
        If List2.Selected(k) = True Then
        jk = Split(List2.List(k), "  -  ", Len(List2.List(k)), vbTextCompare)

acwp = 0: ectc = 0: eac = 0: bdgt = 0
Dim rsACWP As New ADODB.Recordset
If rsACWP.State Then rsACWP.Close
'rsACWP.Open "select bd_extdamt,bd_e_extdamt,bd_tqty,bd_e_tqty,bd_xchg,bd_unitrate from cost where bd_spread ='" & sp(0) & "' and bd_type='" & cbo_type.Text & "' and bd_jobcharge='" & nm(0) & "' and bd_costtype='E' and bd_resccode='" & jk(0) & "'", Cn, 3, 2

'eac = acwp + ectc
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select bd_unitrate,bd_curr ,bd_xchg,  bd_costcode, cc.cc_desc,bd_uom, bd_qty,bd_extdamt, bd_tqty from cost, costcode cc where bd_spread ='" & sp(0) & "' and bd_type='" & cbo_type.Text & "' and bd_jobcharge='" & nm(0) & "' and bd_costtype='B' and bd_resccode='" & jk(0) & "' and bd_costcode = cc.cc_code", Cn, 3, 2
While Not rs1.EOF
'If rs1("bd_UOM") = "MT" Then
If rs1("bd_Curr") <> "RM" Then
bdgt = (rs1("bd_qty") * rs1("bd_unitrate"))
Else
bdgt = (rs1("bd_qty") * rs1("bd_unitrate") / rs1("bd_xchg"))
End If
If rsACWP.State Then rsACWP.Close
rsACWP.Open "select ((bd_qty* bd_unitrate)/ bd_xchg),(bd_qty* bd_unitrate),bd_uom,(bd_extdamt/bd_xchg),bd_curr from cost where bd_spread ='" & sp(0) & "' and bd_jobcharge='" & nm(0) & "' and bd_costtype='E' and bd_type='" & cbo_type.Text & "' and bd_resccode='" & jk(0) & "' and bd_curr = '" & rs1("bd_curr") & "' and bd_costcode = '" & rs1("bd_costcode") & "'", Cn, 3, 2
If Not rsACWP.EOF Then
'If rsACWP("bd_uom") = "MT" Then
If rsACWP("bd_Curr") <> "RM" Then
acwp = rsACWP(1)
Else
acwp = rsACWP(0)
End If
End If
  cnt = cnt + 1 '********************************
If cnt >= 42 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs
cnt = 0
End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If strCode <> Mid(List2.List(k), 2, 3) Then
fs.WriteLine "            <td Nowrap colspan=6 align=left><b>" & Format(List2.List(k), vbUpperCase) & "</b></td>"
fs.WriteLine "        </tr>"
End If
strCode = Mid(List2.List(k), 2, 3)
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & rs1(3) & "</td>"
fs.WriteLine "            <td  >" & rs1(4) & "</td>"
If rs1(1) = "RM" Then
fs.WriteLine "            <td Nowrap align=right>" & Format(bdgt, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(0, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(acwp, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(0, "###,###,###,##0") & "</td>"
dblBudgetTotalRM = dblBudgetTotalRM + bdgt
dblACWPTotalRM = dblACWPTotalRM + acwp
Else
fs.WriteLine "            <td Nowrap align=right>" & Format(0, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(bdgt, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(0, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(acwp, "###,###,###,##0") & "</td>"
dblBudgetTotalUSD = dblBudgetTotalUSD + bdgt
dblACWPTotalUSD = dblACWPTotalUSD + acwp
End If
fs.WriteLine "        </tr>"
rs1.MoveNext
Wend
totbdgt = totbdgt + bdgt
totacwp = totacwp + acwp
totectc = totectc + ectc
toteac = toteac + eac
End If
Next k
  cnt = cnt + 1 '********************************
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs
cnt = 0
End If
'fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
'fs.WriteLine "            <td  colspan=2 ><font color=white>" & List1.List(l) & "</td>"
'fs.WriteLine "            <td Nowrap align=right colspan=2><font color=white>" & Format(totbdgt, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td Nowrap align=right colspan=2><font color=white>" & Format(totacwp, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(totectc, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(toteac, "###,###,###,##0") & "</td>"
'fs.WriteLine "        </tr>"
ntotbdgt = ntotbdgt + totbdgt
ntotacwp = ntotacwp + totacwp
ntotectc = ntotectc + totectc
ntoteac = ntoteac + toteac
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td   colspan=2><font color=white>Total In RM and USD Portion</td>"
fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(dblBudgetTotalRM, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(dblBudgetTotalUSD, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=white>" & Format(dblACWPTotalRM, "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=white>" & Format(dblACWPTotalUSD, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntotectc, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntoteac, "###,###,###,##0") & "</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=white>Total in RM (Convert USD Portion to RM)  </td>"
fs.WriteLine "            <td colspan=2 Nowrap align=right><font color=white>" & Format((dblBudgetTotalRM + (dblBudgetTotalUSD * dblXRate)), "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right colspan=2><font color=white>" & Format((dblACWPTotalRM + (dblACWPTotalUSD * dblXRate)), "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntotectc, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntoteac, "###,###,###,##0") & "</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td   colspan=2><font color=white>Total in USD (Convert RM Portion to USD)</td>"
fs.WriteLine "            <td colspan=2 Nowrap align=right><font color=white>" & Format(((dblBudgetTotalRM / dblXRate) + dblBudgetTotalUSD), "###,###,###,##0") & "</td>"
fs.WriteLine "            <td Nowrap align=right colspan=2><font color=white>" & Format(((dblACWPTotalRM / dblXRate) + dblACWPTotalUSD), "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntotectc, "###,###,###,##0") & "</td>"
'fs.WriteLine "            <td  Nowrap align=right><font color=white>" & Format(ntoteac, "###,###,###,##0") & "</td>"
fs.WriteLine "        </tr>"
dblBudgetTotalRM = 0: dblBudgetTotalUSD = 0: dblACWPTotalRM = 0: dblACWPTotalUSD = 0
End If
Next l
  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub

Private Sub List1_ItemCheck(Item As Integer)
List2.Clear
nn = Split(cbo_spread.Text, "  -  ", Len(cbo_spread.Text), vbTextCompare)
nm = Split(List1.List(Item), "  -  ", Len(List1.List(Item)), vbTextCompare)
Dim rcs As New ADODB.Recordset
If rcs.State Then rcs.Close
rcs.Open "select DISTINCT(bd_resccode) from cost where  bd_spread='" & nn(0) & "' order by bd_resccode", Cn, 3, 2
While Not rcs.EOF
Dim rcd As New ADODB.Recordset
            If rcd.State Then rcd.Close
            rcd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rcs(0) & "' ", Cn, 3, 2
                   If Not rcd.EOF Then
                   List2.AddItem rcs(0) & "  -  " & rcd(0)
                   Else
                   List2.AddItem rcs(0)
                   End If
 rcs.MoveNext
Wend
 Dim j As Integer
 j = 0
 For j = 0 To List2.ListCount - 1
 List2.Selected(j) = True
 Next j
End Sub
Private Sub Option1_Click()
Dim f As Integer
f = 0
For f = 0 To List1.ListCount - 1
List1.Selected(f) = True
Next f
End Sub
Private Sub Option2_Click()
Dim f1 As Integer
f1 = 0
For f1 = 0 To List1.ListCount - 1
List1.Selected(f1) = False
Next f1
End Sub
Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=3><b>" & GetCompanyName & "</td>"
fs.WriteLine "            <td colspan=3><b>SPREAD COST SUMMARY</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
fs.WriteLine "           <td colspan=3 ><b>" & cbo_spread.Text & "</td>"
fs.WriteLine "           <td  colspan=3>Cuttoff Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap align=center colspan=2 ><font color=white>&nbsp;</td>"
fs.WriteLine "            <td Nowrap align=center colspan=2><font color=white>BDGT</td>"
fs.WriteLine "            <td Nowrap align=center colspan=2><font color=white>ACWP</td>"
'fs.WriteLine "            <td Nowrap align=center colspan=3><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan =2 align=center><font color=white>Resc</td>"
fs.WriteLine "            <td Nowrap align=center><font color=white>RM</td>"
fs.WriteLine "            <td Nowrap align=center><font color=white>USD</td>"
fs.WriteLine "            <td Nowrap align=center><font color=white>RM</td>"
fs.WriteLine "            <td Nowrap align=center><font color=white>USD</td>"
'fs.WriteLine "            <td Nowrap align=center><font color=white>ECTC</td>"
'fs.WriteLine "            <td Nowrap align=center><font color=white>EAC</td>"
fs.WriteLine "        </tr>"
End Sub
Public Sub cuttoffdatechange()
Dim j As Integer
j = 0
For j = 0 To List1.ListCount - 1
If List1.Selected(j) = True Then
xk = Split(List1.List(j), "  -  ", Len(List1.List(j)), vbTextCompare)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where bd_jobcharge='" & xk(0) & "' and bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2
While Not fldata.EOF
iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_JobCharge, "  -  ", Len(fldata!bd_JobCharge), vbTextCompare)
Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If
Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
fldata2.Open "select * from cost where    bd_jobcharge='" & fldata!bd_JobCharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'
    If Not fldata2.EOF Then
            fldata2!bd_sdate = dt1
            fldata2!bd_edate = dt2
                    If dt1 <= main.DTPcutdate1.Value And dt2 <= main.DTPcutdate1.Value Then
                    a = dt2 - dt1
                    c = 0
                    ElseIf dt1 <= main.DTPcutdate1.Value And dt2 >= main.DTPcutdate1.Value Then
                    a = main.DTPcutdate1.Value - dt1
                    c = dt2 - main.DTPcutdate1.Value
                    Else
                    a = 0
                    c = dt2 - dt1
                    End If
            Dim d As Double
            d = 0
            Dim f As Double
            f = 0
            fldata2!bd_days = a
            fldata2!bd_e_days = c
            d = CDbl(a) * CDbl(fldata!bd_qty)
            fldata2!bd_e_tqty = CDbl(c) * CDbl(fldata!bd_qty)
            fldata2!bd_tqty = d
            fldata2!bd_extdamt = CDbl(d) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
            fldata2!bd_e_extdamt = CDbl(fldata2!bd_e_tqty) * CDbl(fldata!bd_unitrate) * CDbl(fldata!bd_xchg)
            fldata2.Update
    End If
        fldata.MoveNext
    Wend
End If
Next j
End Sub
