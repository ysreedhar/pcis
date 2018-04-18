VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_obsestimate 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6495
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   11456
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
      Height          =   1455
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9600
         Picture         =   "rpt_obsestimate.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Click to Exit"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_obsestimate.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Click to View"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9960
         Picture         =   "rpt_obsestimate.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Click to Print"
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   480
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   3375
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   120
         Width           =   5055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   735
         Left            =   120
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DC7E5A&
         Caption         =   "    Project key - Description"
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
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11055
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACWP"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ECTC"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EAC"
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transaction Dates"
         Height          =   255
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   8520
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   49283075
         CurrentDate     =   38099
      End
   End
End
Attribute VB_Name = "rpt_obsestimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer
Private Sub cbo_proj_Click()
Option1.Value = False
List1.Clear
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim rc As New ADODB.Recordset
 If rc.State Then rc.Close
 rc.Open "select DISTINCT(c.bd_obs),r.resp_desc from cost c,responsibledetails r where c.bd_obs=r.resp_code and c.bd_projectkey='" & nn(0) & "'", Cn, 3, 2
 While Not rc.EOF
 List1.AddItem rc(0) & "  -  " & rc(1)
 rc.MoveNext
 Wend
 rc.Close
  ' Option1.Value = True
   Check1.Value = 1
        hgg = 0
         For hgg = 0 To List1.ListCount - 1
         List1.Selected(hgg) = False
         Next hgg
         Option1.Value = 0
         Option2.Value = 0
End Sub
Private Sub Check1_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
End If
End Sub
Private Sub Check2_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
End If
End Sub
Private Sub Check3_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
End If
End Sub
Private Sub Check4_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
End If
End Sub
Private Sub Check5_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
End If
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
If cbo_proj.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
DTPicker1_Change
If Check5.Value = 1 Then
Call appcolor
Else
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy
End If
End Sub
Private Sub DTPicker1_Change()
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
                    If dt1 <= DTPicker1.Value And dt2 <= DTPicker1.Value Then
                    a = dt2 - dt1
                    c = 0
                    ElseIf dt1 <= DTPicker1.Value And dt2 >= DTPicker1.Value Then
                    a = DTPicker1.Value - dt1
                    c = dt2 - DTPicker1.Value
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
Dim jj As Integer
jj = 0
For jj = 0 To List1.ListCount - 1
If List1.Selected(jj) = True Then
xk = Split(List1.List(jj), "  -  ", Len(List1.List(jj)), vbTextCompare)
Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where  bd_jobcharge='" & xk(0) & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
While Not cd.EOF
If cd!bd_chk = 1 Then
If cd!bd_sdate <= DTPicker1.Value And cd!bd_edate <= DTPicker1.Value Then
                    a = cd!bd_edate - cd!bd_sdate
                    c = 0
                    ElseIf cd!bd_sdate <= DTPicker1.Value And cd!bd_edate >= DTPicker1.Value Then
                    a = DTPicker1.Value - cd!bd_sdate
                    c = cd!bd_edate - DTPicker1.Value
                    Else
                    a = 0
                    c = cd!bd_edate - cd!bd_sdate
                    End If
                    cd!bd_days = a
                    cd!bd_e_days = c
                    If IsNull(cd!bd_days) = True Then
                    cd!bd_tqty = cd!bd_qty
                    Else
                    cd!bd_tqty = cd!bd_qty * cd!bd_days
                    End If
                    cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
                    If IsNull(cd!bd_e_days) = True Then
                    cd!bd_e_tqty = cd!bd_qty
                    Else
                    cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
                    End If
                    cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
 ElseIf cd!bd_chk = 0 Then
 If cd!bd_chk1 = 0 Then
cd!bd_edate = cd!bd_sdate
If cd!bd_sdate <= DTPicker1.Value And cd!bd_edate <= DTPicker1.Value Then
cd!bd_tqty = cd!bd_qty
cd!bd_days = Null
cd!bd_e_days = 0
cd!bd_e_tqty = 0
Else
cd!bd_e_tqty = cd!bd_qty
cd!bd_e_days = Null
cd!bd_days = 0
cd!bd_tqty = 0
End If
If IsNull(cd!bd_days) = True Then
cd!bd_tqty = cd!bd_qty
Else
cd!bd_tqty = cd!bd_qty * cd!bd_days
End If
cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
If IsNull(cd!bd_e_days) = True Then
cd!bd_e_tqty = cd!bd_qty
Else
cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
End If
cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
ElseIf cd!bd_chk1 = 1 Then
cd!bd_tqty = cd!bd_qty * cd!bd_days
cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
End If
End If
cd.Update
cd.MoveNext
Wend
End If
Next jj
End Sub
Private Sub DTPicker1_Click()
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
fldata2.Open "select * from cost where   bd_jobcharge='" & fldata!bd_JobCharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

If Not fldata2.EOF Then



fldata2!bd_sdate = dt1
fldata2!bd_edate = dt2
If dt1 <= DTPicker1.Value And dt2 <= DTPicker1.Value Then
a = dt2 - dt1
c = 0
ElseIf dt1 <= DTPicker1.Value And dt2 >= DTPicker1.Value Then
a = DTPicker1.Value - dt1
c = dt2 - DTPicker1.Value

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


Dim jj As Integer
jj = 0
For jj = 0 To List1.ListCount - 1
If List1.Selected(jj) = True Then
xk = Split(List1.List(jj), "  -  ", Len(List1.List(jj)), vbTextCompare)
Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where  bd_jobcharge='" & xk(0) & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
While Not cd.EOF


If cd!bd_chk = 1 Then


If cd!bd_sdate <= DTPicker1.Value And cd!bd_edate <= DTPicker1.Value Then
a = cd!bd_edate - cd!bd_sdate
c = 0
ElseIf cd!bd_sdate <= DTPicker1.Value And cd!bd_edate >= DTPicker1.Value Then
a = DTPicker1.Value - cd!bd_sdate
c = cd!bd_edate - DTPicker1.Value

Else
a = 0
c = cd!bd_edate - cd!bd_sdate
End If
cd!bd_days = a
cd!bd_e_days = c
If IsNull(cd!bd_days) = True Then
cd!bd_tqty = cd!bd_qty
Else
cd!bd_tqty = cd!bd_qty * cd!bd_days
End If
cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
If IsNull(cd!bd_e_days) = True Then
cd!bd_e_tqty = cd!bd_qty
Else
cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
End If
cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
ElseIf cd!bd_chk = 0 Then

cd!bd_edate = cd!bd_sdate
If cd!bd_sdate <= DTPicker1.Value And cd!bd_edate <= DTPicker1.Value Then
            cd!bd_tqty = cd!bd_qty
            cd!bd_days = Null
            cd!bd_e_days = 0
            cd!bd_e_tqty = 0
Else
            cd!bd_e_tqty = cd!bd_qty
            cd!bd_e_days = Null
            cd!bd_days = 0
            cd!bd_tqty = 0
End If
If IsNull(cd!bd_days) = True Then
cd!bd_tqty = cd!bd_qty
Else
cd!bd_tqty = cd!bd_qty * cd!bd_days
End If
cd!bd_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_tqty
If IsNull(cd!bd_e_days) = True Then
cd!bd_e_tqty = cd!bd_qty
Else
cd!bd_e_tqty = cd!bd_e_days * cd!bd_qty
End If
cd!bd_e_extdamt = cd!bd_unitrate * cd!bd_xchg * cd!bd_e_tqty
End If
cd.Update
cd.MoveNext
Wend
End If
Next jj
End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "EIC BY OBS"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pk.EOF
cbo_proj.AddItem pk(0) & "  -  " & pk(1)
pk.MoveNext
Wend
pk.Close
DTPicker1.Value = Now
hgg = 0
For hgg = 0 To List1.ListCount - 1
List1.Selected(hgg) = False
Next hgg
Option1.Value = False
Option2.Value = True
Me.Width = 11415
Me.Height = 9750
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List1.ListCount - 1
List1.Selected(f) = True
Next f
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List1.ListCount - 1
List1.Selected(g) = False
Next g
End If
List1.Enabled = True
End Sub
Public Sub appcolor()
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
   fs.WriteLine "           <font size=2.5 COLOR= BLUE face=Arial Narrow>" & GetCompanyName & "</font></font><br> "
   fs.WriteLine "        <font COLOR= BLUE size=2>ESTIMATED INCURRED COST BY JOBCHARGE</font>"
 fs.WriteLine "    <table border=1 cellspacing=1 bgcolor=blue width=95%>"
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=white>OBS Code</font></td>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>Description</font></td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"
Else
fs.WriteLine "            <td colspan=10><font color=white>Description</font></td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor =white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=blue>Resc Cde</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>CostCode</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>SprdCde</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>TrnxType</font> </td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=blue>Start Date</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>End Date</font> </td>"
End If
fs.WriteLine "            <td Nowrap><font color=blue>Total Qty</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>UOM</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>Curcy</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>UnitRate</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>Xrate</font> </td>"
'   fs.WriteLine "           <td Nowrap>DT</td>"
'   fs.WriteLine "           <td Nowrap>Escl</td>"
If Check1.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=blue>ACWP Amt(RM)</font> </td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=blue>Tot Qty</font> </td>"
fs.WriteLine "            <td Nowrap><font color=blue>ECTC Amt(RM)</font> </td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=blue>EAC Amt(RM)</font> </td>"
End If
fs.WriteLine "            <td ><font color=blue>Notes</font> </td>"
fs.WriteLine "        </tr>"
'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "
Dim stot As Double
Dim tot As Double
Dim tot1 As Double
Dim dtot As Double
Dim atot As Double
Dim ktot As Double
Dim wtot As Double
Dim wtot1 As Double
Dim wtot2 As Double
   wtot2 = 0
  tot = 0:  tot1 = 0
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=white>" & nm(0) & "</td>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=brown>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=10><font color=white>" & nm(1) & "</td>"
End If
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_obs='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='E' ", Cn, 3, 2
 dtot = 0
ktot = 0
wtot1 = 0
While Not fl.EOF
yre = fl(0)
stot = 0
atot = 0
wtot = 0
    Dim fldata1 As New ADODB.Recordset
    If fldata1.State Then fldata1.Close
    fldata1.Open "select * from cost  where bd_costtype='E' and bd_obs='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
    
    
    While Not fldata1.EOF
    fs.WriteLine "        <tr bgcolor=white class=TableFont>"
    
    fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_resccode & "</font> </td>"
    fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_costcode & "</font> </td>"
    fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_spread & "</font> </td>"
    fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_tranx & "</font> </td>"
    If Check4.Value = 1 Then
    fs.WriteLine "            <td Nowrap><font color=blue>" & Format(fldata1!bd_sdate, "dd/MM/yyyy") & "</font> </td>"
    fs.WriteLine "            <td Nowrap><font color=blue>" & Format(fldata1!bd_edate, "dd/MM/yyyy") & "</font> </td>"
    
    End If
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</font> </td>"
    fs.WriteLine "            <td Nowrap  ><font color=blue>" & fldata1!bd_uom & "</td>"
    fs.WriteLine "            <td Nowrap ><font color=blue>" & fldata1!bd_curr & "</td>"
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</font> </td>"
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</font> </td>"
    'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
    'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
    If Check1.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</font> </td>"
    stot = stot + fldata1!bd_extdamt
    End If
    If Check2.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_e_tqty, "###,###,##0.00") & "</font> </td>"
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "</font> </td>"
    atot = atot + fldata1!bd_e_extdamt
    End If
    If Check3.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "</font> </td>"
    wtot = wtot + (fldata1!bd_extdamt) + (fldata1!bd_e_extdamt)
    End If
    If fldata1!bd_notes <> "" Then
    fs.WriteLine "            <td ><font color=blue>" & fldata1!bd_notes & "</font> </td>"
    Else
    fs.WriteLine "            <td Nowrap><font color=blue>&nbsp;</font> </td>"
    End If
    fs.WriteLine "       </tr>"
    fldata1.MoveNext
    Wend
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
''fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal -  " & yre & "</font></td>"
Else
fs.WriteLine "            <td  colspan=9><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal -  " & yre & "</font></td>"
End If
 If Check1.Value = 1 Then
fs.WriteLine "            <td align=right ><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
End If
 If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(atot, "###,###,##0.00") & "</td>"
End If
 If Check3.Value = 1 Then
fs.WriteLine "            <td align=right ><font color=brown>" & Format(wtot, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
dtot = dtot + stot
ktot = ktot + atot
wtot1 = wtot1 + wtot
fl.MoveNext
Wend
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
    If Check4.Value = 1 Then
    fs.WriteLine "            <td  colspan=11><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total   - " & List1.List(l) & "</font></td>"
    Else
    fs.WriteLine "            <td  colspan=9><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total   - " & List1.List(l) & "</font></td>"
    End If
         If Check1.Value = 1 Then
        fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</font></td>"
        End If
             If Check2.Value = 1 Then
            fs.WriteLine "            <td  align=right>&nbsp;</td>"
            fs.WriteLine "            <td align=right ><font color=brown>" & Format(ktot, "###,###,##0.00") & "</font></td>"
            End If
                 If Check3.Value = 1 Then
                fs.WriteLine "            <td align=right ><font color=brown>" & Format(wtot1, "###,###,##0.00") & "</font></td>"
                End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
 tot = tot + dtot
 tot1 = tot1 + ktot
 wtot2 = wtot2 + wtot1
End If
Next l
fs.WriteLine "        <tr bgcolor=yellow height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11>NET TOTAL</td>"
Else
fs.WriteLine "            <td  colspan=9>NET TOTAL</td>"
End If
If Check1.Value = 1 Then
fs.WriteLine "            <td  align=right>" & Format(tot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td  align=right>" & Format(tot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right>" & Format(wtot2, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
    fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub
Public Sub nocolor()
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
   fs.WriteLine "           <font size=2.5 face=Arial Narrow>" & GetCompanyName & "</font><br> "
   fs.WriteLine "        <font size=2>ESTIMATED INCURRED COST BY JOBCHARGE</font>"
 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
fs.WriteLine "        <tr bgcolor=#acacac  height=17 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=black>OBS Code</td>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=black>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>Description</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=black>Description</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>Description</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=black>Description</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>Description</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>Description</td>"
Else
fs.WriteLine "            <td colspan=10><font color=black>Description</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor =white height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap> Resc Cde  </td>"
   fs.WriteLine "            <td Nowrap> CostCode  </td>"
   fs.WriteLine "            <td Nowrap> SprdCde </td>"
   fs.WriteLine "            <td Nowrap> TrnxType </td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap> Start Date  </td>"
   fs.WriteLine "            <td Nowrap> End Date  </td>"
   End If
   fs.WriteLine "            <td Nowrap> Total Qty  </td>"
   fs.WriteLine "            <td Nowrap> UOM  </td>"
   fs.WriteLine "            <td Nowrap> Curcy  </td>"
   fs.WriteLine "            <td Nowrap> UnitRate  </td>"
   fs.WriteLine "            <td Nowrap> Xrate  </td>"
'   fs.WriteLine "            <td Nowrap>DT</td>"
'   fs.WriteLine "            <td Nowrap>Escl</td>"
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap> ACWP Amt(RM)</font> </td>"
   End If
   If Check2.Value = 1 Then
   fs.WriteLine "            <td Nowrap> Tot Qty </td>"
   fs.WriteLine "            <td Nowrap> ECTC Amt(RM)  </td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap> EAC Amt(RM)  </td>"
   End If
   fs.WriteLine "            <td > Notes </td>"
   fs.WriteLine "        </tr>"
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "
Dim stot As Double
Dim tot As Double
Dim tot1 As Double
Dim dtot As Double
Dim atot As Double
Dim ktot As Double
Dim wtot As Double
Dim wtot1 As Double
Dim wtot2 As Double
wtot = 0: wtot1 = 0: wtot2 = 0
atot = 0: ktot = 0
stot = 0: tot = 0: dtot = 0: tot1 = 0

Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=#acacac  height=17 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=black>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=black>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>" & nm(1) & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=black>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=black>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=black>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=10><font color=black>" & nm(1) & "</td>"
End If
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_obs='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='E' ", Cn, 3, 2
dtot = 0
ktot = 0
wtot1 = 0
While Not fl.EOF
yre = fl(0)
      stot = 0
      atot = 0
      wtot = 0
    Dim fldata1 As New ADODB.Recordset
    If fldata1.State Then fldata1.Close
    fldata1.Open "select * from cost  where bd_costtype='E' and bd_obs='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
    stot = 0
    While Not fldata1.EOF
    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
    fs.WriteLine "            <td Nowrap> " & fldata1!bd_resccode & " </td>"
    fs.WriteLine "            <td Nowrap> " & fldata1!bd_costcode & " </td>"
    fs.WriteLine "            <td Nowrap> " & fldata1!bd_spread & " </td>"
    fs.WriteLine "            <td Nowrap> " & fldata1!bd_tranx & " </td>"
    If Check4.Value = 1 Then
    fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_sdate, "dd/MM/yyyy") & " </td>"
    fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_edate, "dd/MM/yyyy") & " </td>"
    
    End If
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_tqty, "###,###,##0.00") & " </td>"
    fs.WriteLine "            <td Nowrap  > " & fldata1!bd_uom & "</td>"
    fs.WriteLine "            <td Nowrap > " & fldata1!bd_curr & "</td>"
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_unitrate, "###,###,##0.00") & "  </td>"
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_xchg, "###,###,##0.00") & " </td>"
    'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
    'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
    If Check1.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_extdamt, "###,###,##0.00") & " </td>"
    stot = stot + fldata1!bd_extdamt
    End If
    If Check2.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_tqty, "###,###,##0.00") & " </td>"
    fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "  </td>"
    atot = atot + fldata1!bd_e_extdamt
    End If
    If Check3.Value = 1 Then
    fs.WriteLine "            <td Nowrap align=right> " & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "  </td>"
    wtot = wtot + (fldata1!bd_extdamt) + (fldata1!bd_e_extdamt)
    End If
    If fldata1!bd_notes <> "" Then
    fs.WriteLine "            <td > " & fldata1!bd_notes & " </td>"
    Else
    fs.WriteLine "            <td Nowrap> &nbsp;  </td>"
    End If
    fs.WriteLine "       </tr>"
    fldata1.MoveNext
    Wend
fs.WriteLine "        <tr bgcolor=white height=17 class=TableFont>"
''fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11><b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal  - " & yre & " </td>"
Else
fs.WriteLine "            <td  colspan=9><b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal   - " & yre & " </td>"
End If
 If Check1.Value = 1 Then
fs.WriteLine "            <td align=right ><b> " & Format(stot, "###,###,##0.00") & "</td>"
End If
 If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right ><b> " & Format(atot, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td align=right ><b> " & Format(wtot, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
dtot = dtot + stot
ktot = ktot + atot
wtot1 = wtot1 + wtot
fl.MoveNext
Wend
fs.WriteLine "        <tr bgcolor=white height=17 class=TableFont>"
'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
    If Check4.Value = 1 Then
    fs.WriteLine "            <td  colspan=11><b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total   - " & List1.List(l) & " </td>"
    Else
    fs.WriteLine "            <td  colspan=9><b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total  - " & List1.List(l) & " </td>"
    End If
    If Check1.Value = 1 Then
    fs.WriteLine "            <td align=right ><b> " & Format(dtot, "###,###,##0.00") & " </td>"
        End If
             If Check2.Value = 1 Then
            fs.WriteLine "            <td  align=right>&nbsp;</td>"
            fs.WriteLine "            <td align=right ><b> " & Format(ktot, "###,###,##0.00") & " </td>"
            End If
                 If Check3.Value = 1 Then
                fs.WriteLine "            <td align=right ><b> " & Format(wtot1, "###,###,##0.00") & " </td>"
                End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
 tot = tot + dtot
 tot1 = tot1 + ktot
 wtot2 = wtot2 + wtot1
End If
Next l
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11><font color=black>NET TOTAL</td>"
Else
fs.WriteLine "            <td  colspan=9><font color=black>NET TOTAL</td>"
End If
If Check1.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=black>" & Format(tot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=black>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=black>" & Format(tot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=black>" & Format(wtot2, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right ><font color=black>&nbsp;</td>"
fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub
