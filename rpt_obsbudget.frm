VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_obsbudget 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   15045
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   15045
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6495
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   11655
      ExtentX         =   20558
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
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10440
         Picture         =   "rpt_obsbudget.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Click to Exit"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8760
         Picture         =   "rpt_obsbudget.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Click to View"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9600
         Picture         =   "rpt_obsbudget.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Click to Print"
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   480
         Width           =   4815
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Top             =   120
         Width           =   4815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   735
         Left            =   120
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   10
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
      Top             =   1560
      Width           =   11175
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BCWP"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BDGT"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "rpt_obsbudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer

Private Sub cbo_proj_Click()
List1.Clear
Option1.Value = False
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim rc As New ADODB.Recordset
 If rc.State Then rc.Close
 rc.Open "select DISTINCT(c.bd_obs),r.resp_desc from cost c,responsibledetails r where c.bd_obs=r.resp_code and c.bd_projectkey='" & nn(0) & "'", Cn, 3, 2
 While Not rc.EOF
 List1.AddItem rc(0) & "  -  " & rc(1)
 rc.MoveNext
 Wend
 rc.Close
 'Option1.Value = True
 Check1.Value = 1
         hh = 0
 
         For hgg = 0 To List1.ListCount - 1
         List1.Selected(hgg) = False
         Next hgg
         Option1.Value = 0
         Option2.Value = 0
 
End Sub

Private Sub Check1_Click()
If Check3.Value = 1 Then
  Call appcolor
 Else
  Call nocolor
 End If
End Sub

Private Sub Check2_Click()
If Check3.Value = 1 Then
  Call appcolor
 Else
  Call nocolor
 End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
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
If Check3.Value = 1 Then
  Call appcolor
 Else
  Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy

 End If
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BC BY OBS"
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

Private Sub List1_Click()
'If Check3.Value = 1 Then
'Call appcolor
'Else
'Call nocolor
'End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List1.ListCount - 1
List1.Selected(f) = True
Next f
'If Check3.Value = 1 Then
'Call appcolor
'Else
'Call nocolor
'End If
End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List1.ListCount - 1
List1.Selected(g) = False
Next g
'If Check3.Value = 1 Then
'Call appcolor
'Else
'Call nocolor
'End If
End If
List1.Enabled = True
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
   fs.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
   
   
   fs.WriteLine "            <font size=2.5>" & GetCompanyName & "</font><br>"
    fs.WriteLine "           <font size=2.5>BUDGET BY JOBCHARGE</font><BR>"
   
  


 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
fs.WriteLine "        <tr bgcolor=#acacac  class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=black>OBS Code</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black>Description</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black>Description</td>"
End If
 
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"

   fs.WriteLine "        <tr bgcolor=#bcbcbc height=15 class=TableFont>"
    
   fs.WriteLine "            <td Nowrap><font color=black>Resc Cde</td>"
   fs.WriteLine "            <td Nowrap><font color=black>CostCode</td>"
   fs.WriteLine "            <td Nowrap><font color=black>SprdCde</td>"
   fs.WriteLine "            <td Nowrap><font color=white>TrnxType</td>"
   fs.WriteLine "            <td Nowrap><font color=black>Total Qty</td>"
   fs.WriteLine "            <td Nowrap><font color=black>UOM</td>"
   fs.WriteLine "            <td Nowrap><font color=black>Curcy</td>"
   fs.WriteLine "            <td Nowrap><font color=black>UnitRate</td>"
   fs.WriteLine "            <td Nowrap><font color=black>Xrate</td>"
   fs.WriteLine "            <td Nowrap><font color=black>DT</td>"
   fs.WriteLine "            <td Nowrap><font color=black>Escl</td>"
   fs.WriteLine "            <td Nowrap><font color=black>BDGT Amt(RM)</td>"
   If Check2.Value = 1 Then
      fs.WriteLine "            <td Nowrap><font color=black>% WrkCmp</td>"
      fs.WriteLine "            <td Nowrap><font color=black>BCWP Amt(RM)</td>"
   End If
   fs.WriteLine "            <td ><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
    
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

Dim stot As Double
Dim tot As Double
Dim dtot As Double
stot = 0: tot = 0: dtot = 0
Dim bamt As Double
Dim bamt1 As Double
Dim bamt2 As Double
bamt = 0: bamt1 = 0: bamt2 = 0
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
 
fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black>" & nm(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_obs='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
 
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost  where bd_costtype='B' and bd_obs='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 
fs.WriteLine "            <td Nowrap>" & fldata1!bd_resccode & "</td>"
fs.WriteLine "            <td Nowrap>" & fldata1!bd_costcode & "</td>"
fs.WriteLine "            <td Nowrap>" & fldata1!bd_spread & "</td>"
fs.WriteLine "            <td Nowrap>" & fldata1!bd_tranx & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap  >" & fldata1!bd_uom & "</td>"
fs.WriteLine "            <td Nowrap >" & fldata1!bd_curr & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</td>"
stot = stot + fldata1!bd_extdamt
If Check2.Value = 1 Then
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_wrkcomp, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_bcwpamt, "###,###,##0.00") & "</td>"
bamt = bamt + fldata1!bd_bcwpamt

End If
If fldata1!bd_notes <> "" Then
fs.WriteLine "            <td >" & fldata1!bd_notes & "</td>"
Else
fs.WriteLine "            <td Nowrap>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
fldata1.MoveNext
Wend

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10>SubTotal  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right >" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td align=right >" & Format(bamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10>SubTotal  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right >" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend



fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10>Total  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right >" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
fs.WriteLine "            <td align=right >" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10>Total  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right >" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
 tot = tot + dtot
bamt2 = bamt2 + bamt1
End If
Next l
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=11><font color=black>NET TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=black>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=black>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=black>" & Format(bamt2, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=black>&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=11><font color=black>NET TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=black>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right colspan=1 ><font color=black>&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"


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
   
   
 fs.WriteLine "            <font size=2.5 COLOR=BLUE>" & GetCompanyName & "</font><br>"
    fs.WriteLine "           <font size=2.5 COLOR=BLUE>BUDGET BY JOBCHARGE</font><BR>"
   
  


 fs.WriteLine "    <table border=1 cellspacing=1 bgcolor=blue width=95%>"
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=white>OBS Code</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=white>Description</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=white>Description</td>"
End If
 
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"

   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
    
   fs.WriteLine "            <td Nowrap><font color=blue>Resc Cde</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>CostCode</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>SprdCde</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>TrnxType</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Total Qty</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>UOM</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Curcy</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>UnitRate</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Xrate</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>DT</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Escl</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>BDGT Amt(RM)</td>"
   If Check2.Value = 1 Then
      fs.WriteLine "            <td Nowrap><font color=blue>% WrkCmp</td>"
      fs.WriteLine "            <td Nowrap><font color=blue>BCWP Amt(RM)</td>"
   End If
   fs.WriteLine "            <td ><font color=blue>Notes</td>"
   fs.WriteLine "        </tr>"
    
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

Dim stot As Double
Dim tot As Double
Dim dtot As Double
stot = 0: tot = 0: dtot = 0
Dim bamt As Double
Dim bamt1 As Double
Dim bamt2 As Double
bamt = 0: bamt1 = 0: bamt2 = 0
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
 
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=white>" & nm(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=white>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=white>" & nm(1) & "</td>"
End If
 
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_obs='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost  where bd_costtype='B' and bd_obs='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
fs.WriteLine "        <tr bgcolor=white class=TableFont>"
 
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_resccode & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_costcode & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_spread & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_tranx & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap  ><font color=blue>" & fldata1!bd_uom & "</td>"
fs.WriteLine "            <td Nowrap ><font color=blue>" & fldata1!bd_curr & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</td>"
stot = stot + fldata1!bd_extdamt
If Check2.Value = 1 Then
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_wrkcomp, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_bcwpamt, "###,###,##0.00") & "</td>"
bamt = bamt + fldata1!bd_bcwpamt

End If
If fldata1!bd_notes <> "" Then
fs.WriteLine "            <td ><font color=blue>" & fldata1!bd_notes & "</td>"
Else
fs.WriteLine "            <td Nowrap>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
fldata1.MoveNext
Wend

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=brown>SubTotal   - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(bamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><font color=brown>SubTotal   - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
End If
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend



fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=brown>Total   - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><font color=brown>Total   - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
 tot = tot + dtot
 bamt2 = bamt2 + bamt1
End If
Next l
fs.WriteLine "        <tr bgcolor=yellow height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=11>NET TOTAL</td>"
fs.WriteLine "            <td  align=right>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td  align=right>" & Format(bamt2, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=11>NET TOTAL</td>"
fs.WriteLine "            <td  align=right>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"



End Sub


