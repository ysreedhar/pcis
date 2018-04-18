VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_costdetails 
   BackColor       =   &H00FFFFFF&
   Caption         =   "L3 - PRCR @ DETAILS LEVEL - BY PROJECT KEY / JOB"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6975
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   12303
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "JobNo - Description"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11175
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10080
         Picture         =   "rpt_costdetails.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to Save"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10080
         Picture         =   "rpt_costdetails.frx":057F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to Exit"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_costdetails.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to View"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_costdetails.frx":1199
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to Print"
         Top             =   960
         Width           =   735
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3255
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   120
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.ComboBox cbo_job 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Key - Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "rpt_costdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer
Private Sub cbo_job_Click()
spp = Split(cbo_job.Text, "-", Len(cbo_job.Text), vbTextCompare)
WebBrowser.Navigate "About:Blank"
List1.Clear
Dim lst As String
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & spp(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
List1.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
 hh = 0
         For hgg = 0 To List1.ListCount - 1
         List1.Selected(hgg) = False
         Next hgg
         Option1.Value = 0
         Option2.Value = 0
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
On Error Resume Next
If cbo_job.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_costdetails.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml(False)
Unload frmBusy
End Sub

Private Sub Command1_Click()
On Error Resume Next
If cbo_job.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
filepathl3.Show
SetParent filepathl3.hwnd, rpt_costdetails.hwnd
End Sub

Private Sub command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT:
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "L3 - PRCR @ DETAILS LEVEL - BY PROJECT KEY / JOB"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pk.EOF
cbo_job.AddItem pk(0) & "  -  " & pk(1)
pk.MoveNext
Wend
pk.Close
main.DTPcutdate1.Value = main.DTPcutdate1.Value
Option1.Value = False
Option2.Value = True
' Me.Width = 11415
'Me.Height = 9750
End Sub
Public Sub rephtml(boolSaveAsExcel As Boolean)
On Error Resume Next
nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
 Dim fso As New FileSystemObject
 Dim fs As Object
  If boolSaveAsExcel = True Then
Set fs = fso.CreateTextFile("C:\PCIS-Reports\" & filpat, True)
Else
Set fs = fso.CreateTextFile(App.Path & "\rep.html")
End If
   'Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   
   fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
'   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
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
 Dim cnt As Double

 cnt = 0
Dim atot As Double
Dim ctot As Double
Dim dtot As Double
Dim ftot As Double
Dim gtot As Double
atot = 0: ctot = 0: dtot = 0: ftot = 0: gtot = 0
  Dim Descrp As String
   Dim l As Integer
   l = 0
   For l = 0 To List1.ListCount - 1
   If List1.Selected(l) = True Then
   nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
    RPTHEADING fs, l
    If List1.Selected(l) = True Then ''''''''''''''''''''''''''''''''''''''''''''
 

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(j.job_code),j.job_desc,j.job_proj_key from jobcharge j,cost c where j.job_code=c.bd_jobcharge and j.jobno='" & nm(0) & "' and j.job_proj_key='" & nn(0) & "'  order by j.job_code", Cn, 3, 2
While Not rs.EOF
Dim a As Double
Dim b  As Double
Dim c As Double
Dim d As Double
Dim f As Double
Dim g As Double
a = 0: b = 0: c = 0: d = 0: f = 0: g = 0
cnt = cnt + 1
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs, l
cnt = 0
End If
Descrp = ""
Descrp = nm(1) & "   -    " & rs(1)
fs.WriteLine "        <tr bgcolor=#aeaeae height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=10><font color=black>" & Descrp & "</b></td>"
fs.WriteLine "        </tr>"
Dim rt As New ADODB.Recordset
If rt.State Then rt.Close
rt.Open "select DISTINCT(bd_costcode), bd_chargetype  from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "' and bd_costtype = 'E' GROUP BY  bd_costcode, bd_chargetype   ", Cn, 3, 2
'rt.Open "select DISTINCT(bd_costcode)  from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  ", Cn, 3, 2
While Not rt.EOF
cnt = cnt + 1
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs, l
cnt = 0
End If

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
Dim ct As New ADODB.Recordset
                    If ct.State Then ct.Close
                    ct.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & rt(0) & "' ", Cn, 3, 2
                    If Not ct.EOF Then
                    fs.WriteLine "            <td Nowrap  >" & ct(0) & "</td>"
                    Else
                    fs.WriteLine "            <td Nowrap  >&nbsp;</td>"
                    End If
fs.WriteLine "            <td Nowrap  >" & rs(0) & "</td>"
fs.WriteLine "            <td>" & rt("bd_chargetype") & "</td>"
fs.WriteLine "            <td Nowrap  >" & rt(0) & "</td>"

Dim rt1 As New ADODB.Recordset
If rt1.State Then rt1.Close
rt1.Open "select SUM(bd_extdamt) ,SUM(bd_bcwpamt) from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  and bd_costcode='" & rt(0) & "' and bd_costtype='B' and bd_chargetype = '" & rt("bd_Chargetype") & "'  GROUP BY bd_costcode order by bd_costcode", Cn, 3, 2
If Not rt1.EOF Then




fs.WriteLine "            <td Nowrap  align=right >" & Format(rt1(0), "###,###,##0.00") & "</td>"
a = a + rt1(0) 'bdgt
                If rt1(1) = 0 Then
                fs.WriteLine "            <td Nowrap  align=right>" & Format(rt1(1), "###,###,##0.00") & "</td>"
                b = 0
                b = Format(rt1(1), "###,###,##0.00")
                Else
                fs.WriteLine "            <td Nowrap  align=right>" & Format(Round((100 / (rt1(0) / rt1(1))), 3), "###,###,##0.00") & "</td>"
                b = 0
                b = Round((100 / (rt1(0) / rt1(1))), 3)
                End If
c = c + rt1(1) 'bcwp
fs.WriteLine "            <td Nowrap  align=right>" & Format(rt1(1), "###,###,##0.00") & "</td>"
Else
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"

End If

Dim rt2 As New ADODB.Recordset
If rt2.State Then rt2.Close
rt2.Open "select SUM(bd_extdamt) ,SUM(bd_e_extdamt) from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  and bd_costcode='" & rt(0) & "' and bd_costtype='E' and bd_chargetype = '" & rt("bd_Chargetype") & "' GROUP BY bd_costcode order by bd_costcode ", Cn, 3, 2
If Not rt2.EOF Then
fs.WriteLine "            <td Nowrap  align=right>" & Format(rt2(0), "###,###,##0.00") & "</td>"
d = d + rt2(0) 'acwp
fs.WriteLine "            <td Nowrap  align=right>" & Format(rt2(1), "###,###,##0.00") & "</td>"
f = f + rt2(1) 'ectc
fs.WriteLine "            <td Nowrap  align=right>" & Format(rt2(1) + rt2(0), "###,###,##0.00") & "</td>"
g = g + (rt2(1) + rt2(0)) 'eac
Else
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs.WriteLine "            <td Nowrap  align=right>0.00</td>"
End If
fs.WriteLine "        </tr>"

rt.MoveNext
Wend
cnt = cnt + 1
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs, l
cnt = 0
End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=4><b>Sub Total</B></td>"
'fs.WriteLine "            <td Nowrap >Resp</td>"
fs.WriteLine "            <td align=right><b>" & Format(a, "###,###,##0.00") & "</td>"
'fs.WriteLine "            <td align=right>" & Format(b, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right><b>" & Format(c, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b>" & Format(d, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b>" & Format(f, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b>" & Format(g, "###,###,##0.00") & "</td>"
fs.WriteLine "        </tr>"
atot = atot + a
ctot = ctot + c
dtot = dtot + d
ftot = ftot + f
gtot = gtot + g
rs.MoveNext
Wend
End If
End If
Next l
cnt = cnt + 1
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs, l
cnt = 0
End If
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=4><b><b><font color=white>Total</b></td>"
'fs.WriteLine "            <td Nowrap >Resp</td>"
fs.WriteLine "            <td align=right><b><font color=white>" & Format(atot, "###,###,##0.00") & "</td>"
'fs.WriteLine "            <td align=right>" & Format(b, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right><b><font color=white>" & Format(ctot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b><font color=white>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b><font color=white>" & Format(ftot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><b><font color=white>" & Format(gtot, "###,###,##0.00") & "</td>"
fs.WriteLine "        </tr>"
   
fs.WriteLine " </table>"

   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
   fs.Close
   'WebBrowser.Navigate App.Path & "\rep.html"
   If boolSaveAsExcel = True Then
    WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
Else
    WebBrowser.Navigate App.Path & "\rep.html"
End If

Set fs = Nothing
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
End Sub
Public Sub RPTHEADING(fs As Object, intCurrentIndex As Integer)
On Error Resume Next
  fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
   fs.WriteLine "        <tr bgcolor=white height=25 class=TableFont>"
   fs.WriteLine "            <td Nowrap> " & GetCompanyName & "</td>"
   fs.WriteLine "            <td Nowrap colspan=3> PROJECT COST DETAILS</td>"
   fs.WriteLine "            <td Nowrap align=center colspan=5> (All Amounts Reported in RM)</td>"
   fs.WriteLine "            <td Nowrap> Schedule L3</td>"
   fs.WriteLine "        </tr>"
nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white height=25 class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=8> Project :&nbsp;&nbsp;&nbsp;&nbsp;" & UCase(nn(1)) & "</td>"
fs.WriteLine "            <td Nowrap colspan=2> CuttOffdate :" & main.DTPcutdate1.Value & "</td>"
'fs.WriteLine "            <td Nowrap  >" & Format(Date, "dd/MMM/yyyy") & "</td>"
fs.WriteLine "        </tr>"

  fs.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                            fs.WriteLine "            <td colspan=10>&nbsp;</td>"
                            fs.WriteLine "        </tr>"


fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
fs.WriteLine "            <td colspan=10><font color=white>JobCharge Description - " & List1.List(intCurrentIndex) & " </td>"
fs.WriteLine "        </tr>"

                fs.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
                fs.WriteLine "            <td Nowrap><font color=white>Costcode Desc</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Job-SubJob</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Cg.Type</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Costcode</td>"
                'fs.WriteLine "            <td Nowrap >Resp</td>"
                fs.WriteLine "            <td align=center><font color=white>BDGT</td>"
                fs.WriteLine "            <td align=center><font color=white>%WC</td>"
                fs.WriteLine "            <td align=center><font color=white>BCWP</td>"
                fs.WriteLine "            <td align=center><font color=white>ACWP</td>"
                fs.WriteLine "            <td align=center><font color=white>ECTC</td>"
                fs.WriteLine "            <td align=center><font color=white>EAC</td>"
                fs.WriteLine "        </tr>"
End Sub
Public Sub RPTHEADING1(fs1 As Object)
  fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
 
   fs1.WriteLine "        <tr bgcolor=white height=25 class=TableFont>"
   fs1.WriteLine "            <td Nowrap> " & GetCompanyName & "</td>"
   fs1.WriteLine "            <td Nowrap colspan=2> PROJECT COST DETAILS</td>"
   fs1.WriteLine "            <td Nowrap align=center colspan=5> (All Amounts Reported in RM)</td>"
   fs1.WriteLine "            <td Nowrap> Schedule L3</td>"
   fs1.WriteLine "        </tr>"
    
nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
fs1.WriteLine "        <tr bgcolor=white height=25 class=TableFont>"
fs1.WriteLine "            <td Nowrap colspan=7> Project :&nbsp;&nbsp;&nbsp;&nbsp;" & UCase(nn(1)) & "</td>"
fs1.WriteLine "            <td Nowrap colspan=2> CuttOffdate :" & main.DTPcutdate1.Value & "</td>"
'fs1.WriteLine "            <td Nowrap  >" & Format(Date, "dd/MMM/yyyy") & "</td>"
fs1.WriteLine "        </tr>"

  fs1.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                            fs1.WriteLine "            <td colspan=9>&nbsp;</td>"
                            fs1.WriteLine "        </tr>"


fs1.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
fs1.WriteLine "            <td colspan=9><font color=white>JobCharge Description</td>"
fs1.WriteLine "        </tr>"


                fs1.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
                fs1.WriteLine "            <td Nowrap><font color=white>Costcode Desc</td>"
                fs1.WriteLine "            <td Nowrap ><font color=white>Job-SubJob</td>"
                fs1.WriteLine "            <td Nowrap ><font color=white>Cg. Type</td>"
                fs1.WriteLine "            <td Nowrap ><font color=white>Costcode</td>"
                'fs1.WriteLine "            <td Nowrap >Resp</td>"
                fs1.WriteLine "            <td align=center><font color=white>BDGT</td>"
                fs1.WriteLine "            <td align=center><font color=white>%WC</td>"
                fs1.WriteLine "            <td align=center><font color=white>BCWP</td>"
                fs1.WriteLine "            <td align=center><font color=white>ACWP</td>"
                fs1.WriteLine "            <td align=center><font color=white>ECTC</td>"
                fs1.WriteLine "            <td align=center><font color=white>EAC</td>"
                fs1.WriteLine "        </tr>"
End Sub

Public Sub rephtml1()
Me.Top = 10
Me.Left = 10
nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
 Dim fso As New FileSystemObject
 Dim fs1 As Object
        With fso
        '        strName = .BuildPath(C:\, rep1.html)
        Set fs1 = .CreateTextFile(App.Path & filpat, True)
        
        End With
  fs1.WriteLine " <html> "
  fs1.WriteLine "<style>"
  fs1.WriteLine "    BODY INPUT"
  fs1.WriteLine "    {"
  fs1.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
  fs1.WriteLine "    }"
  fs1.WriteLine "    .TableFont"
  fs1.WriteLine "    {"
  fs1.WriteLine "        COLOR: Black;"
  fs1.WriteLine "        FONT-FAMILY: Arial Narrow;"
  fs1.WriteLine "        FONT-SIZE: 8pt;"
  fs1.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bolder;"
  fs1.WriteLine "        CURSOR:HAND;"
  fs1.WriteLine "    }"
  fs1.WriteLine "    .TrFont"
  fs1.WriteLine "    {"
  fs1.WriteLine "        COLOR: black;"
  fs1.WriteLine "        FONT-FAMILY: Arial Narrow;"
  fs1.WriteLine "        FONT-SIZE: 8pt;"
  fs1.WriteLine "        TEXT-TRANSFORM: capitalize;"
  fs1.WriteLine "        CURSOR:HAND;"
  fs1.WriteLine "   }"
  fs1.WriteLine "</style>"
  fs1.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
  fs1.WriteLine "<body scroll=auto>"
  fs1.WriteLine "    <center>"
    Dim Descr1 As String
 Dim cnt As Integer
 
 RPTHEADING1 fs1
 cnt = 0
   
Dim atot As Double
Dim ctot As Double
Dim dtot As Double
Dim ftot As Double
Dim gtot As Double
atot = 0: ctot = 0: dtot = 0: ftot = 0: gtot = 0

   Dim l As Integer
   l = 0
   For l = 0 To List1.ListCount - 1
   If List1.Selected(l) = True Then
   nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
   If List1.Selected(l) = True Then ''''''''''''''''''''''''''''''''''''''''''''
 

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(j.job_code),j.job_desc,j.job_proj_key from jobcharge j,cost c where j.job_code=c.bd_jobcharge and j.jobno='" & nm(0) & "' and j.job_proj_key='" & nn(0) & "' order by j.job_code", Cn, 3, 2
While Not rs.EOF

Dim a As Double
Dim b  As Double
Dim c As Double
Dim d As Double
Dim f As Double
Dim g As Double
a = 0: b = 0: c = 0: d = 0: f = 0: g = 0
cnt = cnt + 1
If cnt >= 52 Then
fs1.WriteLine "</table><P></P>"
Descr1 = ""
Descr1 = nm(1) & "   -    " & rs(1)
RPTHEADING fs1, l
cnt = 0
End If

fs1.WriteLine "        <tr bgcolor=#aeaeae height=15 class=TableFont>"
fs1.WriteLine "            <td Nowrap colspan=10><font color=black>" & Descr1 & "</b></td>"
fs1.WriteLine "        </tr>"
Dim rt As New ADODB.Recordset
If rt.State Then rt.Close
rt.Open "select DISTINCT(bd_costcode)  from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  GROUP BY  bd_costcode   ", Cn, 3, 2
While Not rt.EOF
cnt = cnt + 1
If cnt >= 52 Then
fs1.WriteLine "</table><P></P>"
RPTHEADING1 fs1
cnt = 0
End If
fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
Dim ct As New ADODB.Recordset
                    If ct.State Then ct.Close
                    ct.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & rt(0) & "' ", Cn, 3, 2
                    If Not ct.EOF Then
                   fs1.WriteLine "            <td Nowrap  >" & ct(0) & "</td>"
                    End If
fs1.WriteLine "            <td Nowrap  >" & rs(0) & "</td>"
fs1.WriteLine "            <td Nowrap  >" & rt(0) & "</td>"
Dim rt1 As New ADODB.Recordset
If rt1.State Then rt1.Close
rt1.Open "select SUM(bd_extdamt) ,SUM(bd_bcwpamt) from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  and bd_costcode='" & rt(0) & "' and bd_costtype='B' GROUP BY bd_costcode order by bd_costcode", Cn, 3, 2
If Not rt1.EOF Then



'fs1.WriteLine "            <td Nowrap   >" & rt(1) & "</td>"
fs1.WriteLine "            <td Nowrap  align=right >" & Format(rt1(0), "###,###,##0.00") & "</td>"
a = a + rt1(0) 'bdgt
                If rt1(1) = 0 Then
               fs1.WriteLine "            <td Nowrap  align=right>" & Format(rt1(1), "###,###,##0.00") & "</td>"
                b = 0
                b = Format(rt1(1), "###,###,##0.00")
                Else
               fs1.WriteLine "            <td Nowrap  align=right>" & Format(Round((100 / (rt1(0) / rt1(1))), 3), "###,###,##0.00") & "</td>"
                b = 0
                b = Round((100 / (rt1(0) / rt1(1))), 3)
                End If
c = c + rt1(1) 'bcwp
fs1.WriteLine "            <td Nowrap  align=right>" & Format(rt1(1), "###,###,##0.00") & "</td>"
Else
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"

End If

Dim rt2 As New ADODB.Recordset
If rt2.State Then rt2.Close
rt2.Open "select SUM(bd_extdamt) ,SUM(bd_e_extdamt) from cost where bd_jobcharge='" & rs(0) & "' and bd_projectkey='" & rs(2) & "'  and bd_costcode='" & rt(0) & "' and bd_costtype='E' GROUP BY bd_costcode order by bd_costcode ", Cn, 3, 2
If Not rt2.EOF Then
fs1.WriteLine "            <td Nowrap  align=right>" & Format(rt2(0), "###,###,##0.00") & "</td>"
d = d + rt2(0) 'acwp
fs1.WriteLine "            <td Nowrap  align=right>" & Format(rt2(1), "###,###,##0.00") & "</td>"
f = f + rt2(1) 'ectc
fs1.WriteLine "            <td Nowrap  align=right>" & Format(rt2(1) + rt2(0), "###,###,##0.00") & "</td>"
g = g + (rt2(1) + rt2(0)) 'eac
Else
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"
fs1.WriteLine "            <td Nowrap  align=right>0.00</td>"
End If
fs1.WriteLine "        </tr>"

rt.MoveNext
Wend
cnt = cnt + 1
If cnt >= 52 Then
fs1.WriteLine "</table><P></P>"
RPTHEADING1 fs1
cnt = 0
End If
fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs1.WriteLine "            <td Nowrap colspan=3><b>Sub Total</B></td>"
 
'fs1.WriteLine "            <td Nowrap >Resp</td>"
fs1.WriteLine "            <td align=right><b>" & Format(a, "###,###,##0.00") & "</td>"
'fs1.WriteLine "            <td align=right>" & Format(b, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right>&nbsp;</td>"
fs1.WriteLine "            <td align=right><b>" & Format(c, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b>" & Format(d, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b>" & Format(f, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b>" & Format(g, "###,###,##0.00") & "</td>"
fs1.WriteLine "        </tr>"
atot = atot + a
ctot = ctot + c
dtot = dtot + d
ftot = ftot + f
gtot = gtot + g
rs.MoveNext
Wend


End If
End If
Next l
cnt = cnt + 1
If cnt >= 52 Then
fs1.WriteLine "</table><P></P>"
RPTHEADING1 fs1
cnt = 0
End If
fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs1.WriteLine "            <td Nowrap colspan=3><b><b><font color=white>Total</b></td>"

'fs1.WriteLine "            <td Nowrap >Resp</td>"
fs1.WriteLine "            <td align=right><b><font color=white>" & Format(atot, "###,###,##0.00") & "</td>"
'fs1.WriteLine "            <td align=right>" & Format(b, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right>&nbsp;</td>"
fs1.WriteLine "            <td align=right><b><font color=white>" & Format(ctot, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b><font color=white>" & Format(dtot, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b><font color=white>" & Format(ftot, "###,###,##0.00") & "</td>"
fs1.WriteLine "            <td align=right><b><font color=white>" & Format(gtot, "###,###,##0.00") & "</td>"
fs1.WriteLine "        </tr>"
   
   
   
  fs1.WriteLine " </table>"
    
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
  fs1.WriteLine "    </table><br>"
  fs1.WriteLine "    </body>"
  fs1.WriteLine "    <html>"

End Sub
