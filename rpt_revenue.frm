VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_revenue 
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
      Height          =   6615
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   11668
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
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00DC7E5A&
      Height          =   480
      Left            =   10200
      Picture         =   "rpt_revenue.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click to Exit"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00DC7E5A&
      Height          =   480
      Left            =   8520
      Picture         =   "rpt_revenue.frx":05FF
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click to View"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00DC7E5A&
      Height          =   480
      Left            =   9360
      Picture         =   "rpt_revenue.frx":0C1A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click to Print"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1205
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   6690
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   480
         Width           =   4050
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   960
         Width           =   1045
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BUDGETED REVENUE / VARIATION ORDER"
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
         Left            =   6720
         TabIndex        =   13
         Top             =   240
         Width           =   4110
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Project key"
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
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5640
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "JobNo."
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
         Left            =   600
         TabIndex        =   10
         Top             =   720
         Width           =   585
      End
   End
End
Attribute VB_Name = "rpt_revenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer
Private Sub cbo_proj_Click()
spp = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
List2.Clear
 
Dim lst As String
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & spp(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
List2.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
  hgg = 0
            For hgg = 0 To List2.ListCount - 1
            List2.Selected(hgg) = False
            Next hgg
            hgg = 0
            For hgg = 0 To List1.ListCount - 1
            List1.Selected(hgg) = False
            Next hgg
            Option1.Value = 0
            Option2.Value = 0
            Option3.Value = 0
            Option4.Value = 0
 

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
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy
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
   nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
        fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
        fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
        fs.WriteLine "            <td colspan=2><b>" & GetCompanyName & "</td>"
        fs.WriteLine "           <td  >Project key</td>"
        fs.WriteLine "           <td  >" & nm(0) & "</td>"
        fs.WriteLine "           <td  >JobNo</td>"
        fs.WriteLine "           <td  >SeeEndOfReport</td>"
        fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><b>BUDGETED REVENUE & VARIATION ORDER</td>"
                fs.WriteLine "           <td  >Cutt-off Date</td>"
                fs.WriteLine "           <td  >" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td colspan=6><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
            fs.WriteLine "            <td colspan=6><font color=white>Revenue Type</td>"
            fs.WriteLine "        </tr>"
            
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap  ><font color=white>Job No & Desc</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>xRate</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount(RM)</td>"
   fs.WriteLine "            <td width=200 align=center><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"

Dim bd As Double
Dim co As Double
Dim ad As Double
Dim bl As Double
Dim ubl As Double
bd = 0: co = 0: ad = 0: bl = 0: ubl = 0
Dim flj1 As New ADODB.Recordset
Dim flj2 As New ADODB.Recordset
Dim flj3 As New ADODB.Recordset

    If List1.Selected(0) = True Then
      If List1.List(0) = "Budgeted Revenue" Then
                    fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=6><font color=black>BDGT(BUDGETED REVENUE)</td>"
                    fs.WriteLine "        </tr>"
                    Dim pnh As String
                    pn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
                    pnh = Mid(pn(0), 1, 3)
  Dim i As Integer
  i = 0
  For i = 0 To List2.ListCount - 1
  If List2.Selected(i) = True Then
  ii = Split(List2.List(i), "  -  ", Len(List2.List(i)), vbTextCompare)
                    Dim fldata As New ADODB.Recordset
                    If fldata.State Then fldata.Close
                    fldata.Open "select * from revenue where rev_type='BGT' and rev_projcode='" & pn(0) & "' and rev_jobno='" & ii(0) & "' order by rev_jobno ", Cn, 3, 2
                 
                    While Not fldata.EOF
                     fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                     
                      If flj1.State Then flj1.Close
                      flj1.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldata!rev_jobno & "'", Cn, 3, 2
                      If Not flj1.EOF Then
                      fs.WriteLine "            <td  >" & fldata!rev_jobno & "  -  " & flj1(0) & "</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldata!rev_jobno & "</td>"
                      End If
                      fs.WriteLine "            <td  align=center>" & fldata!rev_Currency & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_amount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_exchange, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_totamount, "###,###,##0.00") & "</td>"
                      bd = bd + fldata!rev_totamount
                      If fldata!rev_tranxnotes = "" Then
                      fs.WriteLine "            <td  >&nbsp;</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldata!rev_tranxnotes & "</td>"
                      End If
                      fs.WriteLine "        </tr>"
                    fldata.MoveNext
                    Wend
  End If
  Next i
                    fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><font color=white>TOTAL - BUDGETED REVENUE</td>"
                    fs.WriteLine "            <td  align=right><font color=white>" & Format(bd, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
  End If
  End If
  
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td colspan=6><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
  pn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
If List1.Selected(1) = True Then
If List1.List(1) = "Variation Orders - Positive" Then
                    fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=6><font color=black>VO(+) VARIATION ORDERS POSITIVE </td>"
                    fs.WriteLine "        </tr>"
  Dim j As Integer
  j = 0
  For j = 0 To List2.ListCount - 1
  If List2.Selected(j) = True Then
  jj = Split(List2.List(j), "  -  ", Len(List2.List(j)), vbTextCompare)
Dim fldataa As New ADODB.Recordset
If fldataa.State Then fldataa.Close
fldataa.Open "select * from revenue where rev_type='VO(+)' and rev_projcode='" & pn(0) & "' and rev_jobno='" & jj(0) & "' order by rev_jobno ", Cn, 3, 2

 
    While Not fldataa.EOF
                      fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                      If flj2.State Then flj2.Close
                      flj2.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldataa!rev_jobno & "'", Cn, 3, 2
                      If Not flj2.EOF Then
                      fs.WriteLine "            <td  >" & fldataa!rev_jobno & "  -  " & flj2(0) & "</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldataa!rev_jobno & "</td>"
                      End If
                      fs.WriteLine "            <td  align=center>" & fldataa!rev_Currency & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldataa!rev_amount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldataa!rev_exchange, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldataa!rev_totamount, "###,###,##0.00") & "</td>"
                      co = co + fldataa!rev_totamount
                      If fldataa!rev_tranxnotes = "" Then
                      fs.WriteLine "            <td  >&nbsp;</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldataa!rev_tranxnotes & "</td>"
                      End If
                      fs.WriteLine "        </tr>"
      fldataa.MoveNext
    Wend
 End If
 Next j
                    fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><font color=white>TOTAL - VARIATION ORDER(+) POSITIVE</td>"
                    fs.WriteLine "            <td  align=right><font color=white>" & Format(co, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
 
End If
End If

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td colspan=6><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
       If List1.Selected(2) = True Then
       If List1.List(2) = "Variation Orders - Negative" Then
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=6><font color=black>VO(-) VARIATION ORDERS NEGATIVE</font> </td>"
fs.WriteLine "        </tr>"
Dim k As Integer
  k = 0
    pn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
  For k = 0 To List2.ListCount - 1
  If List2.Selected(k) = True Then
  kk = Split(List2.List(k), "  -  ", Len(List2.List(k)), vbTextCompare)
  
       Dim fldatab As New ADODB.Recordset
If fldatab.State Then fldatab.Close
fldatab.Open "select * from revenue where rev_type='VO(-)' and rev_projcode='" & pn(0) & "' and rev_jobno='" & kk(0) & "' order by rev_jobno ", Cn, 3, 2
 
    While Not fldatab.EOF
                      fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                      If flj3.State Then flj3.Close
                      flj3.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & fldatab!rev_jobno & "'", Cn, 3, 2
                      If Not flj3.EOF Then
                      fs.WriteLine "            <td  >" & fldatab!rev_jobno & "  -  " & flj3(0) & "</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldatab!rev_jobno & "</td>"
                      End If
                      fs.WriteLine "            <td  align=center>" & fldatab!rev_Currency & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldatab!rev_amount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldatab!rev_exchange, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldatab!rev_totamount, "###,###,##0.00") & "</td>"
                      ad = ad + fldatab!rev_totamount
                      If fldatab!rev_tranxnotes = "" Then
                      fs.WriteLine "            <td  >&nbsp;</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldatab!rev_tranxnotes & "</td>"
                      End If
                      fs.WriteLine "        </tr>"
      fldatab.MoveNext
    Wend
  End If
  Next k
                    fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><font color=white>TOTAL - VARIATION ORDER(-) NEGATIVE</td>"
                    fs.WriteLine "            <td  align=right><font color=white>" & Format(ad, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
 
       
       End If
       End If
       
       
  
    
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BUDGETED REVENUE/VARIATION ORDER"
Me.Top = 10
Me.Left = 10
Me.Height = 9720
Me.Width = 11220
WebBrowser.Navigate "About:Blank"
 Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pk.EOF
cbo_proj.AddItem pk(0) & "  -  " & pk(1)
pk.MoveNext
Wend
pk.Close

List1.AddItem "Budgeted Revenue"
List1.AddItem "Variation Orders - Positive"
List1.AddItem "Variation Orders - Negative"
            hgg = 0
            For hgg = 0 To List2.ListCount - 1
            List2.Selected(hgg) = False
            Next hgg
            hgg = 0
            For hgg = 0 To List1.ListCount - 1
            List1.Selected(hgg) = False
            Next hgg
            Option1.Value = False
            Option2.Value = True
            Option3.Value = True
            Option4.Value = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub Option1_Click()
Option3.Value = 0
Option4.Value = 0
hgg = 0
            For hgg = 0 To List1.ListCount - 1
            List1.Selected(hgg) = False
            Next hgg
If Option1.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List2.ListCount - 1
List2.Selected(f) = True
Next f
 
End If
 
End Sub

Private Sub Option2_Click()
Option3.Value = 0
Option4.Value = 0
hgg = 0
            For hgg = 0 To List1.ListCount - 1
            List1.Selected(hgg) = False
            Next hgg
If Option2.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List2.ListCount - 1
List2.Selected(g) = False
Next g
 
End If
 
End Sub
Private Sub Option3_Click()
If Option3.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List1.ListCount - 1
List1.Selected(g) = False
Next g
 
End If
 
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List1.ListCount - 1
List1.Selected(f) = True
Next f

End If
 
End Sub

