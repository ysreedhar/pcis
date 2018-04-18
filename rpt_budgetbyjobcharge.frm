VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_budgetbyjobcharge 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   10920
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6255
      Left            =   240
      TabIndex        =   25
      Top             =   2280
      Width           =   9855
      ExtentX         =   17383
      ExtentY         =   11033
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Width           =   10815
      Begin VB.CheckBox chk_xchg 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Calc Xchg"
         Height          =   255
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9960
         Picture         =   "rpt_budgetbyjobcharge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8160
         Picture         =   "rpt_budgetbyjobcharge.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9060
         Picture         =   "rpt_budgetbyjobcharge.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DE-L2"
         Height          =   195
         Left            =   5400
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DE-L3"
         Height          =   195
         Left            =   6960
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "L3"
         Height          =   195
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "L2"
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BDGT"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BCWP"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   720
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1040
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6690
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   4050
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   1215
         End
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
         TabIndex        =   17
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "JobCharge"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   720
         Width           =   930
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5620
         Top             =   120
         Width           =   5175
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
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   5415
      End
   End
End
Attribute VB_Name = "rpt_budgetbyjobcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer

Private Sub cbo_proj_Click()
spp = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
List2.Clear
List1.Clear
Dim lst As String
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & spp(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
List2.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
 
 Check1.Value = 1
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

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Call appcolor
 Else
  Call nocolor
 End If
End Sub

Private Sub Check4_Click()
'Call l2rep
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
If chk_xchg.Value = 1 Then
Call calcbudget
End If


If cbo_proj.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
If Check5.Value = 1 Then
If Check3.Value = 1 Then
  Call appcolor
 Else
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy

 End If
 ElseIf Check4.Value = 1 Then
 Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call l2rep
Unload frmBusy

 ElseIf Check6.Value = 1 Then
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call ldwtesc
Unload frmBusy

 
 ElseIf Check7.Value = 1 Then
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call l2dwtesc
Unload frmBusy

 
 End If
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BC BY JOBCHARGE"
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
            
Me.Width = 11415
Me.Height = 9750
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub List2_Click()
List1.Clear
Option1.Value = False
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim h As Integer
h = 0
For h = 0 To List2.ListCount - 1
If List2.Selected(h) = True Then
ju = Split(List2.List(h), "  -  ", Len(List2.List(h)), vbTextCompare)
            Dim rc As New ADODB.Recordset
            If rc.State Then rc.Close
            rc.Open "select DISTINCT(j.job_code),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and j.job_proj_key = '" & nn(0) & "' and j.jobno='" & ju(0) & "'  order by j.job_code", Cn, 3, 2
            While Not rc.EOF
            List1.AddItem rc(0) & "  -  " & rc(1)
            rc.MoveNext
            Wend
            rc.Close
 End If
 Next h
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
Public Sub nocolor()
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
   'fs.WriteLine "        .table {border-style: solid; border-width:1px; padding:0; table-border-color-light:rgb(102,153,225); table-border-color-dark: rgb(0,0,102);}"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bold;"
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
         
  
 Dim bbamt As Double
 Dim bstot As Double
  
 Dim cnt As Integer
 RPTHEADING fs
 cnt = 0
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
  bbamt = 0
 bstot = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 

 
fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black><b>" & gy(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black><b>" & gy(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
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
 ji = Split(nm(0), "-", Len(nm(0)), vbTextCompare)
 If ji(0) = gy(0) Then
  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                

fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"

fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black>" & nm(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 

Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j  where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
 
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='B' and j.job_code='" & nm(0) & "'  and j.job_desc='" & nm(1) & "'  and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
cnt = cnt + 1 '********************************
                                    If cnt >= 52 Then
                                    fs.WriteLine "</table><P></P>"
                                    RPTHEADING fs
                                    cnt = 0
                                    End If
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                            
                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_resccode & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_costcode & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_spread & "</td>"
                'fs.WriteLine "            <td Nowrap>" & fldata1!bd_tranx & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=center >" & fldata1!bd_uom & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_curr & "</td>"
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
               
            If Check2.Value = 1 Then
                    If fldata1!bd_notes <> "" Then
                    Dim jh111 As String
                    jh111 = Mid(fldata1!bd_notes, 1, 15)
                    fs.WriteLine "            <td ><b> " & jh111 & "</td>"
                    Else
                    Dim cd As New ADODB.Recordset
                    If cd.State Then cd.Close
                    cd.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
                        If Not cd.EOF Then
                        Dim jh As String
                        jh = Mid(cd(0), 1, 15)
                        fs.WriteLine "            <td Nowrap> " & jh & "</td>"
                        End If
                    End If
                    fs.WriteLine "        </tr>"
            Else
            
                If fldata1!bd_notes <> "" Then
                Dim jh11 As String
                jh11 = Mid(fldata1!bd_notes, 1, 15)
                fs.WriteLine "            <td ><b> " & jh11 & "</td>"
                Else
                Dim cd1 As New ADODB.Recordset
                If cd1.State Then cd1.Close
                cd1.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
                    If Not cd1.EOF Then
                    Dim jh1 As String
                    jh1 = Mid(cd1(0), 1, 15)
                    fs.WriteLine "            <td Nowrap> " & jh1 & "</td>"
                    End If
                End If
            fs.WriteLine "        </tr>"
            
            
            End If
           
                
                
fldata1.MoveNext
Wend
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = rscd(0)
End If
                cnt = cnt + 1 '***************************
                    If cnt >= 52 Then
                    fs.WriteLine "</table><P></P>"
                    RPTHEADING fs
                    cnt = 0
                    End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
Dim stt As String
stt = Mid(yre & " - " & assk, 1, 35)
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=9>SubTotal   " & stt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=9>SubTotal   " & stt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend


cnt = cnt + 1 '********************************
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs
cnt = 0
End If
Dim sttt As String
sttt = Mid(List1.List(l), 1, 50)

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10>Total    " & sttt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10>Total    " & sttt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
 tot = tot + dtot
bamt2 = bamt2 + bamt1

End If

End If
Next l

cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                             fs.WriteLine "</table><P></P>"
                            RPTHEADING fs
                            cnt = 0
                            End If

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><b>Total    " & List2.List(w) & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right > &nbsp;</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(bamt2, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right > &nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><b> Total    " & List2.List(w) & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right colspan=1 ><font color=white>&nbsp;</td>"
End If

fs.WriteLine "        </tr>"

bbamt = bbamt + bamt2
bstot = bstot + tot

End If
Next w

cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                             fs.WriteLine "</table><P></P>"
                            RPTHEADING fs
                            cnt = 0
                            End If

fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"

If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bbamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right colspan=1 ><font color=white>&nbsp;</td>"
End If

fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
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
   
  


 fs.WriteLine "    <table border=0 cell spacing=0 cell padding =0 bgcolor=blue width=95%>"
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge</td>"
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
   fs.WriteLine "            <td Nowrap><font color=blue>xRate</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>DwT</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Escl</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>BDGT Amt(RM)</td>"
                            If Check2.Value = 1 Then
                               fs.WriteLine "            <td Nowrap><font color=blue>% WrkCmp</td>"
                               fs.WriteLine "            <td Nowrap><font color=blue>BCWP Amt(RM)</td>"
                            End If
   fs.WriteLine "            <td ><font color=blue>Notes/CostCde Desc</td>"
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
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_jobcharge='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost  where bd_costtype='B' and bd_jobcharge='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
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
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
If Not cd.EOF Then
Dim jh As String
jh = Mid(cd(0), 1, 15)
fs.WriteLine "            <td Nowrap><font color=blue>" & jh & "</td>"
End If
End If
fs.WriteLine "        </tr>"
fldata1.MoveNext
Wend

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=brown>SubTotal for the Job - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(bamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><font color=brown>SubTotal for the Job - " & List1.List(l) & "</td>"
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
fs.WriteLine "            <td  colspan=10><font color=brown>Total for the Job - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right  >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=10><font color=brown>Total for the Job - " & List1.List(l) & "</td>"
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
Public Sub RPTHEADING(fs As Object)
            fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
            ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
                fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td colspan=8><b>" & GetCompanyName & "</td>"
                Else
                fs.WriteLine "            <td colspan=6><b>" & GetCompanyName & "</td>"
                End If
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>JobCharge</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
                    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                    If Check2.Value = 1 Then
                    fs.WriteLine "            <td colspan=8><b>BUDGET BY JOBCHARGE(L3)</td>"
                    Else
                    fs.WriteLine "            <td colspan=6><b>BUDGET BY JOBCHARGE(L3)</td>"
                    End If
                    fs.WriteLine "           <td colspan=2><b>JobNo.</td>"
                                If Option1.Value = True Then
                                fs.WriteLine "           <td colspan=2 align=center>All</td>"
                                Else
                                fs.WriteLine "           <td colspan=2 align=center>SeeEndOfReport</td>"
                                End If
                    fs.WriteLine "           <td><b>PrintDate</td>"
                    fs.WriteLine "           <td align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
                    fs.WriteLine "        </tr>"
                
                                fs.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                                If Check2.Value = 1 Then
                                fs.WriteLine "            <td colspan=14>&nbsp;</td>"
                                Else
                                fs.WriteLine "            <td colspan=12>&nbsp;</td>"
                                End If
                                fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge No.</td>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td colspan=12 ><font color=white>Description</td>"
            Else
            fs.WriteLine "            <td colspan=10 ><font color=white>Description</td>"
            End If
            
            fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
            fs.WriteLine "            <td Nowrap ><font color=white>RescCde</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>CostCde</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>SprdCde</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>TotalQty</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>UOM</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>UnitRate</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>xRate</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>DwT</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>Escl</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>BDGT Amt (RM)</td>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td Nowrap align=right><font color=white>%WC</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>BCWP Amt (RM)</td>"
            fs.WriteLine "            <td align=left width=100><font color=white>Notes/CostCde Desc</td>"
            Else
            fs.WriteLine "            <td align=left width=150><font color=white>Notes/CostCde Desc</td>"
            End If
            fs.WriteLine "        </tr>"

End Sub
Public Sub RPTHEADINGDE(fs As Object)
            fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
            ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
                fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
             
                fs.WriteLine "            <td colspan=6><b>" & GetCompanyName & "</td>"
               
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>JobCharge</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
                    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                   
                    fs.WriteLine "            <td colspan=6><b>BUDGET BY JOBCHARGE(L3)</td>"
                  
                    fs.WriteLine "           <td colspan=2><b>JobNo.</td>"
                                If Option1.Value = True Then
                                fs.WriteLine "           <td colspan=2 align=center>All</td>"
                                Else
                                fs.WriteLine "           <td colspan=2 align=center>SeeEndOfReport</td>"
                                End If
                    fs.WriteLine "           <td><b>PrintDate</td>"
                    fs.WriteLine "           <td align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
                    fs.WriteLine "        </tr>"
                
                                fs.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                                
                                fs.WriteLine "            <td colspan=12>&nbsp;</td>"
                              
                                fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge No.</td>"
           
            fs.WriteLine "            <td colspan=10 ><font color=white>Description</td>"
            
            
            fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
            fs.WriteLine "            <td Nowrap ><font color=white>RescCde</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>CostCde</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>SprdCde</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>TotalQty</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>UOM</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>UnitRate</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>xRate</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>DwT</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>Escl</td>"
            fs.WriteLine "            <td Nowrap align=right><font color=white>DWT Amt (RM)</td>"
           
            fs.WriteLine "            <td Nowrap align=right><font color=white>ESCL Amt (RM)</td>"
           
           
            fs.WriteLine "        </tr>"

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

Public Sub l2rep()
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
   'fs.WriteLine "        .table {border-style: solid; border-width:1px; padding:0; table-border-color-light:rgb(102,153,225); table-border-color-dark: rgb(0,0,102);}"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bold;"
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

Dim cnt As Integer
RPTHEADINGL2 fs
cnt = 0


   Dim bbamt As Double
 Dim bstot As Double
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
  bbamt = 0
 bstot = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 

   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black><b>" & gy(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black><b>" & gy(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
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
 ji = Split(nm(0), "-", Len(nm(0)), vbTextCompare)
 If ji(0) = gy(0) Then

  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
                                

fs.WriteLine "        <tr bgcolor=gray  height=15 class=TableFont>"

fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black>" & nm(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 

Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j  where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
 
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='B' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "'   and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
 
 
                stot = stot + fldata1!bd_extdamt
            If Check2.Value = 1 Then
                bamt = bamt + fldata1!bd_bcwpamt
            End If
               
           
           
                
                
fldata1.MoveNext
Wend
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = Mid(rscd(0), 1, 60)
End If
   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

 
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10>SubTotal   " & yre & " - " & assk & "</td>"
fs.WriteLine "            <td align=right > " & Format(stot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td align=right > " & Format(bamt, "###,###,##0.00") & "</td>"
 
Else
fs.WriteLine "            <td  colspan=10>SubTotal   " & yre & " - " & assk & "</td>"
fs.WriteLine "            <td align=right > " & Format(stot, "###,###,##0.00") & "</td>"
 
End If

fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim hj As String
 hj = Mid(List1.List(l), 1, 75)
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><b>Total</b>    " & hj & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt1, "###,###,##0.00") & "</td>"
Else
fs.WriteLine "            <td  colspan=10><b>Total</b>    " & hj & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
End If

fs.WriteLine "        </tr>"
 tot = tot + dtot
bamt2 = bamt2 + bamt1

End If

End If
Next l


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim hk As String
 hk = Mid(List2.List(w), 1, 75)
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><b>Total    " & hk & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  align=right> <b>" & Format(bamt2, "###,###,##0.00") & "</td>"
 
Else
fs.WriteLine "            <td  colspan=10><b> Total    " & hk & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
 
End If

fs.WriteLine "        </tr>"

bbamt = bbamt + bamt2
bstot = bstot + tot

End If
Next w


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If

fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
 
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bbamt, "###,###,##0.00") & "</td>"
 
Else
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
 
End If

fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
    
    
    fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
fs.WriteLine " </table>"
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"



End Sub

Public Sub RPTHEADINGL2(fs As Object)
      fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
    ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
                fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td colspan=6><b>" & GetCompanyName & "</td>"
                Else
                fs.WriteLine "            <td colspan=5><b>" & GetCompanyName & "</td>"
                End If
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>JobCharge</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
                    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                    If Check2.Value = 1 Then
                    fs.WriteLine "            <td colspan=6><b>BUDGET BY JOBCHARGE(L2)</td>"
                    Else
                    fs.WriteLine "            <td colspan=5><b>BUDGET BY JOBCHARGE(l2)</td>"
                    End If
                    fs.WriteLine "           <td colspan=2><b>JobNo.</td>"
                                If Option1.Value = True Then
                                fs.WriteLine "           <td colspan=2 align=center>All</td>"
                                Else
                                fs.WriteLine "           <td colspan=2 align=center>SeeEndOfReport</td>"
                                End If
                    fs.WriteLine "           <td><b>PrintDate</td>"
                    fs.WriteLine "           <td align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
                    fs.WriteLine "        </tr>"
                
                                fs.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                                If Check2.Value = 1 Then
                                fs.WriteLine "            <td colspan=12>&nbsp;</td>"
                                Else
                                fs.WriteLine "            <td colspan=11>&nbsp;</td>"
                                End If
                                fs.WriteLine "        </tr>"
  
 
            
            
            fs.WriteLine "        <tr bgcolor=black  class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobNo.</td>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td colspan=12 ><font color=white>Description</td>"
            Else
            fs.WriteLine "            <td colspan=10 ><font color=white>Description</td>"
            End If
            
            
            fs.WriteLine "        <tr bgcolor=black  class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge No.</td>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td colspan=8 ><font color=white>Description</td>"
            fs.WriteLine "            <td  ><font color=white>BDGT (RM)</td>"
            fs.WriteLine "            <td  ><font color=white>BCWP (RM)</td>"
            Else
            fs.WriteLine "            <td colspan=8 ><font color=white>Description</td>"
            fs.WriteLine "            <td  ><font color=white>BDGT (RM)</td>"
            End If
            
End Sub


Public Sub ldwtesc()
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
   'fs.WriteLine "        .table {border-style: solid; border-width:1px; padding:0; table-border-color-light:rgb(102,153,225); table-border-color-dark: rgb(0,0,102);}"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bold;"
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
         
  
 Dim bbamt As Double
 Dim bstot As Double
  
 Dim cnt As Integer
 RPTHEADINGDE fs
 cnt = 0
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
  bbamt = 0
 bstot = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 

 
fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black><b>" & gy(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black><b>" & gy(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
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
 ji = Split(nm(0), "-", Len(nm(0)), vbTextCompare)
 If ji(0) = gy(0) Then
  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGDE fs
                                cnt = 0
                                End If
                                

fs.WriteLine "        <tr bgcolor=gray  height=15 class=TableFont>"

fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=13 ><font color=black>" & nm(1) & "</td>"
Else
fs.WriteLine "            <td colspan=11 ><font color=black>" & nm(1) & "</td>"
End If
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 

Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j  where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
 
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='B' and j.job_code='" & nm(0) & "'  and j.job_desc='" & nm(1) & "'  and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
cnt = cnt + 1 '********************************
                                    If cnt >= 52 Then
                                    fs.WriteLine "</table><P></P>"
                                    RPTHEADINGDE fs
                                    cnt = 0
                                    End If
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                            
                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_resccode & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_costcode & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_spread & "</td>"
                'fs.WriteLine "            <td Nowrap>" & fldata1!bd_tranx & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=center >" & fldata1!bd_uom & "</td>"
                fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_curr & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td Nowrap align=right>" & Format((fldata1!bd_downtime / 100) * (fldata1!bd_unitrate), "###,###,##0.00") & "</td>"
                stot = stot + ((fldata1!bd_downtime / 100) * (fldata1!bd_unitrate))
                fs.WriteLine "            <td Nowrap align=right>" & Format((fldata1!bd_escl / 100) * (fldata1!bd_unitrate), "###,###,##0.00") & "</td>"
                bamt = bamt + ((fldata1!bd_escl / 100) * (fldata1!bd_unitrate))
                
               
               fs.WriteLine "        </tr>"
                            
fldata1.MoveNext
Wend
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = rscd(0)
End If
                cnt = cnt + 1 '***************************
                    If cnt >= 52 Then
                    fs.WriteLine "</table><P></P>"
                    RPTHEADINGDE fs
                    cnt = 0
                    End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
Dim stt As String
stt = Mid(yre & " - " & assk, 1, 35)
 
 
fs.WriteLine "            <td  colspan=9>SubTotal   " & stt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend


cnt = cnt + 1 '********************************
If cnt >= 52 Then
fs.WriteLine "</table><P></P>"
RPTHEADINGDE fs
cnt = 0
End If
Dim sttt As String
sttt = Mid(List1.List(l), 1, 50)

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
 
fs.WriteLine "            <td  colspan=10>Total    " & sttt & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt1, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"
 tot = tot + dtot
bamt2 = bamt2 + bamt1

End If

End If
Next l

cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                             fs.WriteLine "</table><P></P>"
                            RPTHEADINGDE fs
                            cnt = 0
                            End If

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

 
fs.WriteLine "            <td  colspan=10><b>Total    " & List2.List(w) & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  align=right> <b>" & Format(bamt2, "###,###,##0.00") & "</td>"
 
fs.WriteLine "        </tr>"

bbamt = bbamt + bamt2
bstot = bstot + tot

End If
Next w

cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                             fs.WriteLine "</table><P></P>"
                            RPTHEADINGDE fs
                            cnt = 0
                            End If

fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"

 
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bbamt, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
fs.WriteLine " </table>"
    
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Public Sub l2dwtesc()
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
   'fs.WriteLine "        .table {border-style: solid; border-width:1px; padding:0; table-border-color-light:rgb(102,153,225); table-border-color-dark: rgb(0,0,102);}"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bold;"
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

Dim cnt As Integer
RPTHEADINGL2DE fs
cnt = 0


   Dim bbamt As Double
 Dim bstot As Double
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
  bbamt = 0
 bstot = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 

   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
 
fs.WriteLine "            <td colspan=13 ><font color=black><b>" & gy(1) & "</td>"
 
'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
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
 ji = Split(nm(0), "-", Len(nm(0)), vbTextCompare)
 If ji(0) = gy(0) Then

  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If
                                

fs.WriteLine "        <tr bgcolor=gray  height=15 class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=13 ><font color=black>" & nm(1) & "</td>"
fs.WriteLine "        </tr>"
 

Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j  where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
dtot = 0
ktot = 0
bamt1 = 0
 
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='B' and j.job_code='" & nm(0) & "' and j.job_desc='" & nm(1) & "'   and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata1.EOF
 
 
                stot = stot + ((fldata1!bd_downtime / 100) * (fldata1!bd_unitrate))
            If Check2.Value = 1 Then
                bamt = bamt + ((fldata1!bd_escl / 100) * (fldata1!bd_unitrate))
            End If
                
                
fldata1.MoveNext
Wend
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = Mid(rscd(0), 1, 60)
End If
   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

 
fs.WriteLine "            <td  colspan=10>SubTotal   " & yre & " - " & assk & "</td>"
fs.WriteLine "            <td align=right > " & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right > " & Format(bamt, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim hj As String
 hj = Mid(List1.List(l), 1, 75)
 
 
fs.WriteLine "            <td  colspan=10><b>Total</b>    " & hj & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td align=right ><b>" & Format(bamt1, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"
 tot = tot + dtot
bamt2 = bamt2 + bamt1

End If

End If
Next l


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim hk As String
 hk = Mid(List2.List(w), 1, 75)
 
 
fs.WriteLine "            <td  colspan=10><b>Total    " & hk & "</td>"
fs.WriteLine "            <td  align=right> <b>" & Format(tot, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  align=right> <b>" & Format(bamt2, "###,###,##0.00") & "</td>"
 

fs.WriteLine "        </tr>"

bbamt = bbamt + bamt2
bstot = bstot + tot

End If
Next w


  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2DE fs
                                cnt = 0
                                End If

fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
 
 
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bstot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td  align=right><font color=white><b>" & Format(bbamt, "###,###,##0.00") & "</td>"
 
fs.WriteLine "        </tr>"
fs.WriteLine " </table>"

fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
fs.WriteLine " </table>"
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Public Sub RPTHEADINGL2DE(fs As Object)
      fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
      ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
                fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                fs.WriteLine "            <td colspan=6><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>JobCharge</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
                    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                    fs.WriteLine "            <td colspan=6><b>BUDGET BY JOBCHARGE(DE-L2)</td>"
                    fs.WriteLine "           <td colspan=2><b>JobNo.</td>"
                                If Option1.Value = True Then
                                fs.WriteLine "           <td colspan=2 align=center>All</td>"
                                Else
                                fs.WriteLine "           <td colspan=2 align=center>SeeEndOfReport</td>"
                                End If
                    fs.WriteLine "           <td><b>PrintDate</td>"
                    fs.WriteLine "           <td align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
                    fs.WriteLine "        </tr>"
                
                                fs.WriteLine "        <tr bgcolor=white  height=8 class=TableFont>"
                                fs.WriteLine "            <td colspan=12>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
  
            fs.WriteLine "        <tr bgcolor=black  class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobNo.</td>"
            fs.WriteLine "            <td colspan=12 ><font color=white>Description</td>"
             
            fs.WriteLine "        <tr bgcolor=black  class=TableFont>"
            fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge No.</td>"
            fs.WriteLine "            <td colspan=8 ><font color=white>Description</td>"
            fs.WriteLine "            <td  ><font color=white>DWT (RM)</td>"
            fs.WriteLine "            <td  ><font color=white>ESCL (RM)</td>"
            

End Sub

Public Sub calcbudget()
 On Error Resume Next
Dim j As Integer
j = 0
For j = 0 To List1.ListCount - 1
If List1.Selected(j) = True Then
xk = Split(List1.List(j), "  -  ", Len(List1.List(j)), vbTextCompare)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where bd_jobcharge='" & xk(0) & "' and bd_costtype='B'  ", Cn, 3, 2
While Not fldata.EOF

fldata!bd_extdamt = (((fldata!bd_tqty) * (fldata!bd_xchg) * (fldata!bd_unitrate)) * ((100 + fldata!bd_downtime)) / 100) * ((100 + fldata!bd_escl) / 100)
fldata!bd_bcwpamt = (fldata!bd_extdamt) * (fldata!bd_wrkcomp / 100)
fldata.Update

fldata.MoveNext
Wend
End If
Next j
End Sub
