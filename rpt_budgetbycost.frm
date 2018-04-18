VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_budgetbycost 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   11055
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   11055
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   8916
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
      TabIndex        =   13
      Top             =   1560
      Width           =   11055
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10200
         Picture         =   "rpt_budgetbycost.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8400
         Picture         =   "rpt_budgetbycost.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9300
         Picture         =   "rpt_budgetbycost.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BDGT"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BCWP"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   720
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5640
         TabIndex        =   6
         Top             =   960
         Width           =   1045
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6690
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   4050
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1205
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   2
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
         TabIndex        =   12
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Resource"
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
         TabIndex        =   11
         Top             =   720
         Width           =   825
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
         TabIndex        =   10
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
Attribute VB_Name = "rpt_budgetbycost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim jh As String
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
      
            
            Dim rc As New ADODB.Recordset
            If rc.State Then rc.Close
            rc.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j where c.bd_jobcharge=j.job_code  and  bd_costtype='B' and c.bd_projectkey='" & spp(0) & "' order by c.bd_resccode", Cn, 3, 2
            While Not rc.EOF
            Dim rcd As New ADODB.Recordset
            If rcd.State Then rcd.Close
            rcd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rc(0) & "' ", Cn, 3, 2
                   If Not rcd.EOF Then
                   List1.AddItem rc(0) & "  -  " & rcd(0)
                   Else
                   List1.AddItem rc(0)
                   End If
            
            rc.MoveNext
            Wend
            rc.Close

  
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
main.lbltitle.Caption = "BC BY RESOURCE/COSTCODE"
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
List2.Enabled = False
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
List2.Enabled = True
End Sub

Public Sub nocolor()
Dim fso As New FileSystemObject
   Dim fs As Object
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
   fs.WriteLine "    }"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
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
        RPTHEADING fs
        cnt = 0

Dim stot As Double
Dim tot As Double
Dim dtot As Double
stot = 0: tot = 0: dtot = 0
Dim bamt As Double
Dim bamt1 As Double
Dim bamt2 As Double
bamt = 0: bamt1 = 0: bamt2 = 0
hk = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rg As New ADODB.Recordset
If rg.State Then rg.Close
rg.Open "select * from resourcemaster r, resourcedetails d where r.resc_code=d.dresc_code and  r.resc_code='" & nm(0) & "' and d.dresc_proj='" & hk(0) & "'", Cn, 3, 2
If Not rg.EOF Then
            cnt = cnt + 1 '********************************
            If cnt >= 53 Then
            fs.WriteLine "</table><P></P>"
            RPTHEADING fs
            cnt = 0
            End If
fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=4 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=7 ><font color=black>" & rg!resc_vendorcode & "</td>"
Else
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
End If
fs.WriteLine "        </tr>"
End If

Dim yre As String
Dim yree As String

dtot = 0
bamt1 = 0
 
Dim Y As Integer
Y = 0
For Y = 0 To List2.ListCount - 1
If List2.Selected(Y) = True Then
fl = Split(List2.List(Y), "  -  ", Len(List2.List(Y)), vbTextCompare)
yre = fl(0)
yree = fl(1)
Dim fldata12 As New ADODB.Recordset
If fldata12.State Then fldata12.Close
fldata12.Open "select DISTINCT(c.bd_costcode) from cost c,jobcharge j where c.bd_jobcharge=j.job_code and c.bd_costtype='B' and c.bd_resccode='" & nm(0) & "' and c.bd_projectkey='" & hk(0) & "'  and j.jobno='" & yre & "' order by c.bd_costcode", Cn, 3, 2
stot = 0
bamt = 0
While Not fldata12.EOF


Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and c.bd_costtype='B' and  c.bd_costcode='" & fldata12!bd_costcode & "' and c.bd_resccode='" & nm(0) & "' and c.bd_projectkey='" & hk(0) & "'  and j.jobno='" & yre & "' order by c.bd_costcode,j.jobno ,c.bd_jobcharge", Cn, 3, 2
stota = 0
bamta = 0
Dim stit As Integer
stit = 0
While Not fldata1.EOF
stit = 1
                            
                            cnt = cnt + 1 '********************************
                            If cnt >= 53 Then
                            fs.WriteLine "</table><P></P>"
                            RPTHEADING fs
                            cnt = 0
                            End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap >" & fldata1!bd_costcode & "</td>"
fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_JobCharge & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_uom & "</td>"
fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_curr & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</td>"
stota = stota + fldata1!bd_extdamt
If Check2.Value = 1 Then
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_wrkcomp, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_bcwpamt, "###,###,##0.00") & "</td>"
bamta = bamta + fldata1!bd_bcwpamt

End If
If Check2.Value = 1 Then
                    If fldata1!bd_notes <> "" Then
                    fs.WriteLine "            <td ><b> " & fldata1!bd_notes & "</td>"
                    Else
                    Dim cd As New ADODB.Recordset
                    If cd.State Then cd.Close
                    cd.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
                    If Not cd.EOF Then
                    
                    jh = Mid(cd(0), 1, 20)
                    fs.WriteLine "            <td Nowrap> " & jh & "</td>"
                    End If
                    End If
                    fs.WriteLine "        </tr>"
            Else
            
            If fldata1!bd_notes <> "" Then
            fs.WriteLine "            <td ><b> " & fldata1!bd_notes & "</td>"
            Else
            Dim cd1 As New ADODB.Recordset
            If cd1.State Then cd1.Close
            cd1.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
            If Not cd1.EOF Then
            Dim jh1 As String
            jh1 = Mid(cd1(0), 1, 28)
            fs.WriteLine "            <td Nowrap> " & jh1 & "</td>"
            End If
            End If
            fs.WriteLine "        </tr>"
            
            
            End If
           

fldata1.MoveNext
Wend


If stit <> 0 Then
                    cnt = cnt + 1 '********************************
                    If cnt >= 53 Then
                    fs.WriteLine "</table><P></P>"
                    RPTHEADING fs
                    cnt = 0
                    End If
            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
            fs.WriteLine "            <td  colspan=8>Sub Total   " & fldata12(0) & "  -  " & jh1 & "</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(stota, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(bamta, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            
            Else
            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
            fs.WriteLine "            <td  colspan=8>Sub Total   " & fldata12(0) & "  -  " & jh1 & "</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(stota, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            End If
            fs.WriteLine "        </tr>"
            stot = stot + stota
            bamt = bamt + bamta
            End If



fldata12.MoveNext
Wend
            If stit <> 0 Then
                    cnt = cnt + 1 '********************************
                    If cnt >= 53 Then
                    fs.WriteLine "</table><P></P>"
                    RPTHEADING fs
                    cnt = 0
                    End If
            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            If Check2.Value = 1 Then
            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
            fs.WriteLine "            <td  colspan=8>Sub Total   " & yre & "  -  " & yree & "</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(bamt, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            
            Else
            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
            fs.WriteLine "            <td  colspan=8>Sub Total   " & yre & "  -  " & yree & "</td>"
            fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
            fs.WriteLine "            <td align=right >&nbsp;</td>"
            End If
            fs.WriteLine "        </tr>"
            dtot = dtot + stot
            bamt1 = bamt1 + bamt
            End If
 
End If
Next Y

                cnt = cnt + 1 '********************************
                If cnt >= 53 Then
                fs.WriteLine "</table><P></P>"
                RPTHEADING fs
                cnt = 0
                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
fs.WriteLine "            <td  colspan=8>Total  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "            <td align=right ><b>" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
fs.WriteLine "            <td  colspan=8>Total  - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right >&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
tot = tot + dtot
bamt2 = bamt2 + bamt1

End If
Next l
                    cnt = cnt + 1 '********************************
                    If cnt >= 53 Then
                    fs.WriteLine "</table><P></P>"
                    RPTHEADING fs
                    cnt = 0
                    End If
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=9><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(bamt2, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=9><font color=white>REPORT TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
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
    fs.WriteLine "           <font size=2.5 COLOR=BLUE>BUDGET BY RESOURCE</font><BR>"
   
  
fs.WriteLine "    <table border=1 cellspacing=1 bgcolor=blue width=95%>"
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=7 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=7 ><font color=white>Vendor Name</td>"
Else
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
End If
fs.WriteLine "        </tr>"

   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=blue>Year</td>"
   fs.WriteLine "            <td Nowrap><font color=blue>JobCharge</td>"
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
hk = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rg As New ADODB.Recordset
If rg.State Then rg.Close
rg.Open "select * from resourcemaster r, resourcedetails d where r.resc_code=d.dresc_code and  r.resc_code='" & nm(0) & "' and d.dresc_proj='" & hk(0) & "'", Cn, 3, 2
If Not rg.EOF Then
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td ><font color=white>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=7 ><font color=white>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=white>" & kj(0) & "</td>"
If Check2.Value = 1 Then
fs.WriteLine "            <td colspan=7 ><font color=white>" & rg!resc_vendorcode & "</td>"
Else
fs.WriteLine "            <td colspan=5 ><font color=white>" & rg!resc_vendorcode & "</td>"
End If
fs.WriteLine "        </tr>"
End If

Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_year) from cost  where bd_costtype='B' ", Cn, 3, 2
dtot = 0
bamt = 0
While Not fl.EOF
yre = fl(0)

Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
fldata1.Open "select * from cost  where bd_costtype='B' and bd_resccode='" & nm(0) & "' and bd_year='" & yre & "' ", Cn, 3, 2
stot = 0
bamt1 = 0
While Not fldata1.EOF
fs.WriteLine "        <tr bgcolor=white class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_year & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_costcode & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_JobCharge & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_spread & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_tranx & "</td>"
fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_uom & "</td>"
fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_curr & "</td>"
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
fs.WriteLine "            <td Nowrap><font color=blue>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
fldata1.MoveNext
Wend

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=1><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td  colspan=11><font color=brown>Sub-Total For the year - " & yre & "</td>"
fs.WriteLine "            <td align=right><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td align=right><font color=brown>" & Format(bamt, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><font color=brown>&nbsp;</td>"

Else
fs.WriteLine "            <td  colspan=1><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td  colspan=11><font color=brown>Sub-Total For the year - " & yre & "</td>"
fs.WriteLine "            <td align=right><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right><font color=brown>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
dtot = dtot + stot
bamt1 = bamt1 + bamt
fl.MoveNext
Wend



fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=1><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td  colspan=11><font color=brown>Total For the Resource - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(bamt1, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=1><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td  colspan=11><font color=brown>Total For the Resource - " & List1.List(l) & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
tot = tot + dtot
bamt2 = bamt2 + bamt1
End If
Next l
fs.WriteLine "        <tr bgcolor=yellow height=15 class=TableFont>"
If Check2.Value = 1 Then
fs.WriteLine "            <td  colspan=12><font color=brown>NET TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=brown>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=brown>" & Format(bamt2, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
Else
fs.WriteLine "            <td  colspan=12><font color=brown>NET TOTAL</td>"
fs.WriteLine "            <td  align=right><font color=brown>" & Format(tot, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td align=right ><font color=brown>&nbsp;</td>"
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
                fs.WriteLine "            <td colspan=7><b>" & GetCompanyName & "</td>"
                Else
                fs.WriteLine "            <td colspan=5><b>" & GetCompanyName & "</td>"
                End If
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>Resource</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
                    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                    If Check2.Value = 1 Then
                    fs.WriteLine "            <td colspan=7><b>BUDGET BY COSTCODE</td>"
                    Else
                    fs.WriteLine "            <td colspan=5><b>BUDGET BY COSTCODE</td>"
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
                                fs.WriteLine "            <td colspan=13>&nbsp;</td>"
                                Else
                                fs.WriteLine "            <td colspan=11>&nbsp;</td>"
                                End If
                                fs.WriteLine "        </tr>"
         



                fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                fs.WriteLine "            <td Nowrap><font color=white>RescCde</td>"
                fs.WriteLine "            <td colspan=4 ><font color=white>Resource Code Description</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td colspan=7 ><font color=white>Vendor Name</td>"
                Else
                fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
                End If
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                'fs.WriteLine "            <td Nowrap><font color=white>Year</td>"
                fs.WriteLine "            <td Nowrap align=center><font color=white>CostCde</td>"
                fs.WriteLine "            <td Nowrap><font color=white>JobCharge</td>"
                'fs.WriteLine "            <td Nowrap><font color=white>SprdCde</td>"
                'fs.WriteLine "            <td Nowrap><font color=white>TrnxType</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>TotalQty</td>"
                fs.WriteLine "            <td Nowrap align=center><font color=white>UOM</td>"
                fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>UnitRate</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>xRate</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>DwT</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>Escl</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>BDGT Amt(RM)</td>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td Nowrap align=right><font color=white>%WC</td>"
                fs.WriteLine "            <td Nowrap align=right><font color=white>BCWP Amt(RM)</td>"
                End If
                fs.WriteLine "            <td ><font color=white>Notes/CostCde Desc</td>"
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
List1.Enabled = True
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

