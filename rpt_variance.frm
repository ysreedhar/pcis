VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_variance 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   10095
      ExtentX         =   17806
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11295
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6800
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   4290
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5760
         TabIndex        =   6
         Top             =   960
         Width           =   1040
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1200
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key"
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
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   6120
         TabIndex        =   9
         Top             =   720
         Width           =   585
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5745
         Top             =   120
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   120
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10320
         Picture         =   "rpt_variance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8640
         Picture         =   "rpt_variance.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9480
         Picture         =   "rpt_variance.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Job Level"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "JobCharge Level"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         Begin VB.OptionButton opt_year 
            BackColor       =   &H00FFC0C0&
            Caption         =   " @ YearEnd"
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton opt_cut 
            BackColor       =   &H00FFC0C0&
            Caption         =   "@ Cuttoff Date"
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   3
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   120
      End
   End
End
Attribute VB_Name = "rpt_variance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
If opt_year.Value = True Then
If Option5.Value = True Then
Call nocolor
Else
Call sumvari
End If
ElseIf opt_cut.Value = True Then
If Option5.Value = True Then
Call nocolor1
Else
Call sumvari1
End If
Else
MsgBox "Select Option"
Exit Sub
End If
Unload frmBusy
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
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"



        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
Dim sm3 As Double
sm3 = 0
        
         Dim w As Integer
 w = 0
 Dim sm1111 As Double
sm1111 = 0
Dim sm2222 As Double
sm2222 = 0
 For w = 0 To List1.ListCount - 1
 If List1.Selected(w) = True Then
 gy = Split(List1.List(w), "  -  ", Len(List1.List(w)), vbTextCompare)
Dim sm111 As Double
sm111 = 0
Dim sm222 As Double
sm222 = 0

  cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=6><font color=white><b>Proj - " & gy(0) & " - " & gy(1) & "</td>"
                
                fs.WriteLine "        </tr>"
        

 
        
                Dim l As Integer
                l = 0
                For l = 0 To List2.ListCount - 1
                If List2.Selected(l) = True Then
                nm = Split(List2.List(l), "  -  ", Len(List2.List(l)), vbTextCompare)
Dim sm11 As Double
sm11 = 0
Dim sm22 As Double
sm22 = 0
  Dim ttt As New ADODB.Recordset
        If ttt.State Then ttt.Close
        ttt.Open "select DISTINCT(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not ttt.EOF Then
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
                                fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                                fs.WriteLine "            <td colspan=6><font color=black><b>Job - " & nm(0) & " - " & nm(1) & "</td>"
                                fs.WriteLine "        </tr>"
 
        End If

Dim sm1 As Double
sm1 = 0
Dim sm2 As Double
sm2 = 0
    Dim tt As New ADODB.Recordset
        If tt.State Then tt.Close
        tt.Open "select DISTINCt(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        While Not tt.EOF
Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim fllg As Integer
fllg = 0
Dim dy As Double
dy = 0
Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select SUM(b.bdgt_days) ,j.job_code from budgeteddurationdetails b  , jobcharge j where b.bdgt_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  b.bdgt_job_key  = '" & tt(0) & "'  Group by j.job_code", Cn, 3, 2
 
If Not rs.EOF Then
fllg = 1
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

Dim jcg As New ADODB.Recordset
        If jcg.State Then jcg.Close
        jcg.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not jcg.EOF Then
        ko = Mid(rs(1) & "  -  " & jcg(0), 1, 40)
        
        fs.WriteLine "            <td  >" & ko & "</td>"
        Else
        fs.WriteLine "            <td  >" & rs(1) & "</td>"
        End If
        jcg.Close
fs.WriteLine "            <td align=right >" & Format(rs(0), "###,###,##0.00") & "</td>"
x1 = rs(0)
 
End If
Dim rsn As New ADODB.Recordset
If rsn.State Then rsn.Close
rsn.Open "select SUM(p.prgs_days) ,j.job_code from progressdurationdetails p  , jobcharge j where p.prgs_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  p.prgs_job_key = '" & tt(0) & "'  Group by    j.job_code", Cn, 3, 2
 
If Not rsn.EOF Then
 If fllg = 0 Then
 '--------------------------
 fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim jcg1 As New ADODB.Recordset
        If jcg1.State Then jcg1.Close
        jcg1.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not jcg1.EOF Then
       
        
        fs.WriteLine "            <td  >" & tt(0) & "  -  " & jcg1(0) & "</td>"
      
        End If
        jcg1.Close
fs.WriteLine "            <td align=right >0.00</td>"
 
 
 '--------------------------
 End If
fs.WriteLine "            <td align=right >" & Format(rsn(0), "###,###,##0.00") & "</td>"
x2 = rsn(0)
Else
If x1 <> 0 Then
fs.WriteLine "            <td align=right >0.00</td>"
End If
End If
                Dim jca As New ADODB.Recordset
                If jca.State Then jca.Close
                jca.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jca.EOF Then
                 If x2 = 0 And x1 = 0 Then
                 Else
                fs.WriteLine "            <td align=right >" & Format(CDbl(x2) - CDbl(x1), "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                End If
                End If
'fs.WriteLine "            <td  align=right>&nbsp;</td>"
sm1 = sm1 + x1
sm2 = sm2 + x2
fs.WriteLine "        </tr>"


 tt.MoveNext
 Wend
sm11 = sm11 + sm1
sm22 = sm22 + sm2
                 Dim jcb As New ADODB.Recordset
                If jcb.State Then jcb.Close
                jcb.Open "select * from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcb.EOF Then
                                If sm11 = 0 And sm22 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b>SubTotal</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm22, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm22 - sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
 
sm111 = sm111 + sm11
sm222 = sm222 + sm22

  End If
  Next l
  
  
  

                 Dim jcc As New ADODB.Recordset
                If jcc.State Then jcc.Close
                jcc.Open "select * from jobcharge where  job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcc.EOF Then
                                If sm111 = 0 And sm222 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b> Total - " & gy(0) & "</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222 - sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
sm1111 = sm1111 + sm111
sm2222 = sm2222 + sm222

   
  End If
  Next w
  
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    '-----------------------
                                  fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><font color=white><b>Report Total </td>"
                                fs.WriteLine "            <td  align=right><font color=white> <b> " & Format(sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222 - sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"


   fs.WriteLine " </table>"
   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"


Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>Project</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r




Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b>JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f




fs.WriteLine " </table>"

   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "DURATION VARIANCE BY PROJECT"
Me.Top = 10
Me.Left = 10
Option5.Value = True
WebBrowser.Navigate "About:Blank"
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not ls.EOF
List1.AddItem ls(0) & "  -  " & ls(1)
ls.MoveNext
Wend
ls.Close

Me.Width = 11415
Me.Height = 9750

End Sub

Private Sub List1_Click()
 List2.Clear
 Option3.Value = True
Dim h As Integer
h = 0
For h = 0 To List1.ListCount - 1
If List1.Selected(h) = True Then
ju = Split(List1.List(h), "  -  ", Len(List1.List(h)), vbTextCompare)
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & ju(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
List2.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
End If
Next h
End Sub

Private Sub Option1_Click()
Option3.Value = 0
Option4.Value = 0
                hgg = 0
                For hgg = 0 To List2.ListCount - 1
                List2.Selected(hgg) = False
                Next hgg
If Option1.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List1.ListCount - 1
List1.Selected(f) = True
Next f

End If

End Sub

Private Sub Option2_Click()
Option3.Value = 0
Option4.Value = 0
                    hgg = 0
                    For hgg = 0 To List2.ListCount - 1
                    List2.Selected(hgg) = False
                    Next hgg
If Option2.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List1.ListCount - 1
List1.Selected(g) = False
Next g

End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List2.ListCount - 1
List2.Selected(g) = False
Next g

End If

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List2.ListCount - 1
List2.Selected(f) = True
Next f

End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Public Sub RPTHEADING(fs As Object)
 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
            fs.WriteLine "           <td colspan=6><b>" & GetCompanyName & "</td>"

            fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
            fs.WriteLine "           <td colspan=4><b>Duration Variance By Project</td>"
            fs.WriteLine "           <td colspan=1 align=center><b>ReportDate</td>"
            fs.WriteLine "           <td colspan=1 align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
           
            fs.WriteLine "        </tr>"
   
   
   fs.WriteLine "        <tr bgcolor=Black height=15 class=TableFont>"
 
   fs.WriteLine "            <td Nowrap align=center><font color=white>JobCharge</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Bdgt</td>"
    
   fs.WriteLine "            <td Nowrap align=center><font color=white>Actual</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Var</td>"
    
    
   fs.WriteLine "            <td colspan=2><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
End Sub

Public Sub sumvari()
Dim fs As Object
Dim fso As New FileSystemObject
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
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"



        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
Dim sm3 As Double
sm3 = 0
        
         Dim w As Integer
 w = 0
 Dim sm1111 As Double
sm1111 = 0
Dim sm2222 As Double
sm2222 = 0
 For w = 0 To List1.ListCount - 1
 If List1.Selected(w) = True Then
 gy = Split(List1.List(w), "  -  ", Len(List1.List(w)), vbTextCompare)
Dim sm111 As Double
sm111 = 0
Dim sm222 As Double
sm222 = 0

  cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=6><font color=white><b>Proj - " & gy(0) & " - " & gy(1) & "</td>"
                
                fs.WriteLine "        </tr>"
        

 
        
                Dim l As Integer
                l = 0
                For l = 0 To List2.ListCount - 1
                If List2.Selected(l) = True Then
                nm = Split(List2.List(l), "  -  ", Len(List2.List(l)), vbTextCompare)
Dim sm11 As Double
sm11 = 0
Dim sm22 As Double
sm22 = 0
 

Dim sm1 As Double
sm1 = 0
Dim sm2 As Double
sm2 = 0
    Dim tt As New ADODB.Recordset
        If tt.State Then tt.Close
        tt.Open "select DISTINCt(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        While Not tt.EOF
Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim fllg As Integer
fllg = 0
Dim dy As Double
dy = 0
Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select SUM(b.bdgt_days) ,j.job_code from budgeteddurationdetails b  , jobcharge j where b.bdgt_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  b.bdgt_job_key  = '" & tt(0) & "'  Group by j.job_code", Cn, 3, 2
 
If Not rs.EOF Then
 

 
x1 = rs(0)
 
End If
Dim rsn As New ADODB.Recordset
If rsn.State Then rsn.Close
rsn.Open "select SUM(p.prgs_days) ,j.job_code from progressdurationdetails p  , jobcharge j where p.prgs_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  p.prgs_job_key = '" & tt(0) & "'  Group by    j.job_code", Cn, 3, 2
 
If Not rsn.EOF Then
 
x2 = rsn(0)
 
End If
 
sm1 = sm1 + x1
sm2 = sm2 + x2
fs.WriteLine "        </tr>"


 tt.MoveNext
 Wend
sm11 = sm11 + sm1
sm22 = sm22 + sm2
                 Dim jcb As New ADODB.Recordset
                If jcb.State Then jcb.Close
                jcb.Open "select * from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcb.EOF Then
                                If sm11 = 0 And sm22 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  >  " & nm(0) & "  -  " & nm(1) & "</td>"
                                fs.WriteLine "            <td  align=right>  " & Format(sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right>  " & Format(sm22, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right>  " & Format(sm22 - sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
 
sm111 = sm111 + sm11
sm222 = sm222 + sm22

  End If
  Next l
  
  
  

                 Dim jcc As New ADODB.Recordset
                If jcc.State Then jcc.Close
                jcc.Open "select * from jobcharge where  job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcc.EOF Then
                                If sm111 = 0 And sm222 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b> Total - " & gy(0) & "  -  " & gy(1) & "</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222 - sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
sm1111 = sm1111 + sm111
sm2222 = sm2222 + sm222

   
  End If
  Next w
  
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    '-----------------------
                                  fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><font color=white><b>Report Total </td>"
                                fs.WriteLine "            <td  align=right><font color=white> <b> " & Format(sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222 - sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"


   fs.WriteLine " </table>"
   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"


Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>Project</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r




Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b>JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f




fs.WriteLine " </table>"

   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub


Public Sub nocolor1()
Dim fs As Object
Dim fso As New FileSystemObject
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
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"



        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
Dim sm3 As Double
sm3 = 0
        
         Dim w As Integer
 w = 0
 Dim sm1111 As Double
sm1111 = 0
Dim sm2222 As Double
sm2222 = 0
 For w = 0 To List1.ListCount - 1
 If List1.Selected(w) = True Then
 gy = Split(List1.List(w), "  -  ", Len(List1.List(w)), vbTextCompare)
Dim sm111 As Double
sm111 = 0
Dim sm222 As Double
sm222 = 0

  cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=6><font color=white><b>Proj - " & gy(0) & " - " & gy(1) & "</td>"
                
                fs.WriteLine "        </tr>"
        

 
        
                Dim l As Integer
                l = 0
                For l = 0 To List2.ListCount - 1
                If List2.Selected(l) = True Then
                nm = Split(List2.List(l), "  -  ", Len(List2.List(l)), vbTextCompare)
Dim sm11 As Double
sm11 = 0
Dim sm22 As Double
sm22 = 0
  Dim ttt As New ADODB.Recordset
        If ttt.State Then ttt.Close
        ttt.Open "select DISTINCT(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not ttt.EOF Then
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
                                fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                                fs.WriteLine "            <td colspan=6><font color=black><b>Job - " & nm(0) & " - " & nm(1) & "</td>"
                                fs.WriteLine "        </tr>"
 
        End If

Dim sm1 As Double
sm1 = 0
Dim sm2 As Double
sm2 = 0
    Dim tt As New ADODB.Recordset
        If tt.State Then tt.Close
        tt.Open "select DISTINCT(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        While Not tt.EOF
Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim fllg As Integer
fllg = 0
Dim dy As Double
dy = 0
Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select SUM(b.bdgt_days) ,j.job_code from budgeteddurationdetails b  , jobcharge j where b.bdgt_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  b.bdgt_job_key  = '" & tt(0) & "'  Group by j.job_code", Cn, 3, 2
 
If Not rs.EOF Then
fllg = 1
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"

Dim jcg As New ADODB.Recordset
        If jcg.State Then jcg.Close
        jcg.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not jcg.EOF Then
        ko = Mid(rs(1) & "  -  " & jcg(0), 1, 40)
        
        fs.WriteLine "            <td  >" & ko & "</td>"
        Else
        fs.WriteLine "            <td  >" & rs(1) & "</td>"
        End If
        jcg.Close
Dim aa As Double
Dim bb As Double
aa = 0: bb = 0
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select bd_wrkcomp from cost where bd_costtype='B' and bd_jobcharge='" & tt(0) & "'", Cn, 3, 2
If Not bd.EOF Then

aa = rs(0) * (bd(0) / 100)

End If

 fs.WriteLine "            <td align=right >" & Format(aa, "###,###,##0.00") & "</td>"
 x1 = aa
End If
Dim flgp As Integer
flgp = 0
Dim rsn As New ADODB.Recordset
If rsn.State Then rsn.Close
rsn.Open "select p.prgs_startdate ,p.prgs_enddate,j.job_code from progressdurationdetails p  , jobcharge j where p.prgs_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  p.prgs_job_key = '" & tt(0) & "'  ", Cn, 3, 2
 If Not rsn.EOF Then
While Not rsn.EOF
 If fllg = 0 Then
 flgp = 1
 '--------------------------
 fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 Dim jcg1 As New ADODB.Recordset
        If jcg1.State Then jcg1.Close
        jcg1.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not jcg1.EOF Then
              
        fs.WriteLine "            <td  >" & tt(0) & "  -  " & jcg1(0) & "</td>"
      
        End If
        jcg1.Close
fs.WriteLine "            <td align=right >0.00</td>"
 
 '--------------------------
 End If
 Dim a As Double
 Dim c As Double
 a = 0: c = 0
 
 'duration calculation
        If rsn(0) <= main.DTPcutdate1.Value And rsn(1) <= main.DTPcutdate1.Value Then
        a = rsn(1) - rsn(0)
        c = 0
        ElseIf rsn(0) <= main.DTPcutdate1.Value And rsn(1) >= main.DTPcutdate1.Value Then
        a = main.DTPcutdate1.Value - rsn(0)
        c = rsn(1) - main.DTPcutdate1.Value
        
        Else
        a = 0
        c = rsn(1) - rsn(0)
        End If
 
 

x2 = x2 + a

rsn.MoveNext
Wend
fs.WriteLine "            <td align=right >" & Format(x2, "###,###,##0.00") & "</td>"
Else

If fllg <> 0 Then
fs.WriteLine "            <td align=right >0.00</td>"
End If
End If
                Dim jca As New ADODB.Recordset
                If jca.State Then jca.Close
                jca.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & tt(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jca.EOF Then
                 If x2 = 0 And x1 = 0 Then
                    If fllg = 1 Or flgp = 1 Then
                        fs.WriteLine "            <td align=right >" & Format(CDbl(x2) - CDbl(x1), "###,###,##0.00") & "</td>"
                        fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                    End If
                 Else
                fs.WriteLine "            <td align=right >" & Format(CDbl(x2) - CDbl(x1), "###,###,##0.00") & "</td>"
                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                End If
                End If
'fs.WriteLine "            <td  align=right>&nbsp;</td>"
sm1 = sm1 + x1
sm2 = sm2 + x2
fs.WriteLine "        </tr>"


 tt.MoveNext
 Wend
sm11 = sm11 + sm1
sm22 = sm22 + sm2
                 Dim jcb As New ADODB.Recordset
                If jcb.State Then jcb.Close
                jcb.Open "select * from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcb.EOF Then
                                If sm11 = 0 And sm22 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b>SubTotal</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm22, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm22 - sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
 
sm111 = sm111 + sm11
sm222 = sm222 + sm22

  End If
  Next l
  
  
  

                 Dim jcc As New ADODB.Recordset
                If jcc.State Then jcc.Close
                jcc.Open "select * from jobcharge where  job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcc.EOF Then
                                If sm111 = 0 And sm222 = 0 Then
                                Else
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b> Total - " & gy(0) & "</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222 - sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
sm1111 = sm1111 + sm111
sm2222 = sm2222 + sm222

   
  End If
  Next w
  
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    '-----------------------
                                  fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><font color=white><b>Report Total </td>"
                                fs.WriteLine "            <td  align=right><font color=white> <b> " & Format(sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222 - sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"


   fs.WriteLine " </table>"
   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"


Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>Project</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r




Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b>JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f




fs.WriteLine " </table>"

   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Public Sub sumvari1()
Dim fs As Object
Dim fso As New FileSystemObject
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
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"



        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
Dim sm3 As Double
sm3 = 0
        
         Dim w As Integer
 w = 0
 Dim sm1111 As Double
sm1111 = 0
Dim sm2222 As Double
sm2222 = 0
 For w = 0 To List1.ListCount - 1
 If List1.Selected(w) = True Then
 gy = Split(List1.List(w), "  -  ", Len(List1.List(w)), vbTextCompare)
Dim sm111 As Double
sm111 = 0
Dim sm222 As Double
sm222 = 0

  cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=6><font color=white><b>Proj - " & gy(0) & " - " & gy(1) & "</td>"
                
                fs.WriteLine "        </tr>"
        

 
        
                Dim l As Integer
                l = 0
                For l = 0 To List2.ListCount - 1
                If List2.Selected(l) = True Then
                nm = Split(List2.List(l), "  -  ", Len(List2.List(l)), vbTextCompare)
Dim sm11 As Double
sm11 = 0
Dim sm22 As Double
sm22 = 0
 

Dim sm1 As Double
sm1 = 0
Dim sm2 As Double
sm2 = 0
    Dim tt As New ADODB.Recordset
        If tt.State Then tt.Close
        tt.Open "select DISTINCT(job_code) from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        While Not tt.EOF
Dim x1 As Double
Dim x2 As Double
x1 = 0: x2 = 0
Dim fllg As Integer
fllg = 0
Dim dy As Double
dy = 0
Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select SUM(b.bdgt_days) ,j.job_code from budgeteddurationdetails b  , jobcharge j where b.bdgt_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  b.bdgt_job_key  = '" & tt(0) & "'  Group by j.job_code", Cn, 3, 2
 
If Not rs.EOF Then
fllg = 1


 
Dim aa As Double
Dim bb As Double
aa = 0: bb = 0
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select bd_wrkcomp from cost where bd_costtype='B' and bd_jobcharge='" & tt(0) & "'", Cn, 3, 2
If Not bd.EOF Then

aa = rs(0) * (bd(0) / 100)

End If
 
 x1 = aa
End If
Dim flgp As Integer
flgp = 0
Dim rsn As New ADODB.Recordset
If rsn.State Then rsn.Close
rsn.Open "select p.prgs_startdate ,p.prgs_enddate,j.job_code from progressdurationdetails p  , jobcharge j where p.prgs_job_key = j.job_code   and   j.job_proj_key='" & gy(0) & "' and  p.prgs_job_key = '" & tt(0) & "'  ", Cn, 3, 2
 
While Not rsn.EOF
 If fllg = 0 Then
 flgp = 1
 '--------------------------
 
 
 
 '--------------------------
 End If
 Dim a As Double
 Dim c As Double
 a = 0: c = 0
 
 'duration calculation
        If rsn(0) <= main.DTPcutdate1.Value And rsn(1) <= main.DTPcutdate1.Value Then
        a = rsn(1) - rsn(0)
        c = 0
        ElseIf rsn(0) <= main.DTPcutdate1.Value And rsn(1) >= main.DTPcutdate1.Value Then
        a = main.DTPcutdate1.Value - rsn(0)
        c = rsn(1) - main.DTPcutdate1.Value
        
        Else
        a = 0
        c = rsn(1) - rsn(0)
        End If
 
 

x2 = x2 + a

rsn.MoveNext
Wend
 
 
 
sm1 = sm1 + x1
sm2 = sm2 + x2
fs.WriteLine "        </tr>"


 tt.MoveNext
 Wend
sm11 = sm11 + sm1
sm22 = sm22 + sm2
                 Dim jcb As New ADODB.Recordset
                If jcb.State Then jcb.Close
                jcb.Open "select * from jobcharge where jobno='" & nm(0) & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcb.EOF Then
                                If sm11 = 0 And sm22 = 0 Then
                                Else
    cnt = cnt + 1 '********************************
    If cnt >= 58 Then
    fs.WriteLine "</table><P></P>"
    RPTHEADING fs
    cnt = 0
    End If
 
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  > " & nm(0) & " - " & nm(1) & "</td>"
                                fs.WriteLine "            <td  align=right>   " & Format(sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right>  " & Format(sm22, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right>  " & Format(sm22 - sm11, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
 
sm111 = sm111 + sm11
sm222 = sm222 + sm22

  End If
  Next l
  
  
  

                 Dim jcc As New ADODB.Recordset
                If jcc.State Then jcc.Close
                jcc.Open "select * from jobcharge where  job_proj_key='" & gy(0) & "'", Cn, 3, 2
                If Not jcc.EOF Then
                                If sm111 = 0 And sm222 = 0 Then
                                Else
 cnt = cnt + 1 '********************************
If cnt >= 58 Then
fs.WriteLine "</table><P></P>"
RPTHEADING fs
cnt = 0
End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><b> Total - " & gy(0) & "</td>"
                                fs.WriteLine "            <td  align=right> <b> " & Format(sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><b> " & Format(sm222 - sm111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"
                                End If
                 End If
sm1111 = sm1111 + sm111
sm2222 = sm2222 + sm222

   
  End If
  Next w
  
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    '-----------------------
                                  fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                                fs.WriteLine "            <td  ><font color=white><b>Report Total </td>"
                                fs.WriteLine "            <td  align=right><font color=white> <b> " & Format(sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td align=right><font color=white><b> " & Format(sm2222 - sm1111, "###,###,##0.00") & "</td>"
                                fs.WriteLine "            <td  colspan=2>&nbsp;</td>"
                                fs.WriteLine "        </tr>"


   fs.WriteLine " </table>"
   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"


Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>Project</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r




Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b>JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f




fs.WriteLine " </table>"

   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub
