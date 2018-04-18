VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_progressdurationjb 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6135
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   11175
      ExtentX         =   19711
      ExtentY         =   10821
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
      TabIndex        =   1
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1200
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
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6810
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   4290
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5760
         TabIndex        =   3
         Top             =   960
         Width           =   1040
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   5535
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5745
         Top             =   120
         Width           =   5415
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
         TabIndex        =   11
         Top             =   720
         Width           =   585
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   11295
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10440
         Picture         =   "rpt_progressdurationjb.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8760
         Picture         =   "rpt_progressdurationjb.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9600
         Picture         =   "rpt_progressdurationjb.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transaction Dates"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   -120
      End
   End
End
Attribute VB_Name = "rpt_progressdurationjb"
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
Call nocolor
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
   fs.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"



        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
Dim sm3 As Double
sm3 = 0
        
         Dim w As Integer
 w = 0
 
 For w = 0 To List1.ListCount - 1
 If List1.Selected(w) = True Then
 gy = Split(List1.List(w), "  -  ", Len(List1.List(w)), vbTextCompare)
 

  cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=7><font color=white><b>Proj - " & gy(0) & " - " & gy(1) & "</td>"
                
                fs.WriteLine "        </tr>"
        
       Dim sm2 As Double
       sm2 = 0
 
        
                Dim l As Integer
                l = 0
                For l = 0 To List2.ListCount - 1
                If List2.Selected(l) = True Then
                nm = Split(List2.List(l), "  -  ", Len(List2.List(l)), vbTextCompare)

Dim rsd As New ADODB.Recordset
If rsd.State Then rsd.Close
rsd.Open "select * from progressdurationdetails b, jobcharge p where b.prgs_job_key =p.job_code and   p.job_proj_key='" & gy(0) & "' and  b.prgs_job_key like '" & nm(0) & "%'  order by b.prgs_job_key", Cn, 3, 2
If Not rsd.EOF Then

 cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
                                fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                                fs.WriteLine "            <td colspan=7><font color=black><b>Job - " & nm(0) & " - " & nm(1) & "</td>"
                                
                                fs.WriteLine "        </tr>"
End If
        
        
 Dim sm1 As Double
 
 sm1 = 0
 Dim dy As Double
 dy = 0
 Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from progressdurationdetails b, jobcharge p where b.prgs_job_key =p.job_code and   p.job_proj_key='" & gy(0) & "' and  b.prgs_job_key like '" & nm(0) & "%'  order by b.prgs_job_key", Cn, 3, 2
 
While Not rs.EOF

 cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
 
Dim jcg As New ADODB.Recordset
        If jcg.State Then jcg.Close
        jcg.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & rs!prgs_job_key & "' and job_proj_key='" & gy(0) & "'", Cn, 3, 2
        If Not jcg.EOF Then
        ko = Mid(rs!prgs_job_key & "  -  " & jcg(0), 1, 40)
        
        fs.WriteLine "            <td  >" & ko & "</td>"
        Else
        fs.WriteLine "            <td  >" & rs!prgs_job_key & "</td>"
        End If
        jcg.Close
        fs.WriteLine "            <td  >" & rs!prgs_spread_code & "</td>"
fs.WriteLine "            <td  >" & rs!prgs_type & "</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  >" & rs!prgs_startdate & "</td>"
fs.WriteLine "            <td  >" & rs!prgs_enddate & "</td>"
End If
fs.WriteLine "            <td  align=right>" & Format(rs!prgs_days, "###,###,##0.00") & "</td>"
sm1 = sm1 + rs!prgs_days
If rs!prgs_remarks <> "" Then
nt = Mid(rs!prgs_remarks, 1, 20)
fs.WriteLine "            <td  >" & nt & "</td>"
Else
fs.WriteLine "            <td  >&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
 
 rs.MoveNext
 Wend
 
 '-----------------------
 Dim rsd1 As New ADODB.Recordset
If rsd1.State Then rsd1.Close
rsd1.Open "select * from progressdurationdetails b, jobcharge p where b.prgs_job_key =p.job_code and   p.job_proj_key='" & gy(0) & "' and  b.prgs_job_key like '" & nm(0) & "%'  order by b.prgs_job_key", Cn, 3, 2
If Not rsd1.EOF Then
 
 
   cnt = cnt + 1 '********************************
                                  If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td colspan=5 ><b>SubTotal   -    " & nm(0) & " </td>"
Else
fs.WriteLine "            <td colspan=3 ><b>SubTotal   -    " & nm(0) & " </td>"
End If
fs.WriteLine "            <td  align=right><b>" & Format(sm1, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  >&nbsp;</td>"
 
fs.WriteLine "        </tr>"
 sm2 = sm2 + sm1
 
 End If
 '------------------------
 
 


  End If
  Next l
  
  
   cnt = cnt + 1 '********************************
                                   If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
                                
                                    '-----------------------
 
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=5><b>SubTotal   -    " & gy(0) & " - " & gy(1) & " </td>"
Else
fs.WriteLine "            <td  colspan=3><b>SubTotal   -    " & gy(0) & " - " & gy(1) & " </td>"
End If
fs.WriteLine "            <td  align=right><b>" & Format(sm2, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  >&nbsp;</td>"
 
fs.WriteLine "        </tr>"
 sm3 = sm3 + sm2
 '------------------------
                                
                                

   
  End If
  Next w
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    '-----------------------
 
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=5><b><font color=white  >Report Total</td>"
Else
fs.WriteLine "            <td  colspan=3><b><font color=white  >Report Total</td>"
End If
fs.WriteLine "            <td  align=right><b><font color=white>" & Format(sm3, "###,###,##0.00") & "</td>"
 
fs.WriteLine "            <td  >&nbsp;</td>"
 
fs.WriteLine "        </tr>"
 sm3 = sm3 + sm2
 '------------------------


   fs.WriteLine " </table>"
   fs.WriteLine "    <table border=0 cellspacing=0 BORDERCOLOR=gray width=95%>"


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
main.lbltitle.Caption = "ESTIMATED PROGRESS DURATION BY PROJECT"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not ls.EOF
List1.AddItem ls(0) & "  -  " & ls(1)
ls.MoveNext
Wend
ls.Close
Dim hgg As Integer
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
            Check4.Value = 1
            
Me.Width = 11415
Me.Height = 9750

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Public Sub RPTHEADING(fs As Object)
 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
           
             
            If Check4.Value = 1 Then
             fs.WriteLine "           <td colspan=7><b>" & GetCompanyName & "</td>"
             Else
              fs.WriteLine "           <td colspan=5><b>" & GetCompanyName & "</td>"
            End If
            fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
            fs.WriteLine "           <td colspan=4><b>Progress Duration By Project</td>"
            
            fs.WriteLine "           <td colspan=1 align=center><b>ReportDate" & Format(Date, "dd/MM/yyyy") & "</td>"
            If Check4.Value = 1 Then
            fs.WriteLine "           <td colspan=2 align=center><b>&nbsp;</td>"
            End If
            fs.WriteLine "        </tr>"
   
   
   fs.WriteLine "        <tr bgcolor=Black height=15 class=TableFont>"
 
   fs.WriteLine "            <td Nowrap align=center><font color=white>JobCharge</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Spread</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Type</td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=center><font color=white>StartDate</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Enddate</td>"
   End If
   fs.WriteLine "            <td Nowrap align=center><font color=white>Days</td>"
   fs.WriteLine "            <td width=130><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
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


If Option4.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List2.ListCount - 1
List2.Selected(f) = True
Next f

End If
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
 

