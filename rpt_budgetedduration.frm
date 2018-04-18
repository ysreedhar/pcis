VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_budgetedduration 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6975
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   10815
      ExtentX         =   19076
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
      BackColor       =   &H00DC7E5A&
      Caption         =   "Budgeted Duration"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10320
         Picture         =   "rpt_budgetedduration.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to Exit"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8520
         Picture         =   "rpt_budgetedduration.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to View"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9420
         Picture         =   "rpt_budgetedduration.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click to Print"
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cbo_job 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Spread - Description"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "rpt_budgetedduration"
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
If cbo_job.Text = "" Then
MsgBox "Select Spread"
Exit Sub
End If
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
   
   nm = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)

        Dim cnt As Integer
        RPTHEADING fs
        cnt = 0
        Dim sn As Integer
        sn = 1
 Dim dy As Double
 dy = 0
 Dim ko As String
Dim nt As String
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from budgeteddurationdetails where bdgt_spread_code='" & nm(0) & "' order by bdgt_job_key", Cn, 3, 2
While Not rs.EOF
 cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
Dim jcg As New ADODB.Recordset
        If jcg.State Then jcg.Close
        jcg.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & rs!bdgt_job_key & "' ", Cn, 3, 2
        If Not jcg.EOF Then
        ko = Mid(rs!bdgt_job_key & "  -  " & jcg(0), 1, 60)
        
        fs.WriteLine "            <td  >" & ko & "</td>"
        Else
        fs.WriteLine "            <td  >" & rs!bdgt_job_key & "</td>"
        End If
        jcg.Close


fs.WriteLine "            <td  align=right>" & Format(rs!bdgt_days, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td  align=right>" & Format(rs!bdgt_per_workcomplete, "###,###,##0.00") & "</td>"
dy = dy + rs!bdgt_days
If rs!bdgt_remarks <> "" Then
nt = Mid(rs!bdgt_remarks, 1, 20)
fs.WriteLine "            <td  >" & nt & "</td>"
Else
fs.WriteLine "            <td  >&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
 sn = sn + 1
 rs.MoveNext
 Wend
 
 
  cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=2><font color=white>Total Duration Days</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(dy, "###,###,##0.00") & "</td>"
fs.WriteLine "            <td  colspan=2><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
 
 
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BUDGETED DURATION BY SPREAD"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim tr As New ADODB.Recordset
            If tr.State Then tr.Close
            tr.Open "select DISTINCT(b.bdgt_spread_code),s.spread_desc   from budgeteddurationdetails b,spreadmaster s where b.bdgt_spread_code=s.spread_code order by b.bdgt_spread_code", Cn, 3, 2
            While Not tr.EOF
            cbo_job.AddItem tr(0) & "  -  " & tr(1)
            tr.MoveNext
            Wend
tr.Close
Me.Width = 11415
Me.Height = 9750

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Public Sub RPTHEADING(fs As Object)
 fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
            fs.WriteLine "           <td colspan=3><b>" & GetCompanyName & "</td>"
            fs.WriteLine "           <td colspan=1 align=center><b>Spread</td>"
            fs.WriteLine "           <td colspan=1 align=center>" & cbo_job.Text & "</td>"
            fs.WriteLine "        </tr>"
            fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
            fs.WriteLine "           <td colspan=3><b>Budgeted Duration</td>"
            fs.WriteLine "           <td colspan=1 align=center><b>ReportDate</td>"
            fs.WriteLine "           <td colspan=1 align=center>" & Format(Date, "dd/MM/yyyy") & "</td>"
            fs.WriteLine "        </tr>"
   
   
   fs.WriteLine "        <tr bgcolor=Black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>JobCharge</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Days</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>%WC</td>"
   fs.WriteLine "            <td width=180><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
End Sub
