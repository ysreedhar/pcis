VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_l2 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      Caption         =   "JobNo - Description"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_l2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to Print"
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_l2.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to View"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9960
         Picture         =   "rpt_l2.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to Exit"
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9960
         Picture         =   "rpt_l2.frx":118D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to Save"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DC7E5A&
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox cbo_job 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key - Description"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7995
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   11145
      ExtentX         =   19659
      ExtentY         =   14102
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
End
Attribute VB_Name = "rpt_l2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim Mode As Integer
Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmd_print_Click()
'''On Error GoTo XIT
'''WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
'''XIT:
On Error GoTo errhandler
Dim pr As Object
Set pr = Printer
pr.Orientation = vbPRORLandscape
   WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
'OLECMDID_PRINT
errhandler:
End Sub

Private Sub cmd_save_Click()
Load filepath

End Sub


Private Sub cmd_show_Click()
If cbo_job.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
If Check1.Value = 1 Then
Call progcost
End If
Call rephtml
Unload frmBusy
 
End Sub
Private Sub Command1_Click()
Call rephtml
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT:
End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "L2 - PRCR @ JOBCHARGE LEVEL - BY PROJECT KEY"
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

    Me.Width = 11415
    Me.Height = 9750


End Sub
Public Sub rephtml()
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   
   fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\FeatherTexture.bmp);"
    
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
'''   fs.WriteLine "    <center>"

'''    fs.WriteLine "           <font size=2.5>PROJECT REVENUE & COST REPORT - LEVEL 2</font><BR>"
  fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=#acacac width=95%>"
 
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=11><b>TL OFFSHORE SDN BHD</td>"
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL</td>"
                fs.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
 
 
    
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs.WriteLine "            <td align=center><font color=white >Baseline Budget</td>"
                fs.WriteLine "            <td align=center><font color=white >Estimate @ Completiion</td>"
                fs.WriteLine "            <td align=center colspan=5><font color=white>Cummulative To Date</td>"
                fs.WriteLine "            <td align=center><font color=white >EstimateTo Complete</td>"
                fs.WriteLine "            <td align=center><font color=white >ProjToDate LastYrEnd</td>"
                fs.WriteLine "            <td align=center><font color=white >YearToDate LastMthEnd</td>"
                fs.WriteLine "            <td align=center><font color=white >YrToDate CurrentYear</td>"
                fs.WriteLine "            <td align=center><font color=white >ChangesIn CurrentMth</td>"
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>JobNo.Description</td>"
                fs.WriteLine "            <td align=center><font color=white>jobNo.</td>"
                fs.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td align=center><font color=white  >Revenue</td>"
                fs.WriteLine "            <td nowrap align=center><font color=white align=center>%EAC</td>"
                fs.WriteLine "            <td align=center><font color=white  >BCWP</td>"
                fs.WriteLine "            <td align=center><font color=white  >ACWP</td>"
                fs.WriteLine "            <td align=center><font color=white >CostVar</td>"
                fs.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "        </tr>"

   Dim l As Integer
    
   nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
   Dim a1 As Double
   Dim a2 As Double
   Dim a3 As Double
   Dim a4 As Double
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  SUM(rev_totamount)  from revenue where rev_projcode='" & nn(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select  SUM(rev_totamount)  from revenue where rev_projcode='" & nn(0) & "'  and rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='UBL' ", Cn, 3, 2
   If Not rv4.EOF Then
   a5 = rv4(0)
   End If
                
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "'", Cn, 3, 2
   While Not pt.EOF
        bpdl = bpdl + pt!ptd_lye_revn
        bydl = bydl + pt!ytd_lme_revn
        updl = updl + pt!ptd_lye_revn1
        uydl = uydl + pt!ytd_lme_revn1
   pt.MoveNext
   Wend
   
                
                
                
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=16><b><u>" & cbo_job.Text & "</td>"
                fs.WriteLine "        </tr>"
    ''''''' one
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE-BUDGETED</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(((a1 + a2 + a3) - (a5)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right >NA</td>"
                fs.WriteLine "        </tr>"
                
                
     ''''''' two
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE - VO(+)</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
  ''''''' three
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE- VO(-)</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
  ''''''' four
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE-BILLED</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a4, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(bpdl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(bydl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format((a4 - bpdl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(((a4 - bpdl) - bydl), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
 ''''''' five
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE-UNBILLED</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a5 - a4, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(updl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(uydl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(((a5 - a4) - updl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format((((a5 - a4) - updl) - uydl), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
                
  ''''total
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL REVENUE</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(a1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((a5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - (a5)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                revt1 = (a4 - bpdl) + ((a5 - a4) - updl)
                fs.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt1), "###,###,##0") & "</td>"
                revt2 = (((a4 - bpdl) - bydl) + (((a5 - a4) - updl) - uydl))
                fs.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt2), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
                
                Dim k1 As Double
                Dim k2 As Double
                Dim k3 As Double
                Dim k4 As Double
                Dim k5 As Double
                Dim k6 As Double
                Dim k7 As Double
                k1 = 0: k2 = 0: k3 = 0: k4 = 0: k5 = 0: k6 = 0: k7 = 0
                Dim ptt As Double
                Dim yt As Double
                ptt = 0: yt = 0
                 Dim cyt As Double
                Dim cct As Double
                cyt = 0: cct = 0
                nf = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
                Dim sl As New ADODB.Recordset
                If sl.State Then sl.Close
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' order by jobno_code", Cn, 3, 2
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=16><b><u>" & cbo_job.Text & "</td>"
                fs.WriteLine "        </tr>"
                While Not sl.EOF
                Dim bdg As Double
                Dim bcw As Double
                Dim acw As Double
                Dim ect As Double
               
                bdg = 0: bcw = 0: acw = 0: ect = 0
                          Dim ct As New ADODB.Recordset
                          If ct.State Then ct.Close
                          ct.Open "select SUM(bd_extdamt),SUM(bd_bcwpamt) from jobcharge j, cost c where j.job_code=c.bd_jobcharge and j.jobno='" & sl(0) & "' and j.job_proj_key='" & nf(0) & "' and c.bd_costtype='B'  ", Cn, 3, 2
                          If Not ct.EOF Then
                          bdg = ct(0)
                          bcw = ct(1)
                                        
                          End If
                          
                          Dim ct1 As New ADODB.Recordset
                          If ct1.State Then ct1.Close
                          ct1.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c where j.job_code=c.bd_jobcharge and j.jobno='" & sl(0) & "' and j.job_proj_key='" & nf(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not ct1.EOF Then
                          acw = ct1(0)
                          ect = ct1(1)
                                    
                          End If
                Dim ytd As Double
                Dim ptd As Double
                ytd = 0: ptd = 0
                Dim ctr As New ADODB.Recordset
                If ctr.State Then ctr.Close
                ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where jobno='" & sl(0) & "' and projkey='" & nf(0) & "'", Cn, 3, 2
                If Not ctr.EOF Then
                ytd = ctr(0)
                ptd = ctr(1)
                End If
                
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & sl(1) & "</td>"
                fs.WriteLine "            <td align=center>" & sl(0) & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(bdg, "###,###,##0") & "</td>"
                k1 = k1 + bdg
                fs.WriteLine "            <td align=right nowrap>" & Format((acw + ect), "###,###,##0") & "</td>"
                k2 = k2 + (acw + ect)
                fs.WriteLine "            <td align=right> NA </td>"
                If Round(Format(acw / (acw + ect)), 3) = 0 Then
                fs.WriteLine "            <td align=right>0</td>"
                Else
                fs.WriteLine "            <td align=right nowrap>" & Round(Format((acw / (acw + ect))) * 100, 1) & "</td>"
                End If
                fs.WriteLine "            <td align=right nowrap>" & Format(bcw, "###,###,##0") & "</td>"
                k4 = k4 + bcw
                fs.WriteLine "            <td align=right nowrap>" & Format(acw, "###,###,##0") & "</td>"
                k5 = k5 + acw
                fs.WriteLine "            <td align=right nowrap>" & Format((bcw - acw), "###,###,##0") & "</td>"
                k6 = k6 + (bcw - acw)
                fs.WriteLine "            <td align=right nowrap>" & Format((ect), "###,###,##0") & "</td>"
                k7 = k7 + (ect)
                fs.WriteLine "            <td align=right nowrap>" & Format(ptd, "###,###,##0") & "</td>"
                ptt = ptt + ptd
                fs.WriteLine "            <td align=right nowrap>" & Format(ytd, "###,###,##0") & "</td>"
                yt = yt + ytd
                fs.WriteLine "            <td align=right nowrap>" & Format((acw - ptd), "###,###,##0") & "</td>"
                cyt = cyt + (acw - ptd)
                fs.WriteLine "            <td align=right nowrap>" & Format(((acw - ptd) - ytd), "###,###,##0") & "</td>"
                cct = cct + ((acw - ptd) - ytd)
                fs.WriteLine "        </tr>"
                    
                               
                sl.MoveNext
                Wend
                            fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                            fs.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL COST</td>"
                            fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(k1, "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((k2), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right> <font color=white>NA </td>"
                            If Round(Format(k5 / (k5 + k7)), 1) = 0 Then
                            fs.WriteLine "            <td align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Round(Format((k5 / (k5 + k7))) * 100, 1) & "</td>"
                            End If
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(k4, "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(k5, "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((k6), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((k7), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((ptt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((yt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((cyt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((cct), "###,###,##0") & "</td>"
                            fs.WriteLine "        </tr>"
                            
                            
                            
                            
                            fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                            fs.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL PROFIT</td>"
                            fs.WriteLine "            <td >&nbsp;</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 - k1), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - k2), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a5) - k5), "###,###,##0") & " </td>"
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((((a1 + a2 + a3) - (a5)) - k7), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt)) & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt)) & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt1 - cyt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt2 - cct), "###,###,##0") & "</td>"
                            fs.WriteLine "        </tr>"
  
   fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub
Public Sub rephtml1()
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
 
       With fso
'        strName = .BuildPath(C:\, rep1.html)
        Set fs1 = .CreateTextFile("C:\PCIS-Reports\" & filpat, True)
        
      End With
   
   fs1.WriteLine " <html> "
   fs1.WriteLine "<style>"
   fs1.WriteLine "    BODY INPUT"
   fs1.WriteLine "    {"
   fs1.WriteLine "      BACKGROUND-IMAGE: url(file://C:\FeatherTexture.bmp);"
    
   fs1.WriteLine "    }"
   fs1.WriteLine "    .TableFont"
   fs1.WriteLine "    {"
   fs1.WriteLine "        COLOR: Black;"
   fs1.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs1.WriteLine "        FONT-SIZE: 8pt;"
   fs1.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs1.WriteLine "        'FONT-WEIGHT: bolder;"
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
   fs1.WriteLine "<body scroll=auto>"
'''   fs1.WriteLine "    <center>"

'''    fs1.WriteLine "           <font size=2.5>PROJECT REVENUE & COST REPORT - LEVEL 2</font><BR>"
  fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=#acacac width=95%>"
 
                fs1.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=11><b>TL OFFSHORE SDN BHD</td>"
                fs1.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs1.WriteLine "        </tr>"
                fs1.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL</td>"
                fs1.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs1.WriteLine "        </tr>"
 
 
    
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs1.WriteLine "            <td align=center><font color=white >Baseline Budget</td>"
                fs1.WriteLine "            <td align=center><font color=white >Estimate @ Completiion</td>"
                fs1.WriteLine "            <td align=center colspan=5><font color=white>Cummulative To Date</td>"
                fs1.WriteLine "            <td align=center><font color=white >EstimateTo Complete</td>"
                fs1.WriteLine "            <td align=center><font color=white >ProjToDate LastYrEnd</td>"
                fs1.WriteLine "            <td align=center><font color=white >YearToDate LastMthEnd</td>"
                fs1.WriteLine "            <td align=center><font color=white >YrToDate CurrentYear</td>"
                fs1.WriteLine "            <td align=center><font color=white >ChangesIn CurrentMth</td>"
                fs1.WriteLine "        </tr>"
                
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>JobNo.Description</td>"
                fs1.WriteLine "            <td align=center><font color=white>jobNo.</td>"
                fs1.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td align=center><font color=white  >Revenue</td>"
                fs1.WriteLine "            <td nowrap align=center><font color=white align=center>%EAC</td>"
                fs1.WriteLine "            <td align=center><font color=white >BCWP</td>"
                fs1.WriteLine "            <td align=center><font color=white >ACWP</td>"
                fs1.WriteLine "            <td align=center><font color=white >CostVar</td>"
                fs1.WriteLine "            <td align=center><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "        </tr>"

   Dim l As Integer
    
   nn = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
   Dim a1 As Double
   Dim a2 As Double
   Dim a3 As Double
   Dim a4 As Double
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  SUM(rev_totamount)  from revenue where rev_projcode='" & nn(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select  SUM(rev_totamount)  from revenue where rev_projcode='" & nn(0) & "'  and rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select SUM(rev_totamount) from revenue where rev_projcode='" & nn(0) & "'  and rev_type='UBL' ", Cn, 3, 2
   If Not rv4.EOF Then
   a5 = rv4(0)
   End If
                
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "'", Cn, 3, 2
   While Not pt.EOF
        bpdl = bpdl + pt!ptd_lye_revn
        bydl = bydl + pt!ytd_lme_revn
        updl = updl + pt!ptd_lye_revn1
        uydl = uydl + pt!ytd_lme_revn1
   pt.MoveNext
   Wend
   
                
                
                
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=16><b><u>" & cbo_job.Text & "</td>"
                fs1.WriteLine "        </tr>"
    ''''''' one
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE-BUDGETED</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(((a1 + a2 + a3) - (a5)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right >NA</td>"
                fs1.WriteLine "        </tr>"
                
                
     ''''''' two
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE - VO(+)</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
  ''''''' three
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE- VO(-)</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
  ''''''' four
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE-BILLED</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a4, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(bpdl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(bydl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format((a4 - bpdl), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(((a4 - bpdl) - bydl), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
 ''''''' five
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE-UNBILLED</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a5 - a4, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(updl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(uydl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(((a5 - a4) - updl), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format((((a5 - a4) - updl) - uydl), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
                
  ''''total
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL REVENUE</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(a1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((a5), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - (a5)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                revt1 = (a4 - bpdl) + ((a5 - a4) - updl)
                fs1.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt1), "###,###,##0") & "</td>"
                revt2 = (((a4 - bpdl) - bydl) + (((a5 - a4) - updl) - uydl))
                fs1.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt2), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
                
                Dim k1 As Double
                Dim k2 As Double
                Dim k3 As Double
                Dim k4 As Double
                Dim k5 As Double
                Dim k6 As Double
                Dim k7 As Double
                k1 = 0: k2 = 0: k3 = 0: k4 = 0: k5 = 0: k6 = 0: k7 = 0
                Dim ptt As Double
                Dim yt As Double
                ptt = 0: yt = 0
                 Dim cyt As Double
                Dim cct As Double
                cyt = 0: cct = 0
                nf = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
                Dim sl As New ADODB.Recordset
                If sl.State Then sl.Close
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' order by jobno_code", Cn, 3, 2
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=16><b><u>" & cbo_job.Text & "</td>"
                fs1.WriteLine "        </tr>"
                While Not sl.EOF
                Dim bdg As Double
                Dim bcw As Double
                Dim acw As Double
                Dim ect As Double
               
                bdg = 0: bcw = 0: acw = 0: ect = 0
                          Dim ct As New ADODB.Recordset
                          If ct.State Then ct.Close
                          ct.Open "select SUM(bd_extdamt),SUM(bd_bcwpamt) from jobcharge j, cost c where j.job_code=c.bd_jobcharge and j.jobno='" & sl(0) & "' and j.job_proj_key='" & nf(0) & "' and c.bd_costtype='B'  ", Cn, 3, 2
                          If Not ct.EOF Then
                          bdg = ct(0)
                          bcw = ct(1)
                                        
                          End If
                          
                          Dim ct1 As New ADODB.Recordset
                          If ct1.State Then ct1.Close
                          ct1.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c where j.job_code=c.bd_jobcharge and j.jobno='" & sl(0) & "' and j.job_proj_key='" & nf(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not ct1.EOF Then
                          acw = ct1(0)
                          ect = ct1(1)
                                    
                          End If
                Dim ytd As Double
                Dim ptd As Double
                ytd = 0: ptd = 0
                Dim ctr As New ADODB.Recordset
                If ctr.State Then ctr.Close
                ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where jobno='" & sl(0) & "' and projkey='" & nf(0) & "'", Cn, 3, 2
                If Not ctr.EOF Then
                ytd = ctr(0)
                ptd = ctr(1)
                End If
                
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & sl(1) & "</td>"
                fs1.WriteLine "            <td align=center>" & sl(0) & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(bdg, "###,###,##0") & "</td>"
                k1 = k1 + bdg
                fs1.WriteLine "            <td align=right nowrap>" & Format((acw + ect), "###,###,##0") & "</td>"
                k2 = k2 + (acw + ect)
                fs1.WriteLine "            <td align=right> NA </td>"
                If Round(Format(acw / (acw + ect)), 3) = 0 Then
                fs1.WriteLine "            <td align=right>0</td>"
                Else
                fs1.WriteLine "            <td align=right nowrap>" & Round(Format((acw / (acw + ect))) * 100, 1) & "</td>"
                End If
                fs1.WriteLine "            <td align=right nowrap>" & Format(bcw, "###,###,##0") & "</td>"
                k4 = k4 + bcw
                fs1.WriteLine "            <td align=right nowrap>" & Format(acw, "###,###,##0") & "</td>"
                k5 = k5 + acw
                fs1.WriteLine "            <td align=right nowrap>" & Format((bcw - acw), "###,###,##0") & "</td>"
                k6 = k6 + (bcw - acw)
                fs1.WriteLine "            <td align=right nowrap>" & Format((ect), "###,###,##0") & "</td>"
                k7 = k7 + (ect)
                fs1.WriteLine "            <td align=right nowrap>" & Format(ptd, "###,###,##0") & "</td>"
                ptt = ptt + ptd
                fs1.WriteLine "            <td align=right nowrap>" & Format(ytd, "###,###,##0") & "</td>"
                yt = yt + ytd
                fs1.WriteLine "            <td align=right nowrap>" & Format((acw - ptd), "###,###,##0") & "</td>"
                cyt = cyt + (acw - ptd)
                fs1.WriteLine "            <td align=right nowrap>" & Format(((acw - ptd) - ytd), "###,###,##0") & "</td>"
                cct = cct + ((acw - ptd) - ytd)
                fs1.WriteLine "        </tr>"
                    
                               
                sl.MoveNext
                Wend
                            fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                            fs1.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL COST</td>"
                            fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(k1, "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((k2), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right> <font color=white>NA </td>"
                            If Round(Format(k5 / (k5 + k7)), 1) = 0 Then
                            fs1.WriteLine "            <td align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Round(Format((k5 / (k5 + k7))) * 100, 1) & "</td>"
                            End If
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(k4, "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(k5, "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((k6), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((k7), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((ptt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((yt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((cyt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((cct), "###,###,##0") & "</td>"
                            fs1.WriteLine "        </tr>"
                            
                            
                            
                            
                            fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                            fs1.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL PROFIT</td>"
                            fs1.WriteLine "            <td >&nbsp;</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 - k1), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - k2), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a5) - k5), "###,###,##0") & " </td>"
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((((a1 + a2 + a3) - (a5)) - k7), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt)) & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt)) & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt1 - cyt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt2 - cct), "###,###,##0") & "</td>"
                            fs1.WriteLine "        </tr>"
  
   fs1.WriteLine " </table>"
  
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"

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

Public Sub progcost()
 asd = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
 
Dim gtotal As Double
gtotal = 0
Dim ntotal As Double
ntotal = 0
Dim iddd As Double
iddd = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where  bd_projectkey='" & asd(0) & "' and  bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2


While Not fldata.EOF

     iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_jobcharge, "  -  ", Len(fldata!bd_jobcharge), vbTextCompare)
mmmm = Split(fldata!bd_resccode, "  -  ", Len(fldata!bd_resccode), vbTextCompare)

Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If

Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
fldata2.Open "select * from cost where   bd_projectkey='" & asd(0) & "' and  bd_jobcharge='" & fldata!bd_jobcharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

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


Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where bd_projectkey='" & asd(0) & "'  and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
While Not cd.EOF
 If cd!bd_chk = 1 Then
           
                    If cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate <= main.DTPcutdate1.Value Then
                    a = cd!bd_edate - cd!bd_sdate
                    c = 0
                    ElseIf cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate >= main.DTPcutdate1.Value Then
                    a = main.DTPcutdate1.Value - cd!bd_sdate
                    c = cd!bd_edate - main.DTPcutdate1.Value
                    
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
                            
                             If cd!bd_sdate <= main.DTPcutdate1.Value And cd!bd_edate <= main.DTPcutdate1.Value Then
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
End Sub


Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
