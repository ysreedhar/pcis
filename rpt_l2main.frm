VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_l2main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "L2 - PRCR @ JOBCHARGE LEVEL - BY PROJECT KEY"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14925
   LinkTopic       =   "Form2"
   ScaleHeight     =   10770
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7335
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   14415
      ExtentX         =   25426
      ExtentY         =   12938
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
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Project - Description"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Close"
         Height          =   255
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print"
         Height          =   255
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save To File"
         Height          =   255
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save To File"
         Height          =   255
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Firm Scope"
         Height          =   255
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_partb 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Variation Order"
         Height          =   255
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo_job 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "rpt_l2main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim Mode As Integer

Private Sub Check1_Click()

End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_partb_Click()
If cbo_job.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l2main.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtmlb(False)
Unload frmBusy
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

filepath.Show
SetParent filepath.hwnd, rpt_l2main.hwnd
End Sub


Private Sub cmd_show_Click()
If cbo_job.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If

frmBusy.Show
SetParent frmBusy.hwnd, rpt_l2main.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml
Unload frmBusy
 
End Sub
Private Sub Command1_Click()

filepathb.Show
SetParent filepathb.hwnd, rpt_l2main.hwnd

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
                fs.WriteLine "            <td colspan=11><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL (FIRM SCOPE)</td>"
                fs.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
 
 
    
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs.WriteLine "            <td align=center><font color=white >Revised Budget</td>"
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
   Dim bvo As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0: bvo = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
    
    Dim rbvo As New ADODB.Recordset
    If rbvo.State Then rbvo.Close
    rbvo.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
    If Not rbvo.EOF Then
    bvo = rbvo(0)
    End If
        
 Dim asam As Double
        Dim esam As Double
        
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & nn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
    
    If asam <> 0 Then
    a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
    End If
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "' and notes='MAIN'", Cn, 3, 2
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
              Dim bb As Double
              Dim ba As Double
              ba = 0
              bb = 0
              bb = ((a4 - bpdl) - bydl)
              ba = (a4 - bpdl)
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
                Dim aa As Double
                Dim ab As Double
                aa = 0
                ab = 0
                aa = (((a5 - a4) - updl) - uydl)
                ab = ((a5 - a4) - updl)
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
                revt1 = ba + ab
                fs.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt1), "###,###,##0") & "</td>"
                revt2 = bb + aa
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
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' and type ='MAIN' and status='Active' order by jobno_code", Cn, 3, 2
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
                ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where jobno='" & sl(0) & "' and projkey='" & nf(0) & "'  ", Cn, 3, 2
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
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt), "###,###,##0") & "</td>"
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
                fs1.WriteLine "            <td colspan=11><b>" & GetCompanyName & "</td>"
                fs1.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs1.WriteLine "        </tr>"
                fs1.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL (Part-A)</td>"
                fs1.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs1.WriteLine "        </tr>"
 
 
    
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs1.WriteLine "            <td align=center><font color=white >Revised Budget</td>"
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
                fs1.WriteLine "            <td align=center><font color=white  >BCWP</td>"
                fs1.WriteLine "            <td align=center><font color=white  >ACWP</td>"
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
   Dim bvo As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0: bvo = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
    
    Dim rbvo As New ADODB.Recordset
    If rbvo.State Then rbvo.Close
    rbvo.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
    If Not rbvo.EOF Then
    bvo = rbvo(0)
    End If
        
 Dim asam As Double
        Dim esam As Double
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & nn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
                
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "' and notes='MAIN'", Cn, 3, 2
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
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' and type ='MAIN' and status='Active' order by jobno_code", Cn, 3, 2
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
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt), "###,###,##0") & "</td>"
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




Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub


Public Sub rephtmlb(boolSaveAsExcel As Boolean)
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
   If boolSaveAsExcel = True Then
Set fs = fso.CreateTextFile("C:\PCIS-Reports\" & filpat, True)
Else
Set fs = fso.CreateTextFile(App.Path & "\rep.html")
End If
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
                fs.WriteLine "            <td colspan=11><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL (VARIATION ORDER)</td>"
                fs.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
 
 
    
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs.WriteLine "            <td align=center><font color=white >Revised Budget</td>"
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
   Dim bvo As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0: bvo = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency, rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and rev_id = " & jn(2), Cn, 3, 2
   'rvv1.Open "select rev_totamount, perc from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' and  r.rev_jobno='" & jn(0) & "'", Cn, 3, 2
   av2 = 0
   If Not rvv1.EOF Then
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   'rvv1.MoveNext
   'r = r + rvv1.RecordCount
   'Wend
   End If
   av3 = av3 + av2
   
   ' Calculation for total revn
'   Dim rsRevTotal As New ADODB.Recordset
'   If rsRevTotal.State Then rsRevTotal.Close
'   rsRevTotal.Open "select sum(rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)'", Cn, 3, 2
'   If Not rsRevTotal.EOF Then
'   RevTotal = CDbl(rsRevTotal(0))
'   End If
'   av3 = RevTotal
   jn.MoveNext
   Wend

   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
    Dim bgvo As New ADODB.Recordset
    If bgvo.State Then bgvo.Close
    bgvo.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
    If Not bgvo.EOF Then
    bvo = bgvo(0)
    End If
 Dim asam As Double
        Dim esam As Double
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='CO' and j.job_proj_key='" & nn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
                
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "' and notes='CO'", Cn, 3, 2
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
                fs.WriteLine "            <td align=right >NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right >NA</td>"
                fs.WriteLine "        </tr>"
                
                
     ''''''' two
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>REVENUE - VO(+)</td>"
                fs.WriteLine "            <td >&nbsp;</td>"
                fs.WriteLine "            <td align=right>" & Format(bvo, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(a2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(((a1 + a2 + a3) - (av3)), "###,###,##0") & "</td>"
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
                fs.WriteLine "            <td align=right nowrap>" & Format(av3 - a4, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right>NA</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(updl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(uydl, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format(((av3 - a4) - updl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap>" & Format((((av3 - a4) - updl) - uydl), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
                
  ''''total
                fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL REVENUE</td>"
                fs.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(a1 + bvo, "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((av3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right><font color=white>NA</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - (av3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                revt1 = (a4 - bpdl) + ((av3 - a4) - updl)
                
                fs.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt1), "###,###,##0") & "</td>"
                revt2 = (((a4 - bpdl) - bydl) + ((av3 - a4) - updl) - uydl)
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
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' and type ='CO' and status='Active' order by jobno_code", Cn, 3, 2
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
                ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where jobno='" & sl(0) & "' and projkey='" & nf(0) & "' ", Cn, 3, 2
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
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + bvo) - k1), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - k2), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((av3) - k5), "###,###,##0") & " </td>"
                            fs.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((((a1 + a2 + a3) - (av3)) - k7), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt1 - cyt), "###,###,##0") & "</td>"
                            fs.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt2 - cct), "###,###,##0") & "</td>"
                            fs.WriteLine "        </tr>"
  
   fs.WriteLine " </table>"
      If boolSaveAsExcel = True Then
  WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
  Else
   WebBrowser.Navigate App.Path & "\rep.html"
 End If
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
   End Sub

Public Sub rephtmlb1()
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
                fs1.WriteLine "            <td colspan=11><b>" & GetCompanyName & "</td>"
                fs1.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & cbo_job.Text & "</td>"
                fs1.WriteLine "        </tr>"
                fs1.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=11><b>PROJECT REVENUE & COST REPORT - L2 JOBCHARGE LEVEL (Part-B)</td>"
                fs1.WriteLine "           <td colspan=2><b>Cut-OffDate</td>"
                fs1.WriteLine "           <td colspan=3 align=center>" & main.DTPcutdate1.Value & "</td>"
                fs1.WriteLine "        </tr>"
 
 
    
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white>All Amounts in RM</td>"
                fs1.WriteLine "            <td align=center><font color=white >Revised Budget</td>"
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
                fs1.WriteLine "            <td align=center><font color=white  >BCWP</td>"
                fs1.WriteLine "            <td align=center><font color=white  >ACWP</td>"
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
   Dim bvo As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0: bvo = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
    Dim bgvo As New ADODB.Recordset
    If bgvo.State Then bgvo.Close
    bgvo.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
    If Not bgvo.EOF Then
    bvo = bgvo(0)
    End If
 Dim asam As Double
        Dim esam As Double
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='CO' and j.job_proj_key='" & nn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
                
   Dim bpdl As Double
   Dim bydl As Double
   Dim updl As Double
   Dim uydl As Double
   bpdl = 0: bydl = 0: updl = 0: uydl = 0
   Dim pt As New ADODB.Recordset
   If pt.State Then pt.Close
   pt.Open "select * from projecttransaction where pk_projkey='" & nn(0) & "' and notes='CO'", Cn, 3, 2
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
                fs1.WriteLine "            <td align=right >NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right >NA</td>"
                fs1.WriteLine "        </tr>"
                
                
     ''''''' two
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>REVENUE - VO(+)</td>"
                fs1.WriteLine "            <td >&nbsp;</td>"
                fs1.WriteLine "            <td align=right>" & Format(bvo, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(a2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(((a1 + a2 + a3) - (av3)), "###,###,##0") & "</td>"
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
                fs1.WriteLine "            <td align=right nowrap>" & Format(av3 - a4, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right>NA</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(updl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(uydl, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format(((av3 - a4) - updl), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap>" & Format((((av3 - a4) - updl) - uydl), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
                
  ''''total
                fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>TOTAL REVENUE</td>"
                fs1.WriteLine "            <td ><font color=white>&nbsp;</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(a1 + bvo, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((av3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - (av3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                revt1 = (a4 - bpdl) + ((av3 - a4) - updl)
                
                fs1.WriteLine "            <td align=right><font color=white nowrap>" & Format((revt1), "###,###,##0") & "</td>"
                revt2 = (((a4 - bpdl) - bydl) + ((av3 - a4) - updl) - uydl)
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
                sl.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & nf(0) & "' and type ='CO' and status='Active' order by jobno_code", Cn, 3, 2
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
                ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where jobno='" & sl(0) & "' and projkey='" & nf(0) & "' ", Cn, 3, 2
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
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + bvo) - k1), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((a1 + a2 + a3) - k2), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                           
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((av3) - k5), "###,###,##0") & " </td>"
                            fs1.WriteLine "            <td align=right><font color=white>NA</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((((a1 + a2 + a3) - (av3)) - k7), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bpdl + updl) - ptt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format(((bydl + uydl) - yt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt1 - cyt), "###,###,##0") & "</td>"
                            fs1.WriteLine "            <td align=right nowrap><font color=white>" & Format((revt2 - cct), "###,###,##0") & "</td>"
                            fs1.WriteLine "        </tr>"
  
   fs1.WriteLine " </table>"
    
   
   
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"
End Sub

