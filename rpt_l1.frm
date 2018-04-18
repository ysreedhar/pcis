VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_l1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "L1 - PRCR @ PROJECT KEY LEVEL - ALL PROJECTS"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   14940
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7935
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   14415
      ExtentX         =   25426
      ExtentY         =   13996
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PARTC"
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   255
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      Height          =   255
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PART B"
      Height          =   255
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PART A"
      Height          =   255
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "rpt_l1"
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

Private Sub cmd_save_Click()
filepath_l1.Show
SetParent filepath_l1.hwnd, rpt_l1.hwnd
End Sub

Private Sub Command1_Click()
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l1.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml
Unload frmBusy
End Sub

Private Sub command2_Click()
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l1.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml1
Unload frmBusy
End Sub

Private Sub Command3_Click()
filepath_l1b.Show
SetParent filepath_l1b.hwnd, rpt_l1.hwnd

End Sub

Private Sub Command4_Click()
filepathl1c.Show
SetParent filepathl1c.hwnd, rpt_l1.hwnd
End Sub

Private Sub Command5_Click()
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l1.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call repbp
Unload frmBusy
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "L1 - PRCR @ PROJECT KEY LEVEL - ALL PROJECTS"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
'    Me.Width = 11415
'    Me.Height = 9750

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
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
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
 
    fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
 
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap>(PART-A) </td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Revised Budget</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate @ Completiion</td>"
                fs.WriteLine "            <td align=center colspan=3 nowrap><font color=white>Cummu-Revenue TD</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Cummu-Cost TD</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Cummu-Profit TD</td>"
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>Proj Key Description</td>"
                fs.WriteLine "            <td ><font color=white>Proj Key</td>"
                
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap><font color=white>Billed</td>"
                fs.WriteLine "            <td><font color=white>UnBilled</td>"
                fs.WriteLine "            <td><font color=white>Total</td>"
                fs.WriteLine "            <td Nowrap><font color=white>%WC</td>"
                fs.WriteLine "            <td ><font color=white>BCWP</td>"
                fs.WriteLine "            <td ><font color=white>ACWP</td>"
                fs.WriteLine "            <td ><font color=white>CostVar</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td ><font color=white>GP%</td>"
                fs.WriteLine "        </tr>"
               
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
Dim q5 As Double
Dim q6 As Double
Dim q7 As Double
Dim q8 As Double
Dim q9 As Double
Dim q10 As Double
Dim q11 As Double
Dim q12 As Double
Dim q13 As Double
Dim bp1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bp1 = 0
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(proj_key) from projectmaster order by proj_key", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl
Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double
Dim z5 As Double
Dim z6 As Double
Dim z7 As Double
Dim z8 As Double
Dim z9 As Double
Dim z10 As Double
Dim z11 As Double
Dim z12 As Double
Dim z13 As Double
Dim bp As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bp = 0
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' order by proj_key", Cn, 3, 2
 While Not pl.EOF
                
 
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If

                
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
   rv.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
        Dim rv3 As New ADODB.Recordset
        If rv3.State Then rv3.Close
        rv3.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BLD' ", Cn, 3, 2
        While Not rv3.EOF
        a4 = a4 + rv3(0)
        rv3.MoveNext
        Wend
        
   Dim bgvo As New ADODB.Recordset
   If bgvo.State Then bgvo.Close
   bgvo.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT VO' ", Cn, 3, 2
   While Not bgvo.EOF
   bvo = bvo + bgvo(0)
   bgvo.MoveNext
   Wend
'------------------------------------------------------------
 aa1 = 0: aa2 = 0: aa3 = 0
Dim rav As New ADODB.Recordset
   If rav.State Then rav.Close
   rav.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rav.EOF Then
   aa1 = rav(0)
   End If
   
   Dim rav1 As New ADODB.Recordset
   If rav1.State Then rav1.Close
   rav1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rav1.EOF Then
   aa2 = rav1(0)
    End If
    Dim rav2 As New ADODB.Recordset
   If rav2.State Then rav2.Close
   rav2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rav2.EOF Then
   aa3 = rav2(0)
   End If
   
'   Dim rav3 As New ADODB.Recordset
'   If rav3.State Then rav3.Close
'   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
'   If Not rav3.EOF Then
'   aa4 = rav3(0)
'    End If


            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        If aa1 = Null Then aa1 = 0
        If aa2 = Null Then aa2 = 0
        If aa3 = Null Then aa3 = 0
        
         If IsNull(aa1) Then aa1 = 0
        If IsNull(aa2) Then aa2 = 0
        If IsNull(aa3) Then aa3 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (aa1 + aa2 + aa3)

Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
                              Dim bv As Double
                              bv = 0
                              bv = bvo + a1
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & pl(0) & "</td>"
 
                fs.WriteLine "            <td nowrap align=right>" & Format(bv, "###,###,##0") & "</td>"
                z1 = z1 + bv
                fs.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs.WriteLine "            <td nowrap align=right>" & Format((bv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bv - bdg)
                If a1 = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((bv - bdg) / bv) * 100), "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                z4 = z4 + (a1 + a2 + a3)
                fs.WriteLine "            <td nowrap align=right>" & Format((acw + ect), "###,###,##0") & "</td>"
                z5 = z5 + (acw + ect)
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - (acw + ect)), "###,###,##0") & "</td>"
                z6 = z6 + ((a1 + a2 + a3) - (acw + ect))
                If (a1 + a2 + a3) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - (acw + ect)) / (a1 + a2 + a3)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(a4, "###,###,##0") & "</td>"
                z7 = z7 + a4
                fs.WriteLine "            <td nowrap align=right>" & Format((a5 + av3) - a4, "###,###,##0") & "</td>"
                z8 = z8 + ((a5 + av3) - a4)
                fs.WriteLine "            <td nowrap align=right>" & Format(((a5 + av3)), "###,###,##0") & "</td>"
                z9 = z9 + ((a5 + av3))
                If (acw + ect) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((acw) / (acw + ect)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(bcw, "###,###,##0") & "</td>"
                z10 = z10 + bcw
                fs.WriteLine "            <td nowrap align=right>" & Format(acw, "###,###,##0") & "</td>"
                z11 = z11 + acw
                fs.WriteLine "            <td nowrap align=right>" & Format((bcw - acw), "###,###,##0") & "</td>"
                z12 = z12 + (bcw - acw)
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3)) - acw), "###,###,##0") & "</td>"
                z13 = z13 + (((a5 + av3)) - acw)
                If (((a5 + av3))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - acw) / ((a5 + av3))) * 100), "###,###,##0") & "</td>"
                End If
               
                fs.WriteLine "        </tr>"
                
                
                
pl.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Sub Total</td>"
                 
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                q1 = q1 + z1
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                q2 = q2 + z2
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z3, "###,###,##0") & "</td>"
                q3 = q3 + z3
                            If z1 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z3 / z1) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z4), "###,###,##0") & "</td>"
                q4 = q4 + z4
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z5, "###,###,##0") & "</td>"
                q5 = q5 + z5
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z6), "###,###,##0") & "</td>"
                q6 = q6 + z6
                          If z4 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z6 / z4) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z7, "###,###,##0") & "</td>"
                q7 = q7 + z7
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z8, "###,###,##0") & "</td>"
                q8 = q8 + z8
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z9), "###,###,##0") & "</td>"
                q9 = q9 + z9
                            If z5 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z9 / z5) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z10, "###,###,##0") & "</td>"
                q10 = q10 + z10
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z11, "###,###,##0") & "</td>"
                q11 = q11 + z11
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z12), "###,###,##0") & "</td>"
                q12 = q12 + z12
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z13), "###,###,##0") & "</td>"
                q13 = q13 + z13
                If z9 = 0 Then
                fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z13 / z9) * 100, 2), "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
assad:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                 
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q3 / q1) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q6 / q4) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q9 / q5) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
             
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                If q9 = 0 Then
                fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q13 / q9) * 100, 2), "###,###,##0") & "</td>"
                End If
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
   
    fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap>(PART-B) </td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs.WriteLine "        </tr>"
    
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate To Complete</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Proj Todate Last YrEnd </td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Yr TODate LastMonthEnd</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Current Yr ToDate</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Changes in Current Month</td>"
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>Proj Key Description</td>"
                fs.WriteLine "            <td ><font color=white>ProjKey</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                 fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "        </tr>"
                
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
Dim q5 As Double
Dim q6 As Double
Dim q7 As Double
Dim q8 As Double
Dim q9 As Double
Dim q10 As Double
Dim q11 As Double
Dim q12 As Double
Dim q13 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0
                
Dim jh As String

Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(proj_key) from projectmaster order by proj_key", Cn, 3, 2
While Not hh.EOF
Dim kl As String
kl = Mid(hh(0), 1, 3)
If jh = kl Then GoTo assad1
jh = kl

Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double
Dim z5 As Double
Dim z6 As Double
Dim z7 As Double
Dim z8 As Double
Dim z9 As Double
Dim z10 As Double
Dim z11 As Double
Dim z12 As Double
Dim z13 As Double

z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If
                
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
   rv.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BLD' ", Cn, 3, 2
    While Not rv3.EOF
    a4 = a4 + rv3(0)
    rv3.MoveNext
    Wend
        
'   Dim rv4 As New ADODB.Recordset
'   If rv4.State Then rv4.Close
'   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'   While Not rv4.EOF
'   a5 = a5 + rv4(0)
'   rv4.MoveNext
'   Wend
                    
                    
            '---------------------------------------------------------------
            
             aa1 = 0: aa2 = 0: aa3 = 0
Dim rav As New ADODB.Recordset
   If rav.State Then rav.Close
   rav.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rav.EOF Then
   aa1 = rav(0)
   End If
   
   Dim rav1 As New ADODB.Recordset
   If rav1.State Then rav1.Close
   rav1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rav1.EOF Then
   aa2 = rav1(0)
    End If
    Dim rav2 As New ADODB.Recordset
   If rav2.State Then rav2.Close
   rav2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rav2.EOF Then
   aa3 = rav2(0)
   End If
   
'   Dim rav3 As New ADODB.Recordset
'   If rav3.State Then rav3.Close
'   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
'   If Not rav3.EOF Then
'   aa4 = rav3(0)
'    End If


            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        If aa1 = "" Then aa1 = 0
        If aa2 = "" Then aa2 = 0
        If aa3 = "" Then aa3 = 0
        
         If IsNull(aa1) Then aa1 = 0
        If IsNull(aa2) Then aa2 = 0
        If IsNull(aa3) Then aa3 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (aa1 + aa2 + aa3)

Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
            
            
            '---------------------------------------------------------------
                    
                    Dim bpdl As Double
                    Dim bydl As Double
                    Dim updl As Double
                    Dim uydl As Double
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where  projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = ctr(0)
                        ptd = ctr(1)
                        End If
                                        
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & pl(0) & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - (a5 + av3)), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - (a5 + av3))
                fs.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - (a5 + av3)) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - (a5 + av3)) - ect)
                If ((a1 + a2 + a3) - (a5 + av3)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - (a5 + av3)) - ect) / ((a1 + a2 + a3) - (a5 + av3))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                fs.WriteLine "            <td nowrap align=right>" & Format(((a5 + av3) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + ((a5 + av3) - (bpdl + updl))
                fs.WriteLine "            <td nowrap align=right>" & Format((acw - ptd), "###,###,##0") & "</td>"
                z9 = z9 + (acw - ptd)
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3) - (bpdl + updl)) - (acw - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5 + av3)) - (bpdl + updl)) - (acw - ptd))
                If ((a5 + av3) - (bpdl + updl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)) / ((a5 + av3) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl))
                fs.WriteLine "            <td nowrap align=right>" & Format(((acw - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + ((acw - ptd) - ytd)
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd))
                If ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd)) / ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
                                             
                                        
pl.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Sub Total</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                q1 = q1 + z1
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                q2 = q2 + z2
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z3, "###,###,##0") & "</td>"
                q3 = q3 + z3
                            If z1 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z3 / z1) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z4), "###,###,##0") & "</td>"
                q4 = q4 + z4
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z5, "###,###,##0") & "</td>"
                q5 = q5 + z5
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z6), "###,###,##0") & "</td>"
                q6 = q6 + z6
                 
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z7, "###,###,##0") & "</td>"
                q7 = q7 + z7
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z8, "###,###,##0") & "</td>"
                q8 = q8 + z8
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z9), "###,###,##0") & "</td>"
                q9 = q9 + z9
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z10, "###,###,##0") & "</td>"
                          If z8 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z10 / z8) * 100, 2), "###,###,##0") & "</td>"
                            End If
                q10 = q10 + z10
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z11, "###,###,##0") & "</td>"
                q11 = q11 + z11
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z12), "###,###,##0") & "</td>"
                q12 = q12 + z12
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z13), "###,###,##0") & "</td>"
                q13 = q13 + z13
                            If z11 = 0 Then
                            fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z13 / z11) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs.WriteLine "        </tr>"
assad1:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q3 / q1) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
                 fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q10 / q8) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q13 / q11) * 100, 2), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
        fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub


Public Sub rephtmlfile()
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
   fs1.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
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
   fs1.WriteLine "    <center>"
 
    fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
 
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap>(PART-A) </td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs1.WriteLine "        </tr>"
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Revised Budget</td>"
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate @ Completiion</td>"
                fs1.WriteLine "            <td align=center colspan=3 nowrap><font color=white>Cummu-Revenue TD</td>"
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Cummu-Cost TD</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Cummu-Profit TD</td>"
                fs1.WriteLine "        </tr>"
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>Proj Key Description</td>"
                fs1.WriteLine "            <td ><font color=white>Proj Key</td>"
                
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                fs1.WriteLine "            <td nowrap><font color=white>Billed</td>"
                fs1.WriteLine "            <td><font color=white>UnBilled</td>"
                fs1.WriteLine "            <td><font color=white>Total</td>"
                fs1.WriteLine "            <td Nowrap><font color=white>%WC</td>"
                fs1.WriteLine "            <td ><font color=white>BCWP</td>"
                fs1.WriteLine "            <td ><font color=white>ACWP</td>"
                fs1.WriteLine "            <td ><font color=white>CostVar</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td ><font color=white>GP%</td>"
                fs1.WriteLine "        </tr>"
               
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
Dim q5 As Double
Dim q6 As Double
Dim q7 As Double
Dim q8 As Double
Dim q9 As Double
Dim q10 As Double
Dim q11 As Double
Dim q12 As Double
Dim q13 As Double
Dim bp1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bp1 = 0
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(proj_key) from projectmaster order by proj_key", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl
Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double
Dim z5 As Double
Dim z6 As Double
Dim z7 As Double
Dim z8 As Double
Dim z9 As Double
Dim z10 As Double
Dim z11 As Double
Dim z12 As Double
Dim z13 As Double
Dim bp As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bp = 0
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' order by proj_key", Cn, 3, 2
 While Not pl.EOF
                
 
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If

                
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
   rv.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
        Dim rv3 As New ADODB.Recordset
        If rv3.State Then rv3.Close
        rv3.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BLD' ", Cn, 3, 2
        While Not rv3.EOF
        a4 = a4 + rv3(0)
        rv3.MoveNext
        Wend
        
   Dim bgvo As New ADODB.Recordset
   If bgvo.State Then bgvo.Close
   bgvo.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT VO' ", Cn, 3, 2
   While Not bgvo.EOF
   bvo = bvo + bgvo(0)
   bgvo.MoveNext
   Wend
'------------------------------------------------------------
 aa1 = 0: aa2 = 0: aa3 = 0
Dim rav As New ADODB.Recordset
   If rav.State Then rav.Close
   rav.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rav.EOF Then
   aa1 = rav(0)
   End If
   
   Dim rav1 As New ADODB.Recordset
   If rav1.State Then rav1.Close
   rav1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rav1.EOF Then
   aa2 = rav1(0)
    End If
    Dim rav2 As New ADODB.Recordset
   If rav2.State Then rav2.Close
   rav2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rav2.EOF Then
   aa3 = rav2(0)
   End If
   
'   Dim rav3 As New ADODB.Recordset
'   If rav3.State Then rav3.Close
'   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
'   If Not rav3.EOF Then
'   aa4 = rav3(0)
'    End If


            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        If aa1 = "" Then aa1 = 0
        If aa2 = "" Then aa2 = 0
        If aa3 = "" Then aa3 = 0
        
         If IsNull(aa1) Then aa1 = 0
        If IsNull(aa2) Then aa2 = 0
        If IsNull(aa3) Then aa3 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (aa1 + aa2 + aa3)

Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
                              Dim bv As Double
                              bv = 0
                              bv = bvo + a1
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & pl(0) & "</td>"
 
                fs1.WriteLine "            <td nowrap align=right>" & Format(bv, "###,###,##0") & "</td>"
                z1 = z1 + bv
                fs1.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs1.WriteLine "            <td nowrap align=right>" & Format((bv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bv - bdg)
                If a1 = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((bv - bdg) / bv) * 100), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format((a1 + a2 + a3), "###,###,##0") & "</td>"
                z4 = z4 + (a1 + a2 + a3)
                fs1.WriteLine "            <td nowrap align=right>" & Format((acw + ect), "###,###,##0") & "</td>"
                z5 = z5 + (acw + ect)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - (acw + ect)), "###,###,##0") & "</td>"
                z6 = z6 + ((a1 + a2 + a3) - (acw + ect))
                If (a1 + a2 + a3) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - (acw + ect)) / (a1 + a2 + a3)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(a4, "###,###,##0") & "</td>"
                z7 = z7 + a4
                fs1.WriteLine "            <td nowrap align=right>" & Format((a5 + av3) - a4, "###,###,##0") & "</td>"
                z8 = z8 + ((a5 + av3) - a4)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a5 + av3)), "###,###,##0") & "</td>"
                z9 = z9 + ((a5 + av3))
                If (acw + ect) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((acw) / (acw + ect)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(bcw, "###,###,##0") & "</td>"
                z10 = z10 + bcw
                fs1.WriteLine "            <td nowrap align=right>" & Format(acw, "###,###,##0") & "</td>"
                z11 = z11 + acw
                fs1.WriteLine "            <td nowrap align=right>" & Format((bcw - acw), "###,###,##0") & "</td>"
                z12 = z12 + (bcw - acw)
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3)) - acw), "###,###,##0") & "</td>"
                z13 = z13 + (((a5 + av3)) - acw)
                If (((a5 + av3))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - acw) / ((a5 + av3))) * 100), "###,###,##0") & "</td>"
                End If
               
                fs1.WriteLine "        </tr>"
                
                
                
pl.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Sub Total</td>"
                 
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                q1 = q1 + z1
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                q2 = q2 + z2
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z3, "###,###,##0") & "</td>"
                q3 = q3 + z3
                            If z1 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z3 / z1) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z4), "###,###,##0") & "</td>"
                q4 = q4 + z4
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z5, "###,###,##0") & "</td>"
                q5 = q5 + z5
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z6), "###,###,##0") & "</td>"
                q6 = q6 + z6
                          If z4 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z6 / z4) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z7, "###,###,##0") & "</td>"
                q7 = q7 + z7
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z8, "###,###,##0") & "</td>"
                q8 = q8 + z8
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z9), "###,###,##0") & "</td>"
                q9 = q9 + z9
                            If z5 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z9 / z5) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z10, "###,###,##0") & "</td>"
                q10 = q10 + z10
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z11, "###,###,##0") & "</td>"
                q11 = q11 + z11
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z12), "###,###,##0") & "</td>"
                q12 = q12 + z12
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z13), "###,###,##0") & "</td>"
                q13 = q13 + z13
                If z9 = 0 Then
                fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z13 / z9) * 100, 2), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
assad:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                 
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q3 / q1) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q6 / q4) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q9 / q5) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
             
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                If q9 = 0 Then
                fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q13 / q9) * 100, 2), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
                
        fs1.WriteLine " </table>"
    
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"

End Sub



Public Sub rephtmlfile1()
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
   fs1.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
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
   fs1.WriteLine "    <center>"
   
    fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap>(PART-B) </td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs1.WriteLine "        </tr>"
    
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate To Complete</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Proj Todate Last YrEnd </td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Yr TODate LastMonthEnd</td>"
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Current Yr ToDate</td>"
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Changes in Current Month</td>"
                fs1.WriteLine "        </tr>"
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>Proj Key Description</td>"
                fs1.WriteLine "            <td ><font color=white>ProjKey</td>"
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                 fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                fs1.WriteLine "        </tr>"
                
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
Dim q5 As Double
Dim q6 As Double
Dim q7 As Double
Dim q8 As Double
Dim q9 As Double
Dim q10 As Double
Dim q11 As Double
Dim q12 As Double
Dim q13 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0
                
Dim jh As String

Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(proj_key) from projectmaster order by proj_key", Cn, 3, 2
While Not hh.EOF
Dim kl As String
kl = Mid(hh(0), 1, 3)
If jh = kl Then GoTo assad1
jh = kl

Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double
Dim z5 As Double
Dim z6 As Double
Dim z7 As Double
Dim z8 As Double
Dim z9 As Double
Dim z10 As Double
Dim z11 As Double
Dim z12 As Double
Dim z13 As Double

z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and c.bd_projectkey='" & pl(0) & "' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If
                
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
   rv.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select  rev_totamount  from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BLD' ", Cn, 3, 2
    While Not rv3.EOF
    a4 = a4 + rv3(0)
    rv3.MoveNext
    Wend
        
'   Dim rv4 As New ADODB.Recordset
'   If rv4.State Then rv4.Close
'   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'   While Not rv4.EOF
'   a5 = a5 + rv4(0)
'   rv4.MoveNext
'   Wend
                    
                    
            '---------------------------------------------------------------
            
             aa1 = 0: aa2 = 0: aa3 = 0
Dim rav As New ADODB.Recordset
   If rav.State Then rav.Close
   rav.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rav.EOF Then
   aa1 = rav(0)
   End If
   
   Dim rav1 As New ADODB.Recordset
   If rav1.State Then rav1.Close
   rav1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rav1.EOF Then
   aa2 = rav1(0)
    End If
    Dim rav2 As New ADODB.Recordset
   If rav2.State Then rav2.Close
   rav2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rav2.EOF Then
   aa3 = rav2(0)
   End If
   
'   Dim rav3 As New ADODB.Recordset
'   If rav3.State Then rav3.Close
'   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & nn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
'   If Not rav3.EOF Then
'   aa4 = rav3(0)
'    End If


            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        If aa1 = "" Then aa1 = 0
        If aa2 = "" Then aa2 = 0
        If aa3 = "" Then aa3 = 0
        
         If IsNull(aa1) Then aa1 = 0
        If IsNull(aa2) Then aa2 = 0
        If IsNull(aa3) Then aa3 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (aa1 + aa2 + aa3)

Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "' ", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
            
            
            '---------------------------------------------------------------
                    
                    Dim bpdl As Double
                    Dim bydl As Double
                    Dim updl As Double
                    Dim uydl As Double
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 where  projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = ctr(0)
                        ptd = ctr(1)
                        End If
                                        
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & pl(0) & "</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - (a5 + av3)), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - (a5 + av3))
                fs1.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - (a5 + av3)) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - (a5 + av3)) - ect)
                If ((a1 + a2 + a3) - (a5 + av3)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - (a5 + av3)) - ect) / ((a1 + a2 + a3) - (a5 + av3))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                fs1.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a5 + av3) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + ((a5 + av3) - (bpdl + updl))
                fs1.WriteLine "            <td nowrap align=right>" & Format((acw - ptd), "###,###,##0") & "</td>"
                z9 = z9 + (acw - ptd)
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3) - (bpdl + updl)) - (acw - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5 + av3)) - (bpdl + updl)) - (acw - ptd))
                If ((a5 + av3) - (bpdl + updl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)) / ((a5 + av3) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl))
                fs1.WriteLine "            <td nowrap align=right>" & Format(((acw - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + ((acw - ptd) - ytd)
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd))
                If ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) - ((acw - ptd) - ytd)) / ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
                                             
                                        
pl.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Sub Total</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                q1 = q1 + z1
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                q2 = q2 + z2
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z3, "###,###,##0") & "</td>"
                q3 = q3 + z3
                            If z1 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z3 / z1) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z4), "###,###,##0") & "</td>"
                q4 = q4 + z4
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z5, "###,###,##0") & "</td>"
                q5 = q5 + z5
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z6), "###,###,##0") & "</td>"
                q6 = q6 + z6
                 
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z7, "###,###,##0") & "</td>"
                q7 = q7 + z7
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z8, "###,###,##0") & "</td>"
                q8 = q8 + z8
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z9), "###,###,##0") & "</td>"
                q9 = q9 + z9
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z10, "###,###,##0") & "</td>"
                          If z8 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z10 / z8) * 100, 2), "###,###,##0") & "</td>"
                            End If
                q10 = q10 + z10
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z11, "###,###,##0") & "</td>"
                q11 = q11 + z11
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z12), "###,###,##0") & "</td>"
                q12 = q12 + z12
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z13), "###,###,##0") & "</td>"
                q13 = q13 + z13
                            If z11 = 0 Then
                            fs1.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                            Else
                            fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((z13 / z11) * 100, 2), "###,###,##0") & "</td>"
                            End If
                fs1.WriteLine "        </tr>"
assad1:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q3 / q1) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
                 fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q10 / q8) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((q13 / q11) * 100, 2), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
        fs1.WriteLine " </table>"
    
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"
End Sub

Public Sub repbp()
On Error Resume Next
Me.Top = 10
Me.Left = 10
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
 
    fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
 
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2>" & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=1 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap>(PART-C) </td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Revised Budget</td>"

                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2 nowrap><font color=white>Proj Key Description</td>"
                fs.WriteLine "            <td ><font color=white>Proj Key</td>"
                
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                
                fs.WriteLine "        </tr>"
               
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
 
q1 = 0: q2 = 0: q3 = 0: q4 = 0:
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(proj_key) from projectmaster where status <> 'InActive'  order by proj_key", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl
Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double

Dim bp As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0:
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' and status <> 'InActive'  order by proj_key", Cn, 3, 2
 While Not pl.EOF
                
      Dim blb As New ADODB.Recordset
      If blb.State Then blb.Close
      blb.Open "select SUM(revn),SUM(cost) from baseline where proj_key = '" & pl(0) & "' ", Cn, 3, 2
                        
            If Not blb.EOF Then
                                        
            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            fs.WriteLine "            <td colspan=2 nowrap>" & pl(1) & "</td>"
            fs.WriteLine "            <td nowrap>" & pl(0) & "</td>"
            If IsNull(blb(0)) Then
            fs.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs.WriteLine "            <td nowrap align=right>" & Format(blb(0), "###,###,##0") & "</td>"
            End If
            z1 = z1 + blb(0)
            If IsNull(blb(1)) Then
            fs.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs.WriteLine "            <td nowrap align=right>" & Format(blb(1), "###,###,##0") & "</td>"
            End If
                  
            z2 = z2 + blb(1)
            If IsNull(blb(0) - blb(1)) Then
            fs.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs.WriteLine "            <td nowrap align=right>" & Format((blb(0) - blb(1)), "###,###,##0") & "</td>"
            End If
            If IsNull(((blb(0) - blb(1)) / blb(0)) * 100) Then
              fs.WriteLine "            <td nowrap align=right>0</td>"
            Else
           fs.WriteLine "            <td nowrap align=right>" & Format((((blb(0) - blb(1)) / blb(0)) * 100), "###,###,##0") & "</td>"
            End If
            fs.WriteLine "        </tr>"
            
                
             End If
                
pl.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>Sub Total</td>"
                If IsNull(z1) Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                End If
                q1 = q1 + z1
                If IsNull(z2) Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                End If
                q2 = q2 + z2
                If IsNull(z1 - z2) Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1 - z2, "###,###,##0") & "</td>"
                End If
                If IsNull((z1 - z2) / z1) Then
                 fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((z1 - z2) / z1, "###,###,##0") & "</td>"
                End If

                fs.WriteLine "        </tr>"
assad:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>Total</td>"
                 
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1 - q2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q1 - q2) / q1)), "###,###,##0") & "</td>"
                
                ''Format(Round(((q1 - q2) / q1) * 100, 2), "###,###,##0")

                fs.WriteLine "        </tr>"
                
        fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"


End Sub


Public Sub repbp1()
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
   fs1.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
    
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
   fs1.WriteLine "    <center>"
 
    fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
 
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=2>" & GetCompanyName & "</td>"
                fs1.WriteLine "            <td align=center colspan=1 nowrap>PROJECT REVENUE & COST REPORT - L1 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap>(PART-C) </td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap>CuttOffDate:" & main.DTPcutdate1.Value & "</td>"

                fs1.WriteLine "        </tr>"
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3><font color=white>Reporting Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                
                fs1.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Revised Budget</td>"

                fs1.WriteLine "        </tr>"
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=2 nowrap><font color=white>Proj Key Description</td>"
                fs1.WriteLine "            <td ><font color=white>Proj Key</td>"
                
                fs1.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs1.WriteLine "            <td ><font color=white>Cost</td>"
                fs1.WriteLine "            <td ><font color=white>Profit</td>"
                fs1.WriteLine "            <td><font color=white>GP%</td>"
                
                fs1.WriteLine "        </tr>"
               
Dim q1 As Double
Dim q2 As Double
Dim q3 As Double
Dim q4 As Double
 
q1 = 0: q2 = 0: q3 = 0: q4 = 0:
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(proj_key) from projectmaster where status <> 'InActive'  order by proj_key", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl
Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double

Dim bp As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0:
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key like '" & kl & "%' and status <> 'InActive'  order by proj_key", Cn, 3, 2
 While Not pl.EOF
                
      Dim blb As New ADODB.Recordset
      If blb.State Then blb.Close
      blb.Open "select SUM(revn),SUM(cost) from baseline where proj_key = '" & pl(0) & "' ", Cn, 3, 2
                        
            If Not blb.EOF Then
                                        
            fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            fs1.WriteLine "            <td colspan=2 nowrap>" & pl(1) & "</td>"
            fs1.WriteLine "            <td nowrap>" & pl(0) & "</td>"
            If IsNull(blb(0)) Then
            fs1.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs1.WriteLine "            <td nowrap align=right>" & Format(blb(0), "###,###,##0") & "</td>"
            End If
            z1 = z1 + blb(0)
            If IsNull(blb(1)) Then
            fs1.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs1.WriteLine "            <td nowrap align=right>" & Format(blb(1), "###,###,##0") & "</td>"
            End If
                  
            z2 = z2 + blb(1)
            If IsNull(blb(0) - blb(1)) Then
            fs1.WriteLine "            <td nowrap align=right>0</td>"
            Else
            fs1.WriteLine "            <td nowrap align=right>" & Format((blb(0) - blb(1)), "###,###,##0") & "</td>"
            End If
            If IsNull(((blb(0) - blb(1)) / blb(0)) * 100) Then
              fs1.WriteLine "            <td nowrap align=right>0</td>"
            Else
           fs1.WriteLine "            <td nowrap align=right>" & Format((((blb(0) - blb(1)) / blb(0)) * 100), "###,###,##0") & "</td>"
            End If
            fs1.WriteLine "        </tr>"
            
                
             End If
                
pl.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>Sub Total</td>"
                If IsNull(z1) Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
                End If
                q1 = q1 + z1
                If IsNull(z2) Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                End If
                q2 = q2 + z2
                If IsNull(z1 - z2) Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1 - z2, "###,###,##0") & "</td>"
                End If
                If IsNull((z1 - z2) / z1) Then
                 fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((z1 - z2) / z1, "###,###,##0") & "</td>"
                End If

                fs1.WriteLine "        </tr>"
assad:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>Total</td>"
                 
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1 - q2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q1 - q2) / q1)), "###,###,##0") & "</td>"
                
                ''Format(Round(((q1 - q2) / q1) * 100, 2), "###,###,##0")

                fs1.WriteLine "        </tr>"
                
        fs1.WriteLine " </table>"
    
   

WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"



End Sub
