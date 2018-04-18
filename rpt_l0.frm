VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_l0 
   BackColor       =   &H00FFFFFF&
   Caption         =   "L0 - PRCR @ COMPANY LEVEL - ALL PROJECTS"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8415
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   14415
      ExtentX         =   25426
      ExtentY         =   14843
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
   Begin VB.CheckBox chkSplitRevenue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Split Revenue for PART B"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   360
      Width           =   3735
   End
   Begin VB.CheckBox chkLastYrEndData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Project to Date last year end for PART B"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   0
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FirmScope/CO"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   1560
      TabIndex        =   10
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmd_savea 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save To File"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   550
         Width           =   1215
      End
      Begin VB.CommandButton cmd_saveb 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save To File"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   550
         Width           =   1215
      End
      Begin VB.CommandButton cmd_main 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PART A"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_co 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PART B"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PART C"
      Height          =   255
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save To File"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cbo_year 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PART A"
      Height          =   255
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PART B"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      Height          =   255
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   255
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "rpt_l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmd_co_Click()
If cbo_year.Text = "" Then
MsgBox "Select Year"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l0.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
If chkLastYrEndData.Value Then
chkSplitRevenue.Value = False
chkSplitRevenue.Visible = False
Call rephtml1main
Else
If chkSplitRevenue.Visible = False Then chkSplitRevenue.Visible = True
If chkSplitRevenue.Value Then
Call repPartBDetail(False)
Else
Call repPartB(False)
End If
End If
Unload frmBusy
End Sub
Private Sub cmd_main_Click()
If cbo_year.Text = "" Then
MsgBox "Select Year"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l0.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtmlmain(False)
Unload frmBusy
End Sub
Private Sub cmd_print_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT::
End Sub
Private Sub cmd_save_Click()
filepathl0a.Show
SetParent filepathl0a.hwnd, rpt_l0.hwnd
End Sub
Private Sub cmd_savea_Click()
filepath_l0AMC.Show
SetParent filepath_l0AMC.hwnd, rpt_l0.hwnd
End Sub
Private Sub cmd_saveb_Click()
filepath_l0BMC.Show
SetParent filepath_l0BMC.hwnd, rpt_l0.hwnd
End Sub
Private Sub Command1_Click()
If cbo_year.Text = "" Then
MsgBox "Select Year"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l0.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml
Unload frmBusy
End Sub
Private Sub command2_Click()
If cbo_year.Text = "" Then
MsgBox "Select Year"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l0.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call rephtml1
Unload frmBusy
End Sub
Private Sub Command3_Click()
filepathl0b.Show
SetParent filepathl0b.hwnd, rpt_l0.hwnd
End Sub
Private Sub Command4_Click()
filepathl0c.Show
SetParent filepathl0c.hwnd, rpt_l0.hwnd
End Sub
Private Sub Command5_Click()
If cbo_year.Text = "" Then
MsgBox "Select Year"
Exit Sub
End If
frmBusy.Show
SetParent frmBusy.hwnd, rpt_l0.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call repbp
Unload frmBusy
End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "L0 - PRCR @ COMPANY LEVEL - ALL PROJECTS"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim h As Integer
h = 0
For h = 2000 To 2050
cbo_year.AddItem h
Next
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
                fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>" & asd
                fs.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap> (PART-A)</td>"
                fs.WriteLine "            <td align=center colspan=7 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
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
                fs.WriteLine "            <td ><font color=white>ProjKey</td>"
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
Dim bpg1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bpg1 = 0
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "'  order by bd_projectkey", Cn, 3, 2
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
Dim bpg As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bpg = 0
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key = '" & hh(0) & "' order by proj_key", Cn, 3, 2
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
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT VO' ", Cn, 3, 2
   While Not rv4.EOF
   bvo = bvo + rv4(0)
   rv4.MoveNext
   Wend
   
'''
'----------------------------------------------------
 aa1 = 0: aa2 = 0: aa3 = 0: aa4 = 0
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
   
   Dim rav3 As New ADODB.Recordset
   If rav3.State Then rav3.Close
   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not rav3.EOF Then
   aa4 = rav3(0)
    End If


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
        If IsNull(aa4) Then aa4 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (aa1 + aa2 + aa3)

Dim av3 As Double
Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency, rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
   Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_id='" & jn(2) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   jn.MoveNext
   Wend
   
   '-----------------------------------------------------------
                Dim bgv As Double
                bgv = 0
                bgv = CDbl(a1) + CDbl(bvo)
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & pl(0) & "</td>"
 
                fs.WriteLine "            <td nowrap align=right>" & Format(bgv, "###,###,##0") & "</td>"
                z1 = z1 + (bgv)
                fs.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs.WriteLine "            <td nowrap align=right>" & Format((bgv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bgv - bdg)
                If a1 = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((bgv) - bdg) / (bgv)) * 100), "###,###,##0") & "</td>"
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
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
assad:
hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q3 / q1) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q6 / q4) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                ''wrk
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q11 / q5) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q13 / q9) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
                Dim d1, d2, d3, d4, d5, d6, d7, dinc, dinc1 As Integer
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0: dinc = 0: dinc1 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!bdgt, "###,###,##0") & "</td>"
                d1 = d1 + oi!bdgt
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!bdgt * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!bdgt * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                If oi!exin = "Expenditure" Then
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                d3 = d3 + oi!eac
                fs.WriteLine "            <td nowrap align=right>" & Format(((oi!eac) * -1), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac * -1)
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                dinc = dinc + oi!eac
                fs.WriteLine "            <td nowrap align=right>" & Format(((oi!eac)), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac)
                End If
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!bcwp, "###,###,##0") & "</td>"
                d5 = d5 + oi!bcwp
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                d6 = d6 + oi!acwp
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!bcwp - oi!acwp), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!bcwp - oi!acwp)
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                  
                oi.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp + Nett O/H Recovery</td>"
'                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(((d1) * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(dinc, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc - d3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d5, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5 - d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6 * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
'                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + q3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d3) * -1) + q6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                'potential items
                Dim p1 As Double
                Dim p2 As Double
                Dim p3 As Double
                p1 = 0: p2 = 0: p3 = 0
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                
                 ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
'                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(pti(0), "###,###,##0") & "</td>"
                p1 = p1 + pti(0)
                fs.WriteLine "            <td nowrap align=right>" & Format(pti(1), "###,###,##0") & "</td>"
                p2 = p2 + pti(1)
                fs.WriteLine "            <td nowrap align=right>" & Format((pti(0) - pti(1)), "###,###,##0") & "</td>"
                p3 = p3 + (pti(0) - pti(1))
                If pti(0) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((pti(0) - pti(1)) / pti(0)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                 fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
'                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                
                 'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
'                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + q3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((d3) * -1) + q6) + p3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
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
                fs.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap> (PART-B)</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
                fs.WriteLine "        </tr>"
    
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
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
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
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
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend
             '-----------------------------------------------
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
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5 + av3))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5 + av3)))
                fs.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5 + av3))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5 + av3))) - ect)
                If ((a1 + a2 + a3) - ((a5 + av3))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5 + av3))) - ect) / ((a1 + a2 + a3) - ((a5 + av3)))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5 + av3)) - (bpdl + updl))
                fs.WriteLine "            <td nowrap align=right>" & Format((acw - ptd), "###,###,##0") & "</td>"
                z9 = z9 + (acw - ptd)
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5 + av3)) - (bpdl + updl)) - (acw - ptd))
                If (((a5 + av3)) - (bpdl + updl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)) / (((a5 + av3)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
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
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
assad1:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3), "###,###,##0") & "</td>"
                 fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q3 / q1) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q10 / q8) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q13 / q11) * 100), 2), "###,###,##0") & "</td>"
                
                fs.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + oi!ctd
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                d6 = d6 + oi!chng
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!chng * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((-d6), "###,###,##0") & "</td>"
                d7 = -d6
                 fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + q3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + q10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d7 + q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                
                
                'potential items
                
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                               
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                    
                fs.WriteLine "        </tr>"
                
                
                    'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + q3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + q10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d7 + q13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                
        fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub


Private Sub Form_Resize()
WebBrowser.Width = Me.Width - (Me.Width * 0.06)
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
Public Sub WriteReportLine(fs As Object, strWrite As String)
   fs.WriteLine strWrite
End Sub

Public Sub rephtmla()
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
                fs1.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap> (PART-A)</td>"
                fs1.WriteLine "            <td align=center colspan=7 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
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
                fs1.WriteLine "            <td ><font color=white>ProjKey</td>"
           
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
Dim bpg1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bpg1 = 0
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "'  order by bd_projectkey", Cn, 3, 2
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
Dim bpg As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bpg = 0
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key = '" & hh(0) & "' order by proj_key", Cn, 3, 2
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
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='BGT VO' ", Cn, 3, 2
   While Not rv4.EOF
   bvo = bvo + rv4(0)
   rv4.MoveNext
   Wend
'''
'----------------------------------------------------
 aa1 = 0: aa2 = 0: aa3 = 0: aa4 = 0
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
   
   Dim rav3 As New ADODB.Recordset
   If rav3.State Then rav3.Close
   rav3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not rav3.EOF Then
   aa4 = rav3(0)
    End If


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
        If IsNull(aa4) Then aa4 = 0
 
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
                Dim bgv As Double
                bgv = 0
                bgv = CDbl(a1) + CDbl(bvo)
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & pl(0) & "</td>"
 
                fs1.WriteLine "            <td nowrap align=right>" & Format(bgv, "###,###,##0") & "</td>"
                z1 = z1 + (bgv)
                fs1.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs1.WriteLine "            <td nowrap align=right>" & Format((bgv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bgv - bdg)
                If a1 = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((bgv) - bdg) / (bgv)) * 100), "###,###,##0") & "</td>"
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
               
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
 
assad:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q3 / q1) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q6 / q4) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                ''wrk
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q11 / q5) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q13 / q9) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
                
                Dim d1, d2, d3, d4, d5, d6, d7, dinc, dinc1 As Integer
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0: dinc = 0: dinc1 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                
                 fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
'                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!bdgt, "###,###,##0") & "</td>"
                d1 = d1 + oi!bdgt
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!bdgt * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!bdgt * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                If oi!exin = "Expenditure" Then
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                d3 = d3 + oi!eac
                fs1.WriteLine "            <td nowrap align=right>" & Format(((oi!eac) * -1), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac * -1)
                ElseIf oi!exin = "Income" Then
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                dinc = dinc + oi!eac
                fs1.WriteLine "            <td nowrap align=right>" & Format(((oi!eac)), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac)
                End If

                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!bcwp, "###,###,##0") & "</td>"
                d5 = d5 + oi!bcwp
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                d6 = d6 + oi!acwp
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!bcwp - oi!acwp), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!bcwp - oi!acwp)
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                  
                oi.MoveNext
                Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp + Nett O/H Recovery</td>"
'                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(((d1) * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(dinc, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc - d3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d5, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d6, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5 - d6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6 * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
'                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + q3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d3) * -1) + q6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                
                
                'potential items
                Dim p1 As Double
                Dim p2 As Double
                Dim p3 As Double
                p1 = 0: p2 = 0: p3 = 0
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                
                 ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
'                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(pti(0), "###,###,##0") & "</td>"
               p1 = p1 + pti(0)
                fs1.WriteLine "            <td nowrap align=right>" & Format(pti(1), "###,###,##0") & "</td>"
               p2 = p2 + pti(1)
                fs1.WriteLine "            <td nowrap align=right>" & Format((pti(0) - pti(1)), "###,###,##0") & "</td>"
                p3 = p3 + (pti(0) - pti(1))
                If pti(0) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((pti(0) - pti(1)) / pti(0)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                 fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
'                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                 'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
'                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + q3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((d3) * -1) + q6) + p3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                
        fs1.WriteLine " </table>"
    
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"

End Sub

Public Sub rephtmlb()
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
                fs1.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap> (PART-B)</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
                fs1.WriteLine "        </tr>"
    
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
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
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
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
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend
             '-----------------------------------------------
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
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency ='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
                    
                    
                    
                    
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
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5 + av3))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5 + av3)))
                fs1.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5 + av3))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5 + av3))) - ect)
                If ((a1 + a2 + a3) - ((a5 + av3))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5 + av3))) - ect) / ((a1 + a2 + a3) - ((a5 + av3)))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                fs1.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a5 + av3)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5 + av3)) - (bpdl + updl))
                fs1.WriteLine "            <td nowrap align=right>" & Format((acw - ptd), "###,###,##0") & "</td>"
                z9 = z9 + (acw - ptd)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5 + av3)) - (bpdl + updl)) - (acw - ptd))
                If (((a5 + av3)) - (bpdl + updl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5 + av3)) - (bpdl + updl)) - (acw - ptd)) / (((a5 + av3)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
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
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
assad1:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3), "###,###,##0") & "</td>"
                 fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q3 / q1) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q5), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q10 / q8) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((q13 / q11) * 100), 2), "###,###,##0") & "</td>"
                
                fs1.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + oi!ctd
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                d6 = d6 + oi!chng
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!chng * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((-d6), "###,###,##0") & "</td>"
                d7 = -d6
                 fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d2 + q3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d5 + q10), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d7 + q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                'potential items
                
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                               
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                    
                fs1.WriteLine "        </tr>"
                
                
                    'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d2 + q3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d5 + q10), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d7 + q13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
        fs1.WriteLine " </table>"
    
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"
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
                fs1.WriteLine "            <td align=center colspan=1 nowrap>PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
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
               
 
 Dim z1 As Double
Dim z2 As Double
 
 Dim jh As String
 Dim hh As New ADODB.Recordset
 Dim bpg As Double
z1 = 0: z2 = 0:
 If hh.State Then hh.Close
 
 hh.Open "select DISTINCT(c.bd_projectkey),p.proj_title from cost c,projectmaster p where c.bd_projectkey=p.proj_key and c.bd_year='" & cbo_year.Text & "' and p.status <> 'InActive'   order by bd_projectkey", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl

 

 
 Dim blb As New ADODB.Recordset
      If blb.State Then blb.Close
      blb.Open "select SUM(revn),SUM(cost) from baseline where proj_key = '" & hh(0) & "' ", Cn, 3, 2
                        
            If Not blb.EOF Then
                                        
            fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            fs1.WriteLine "            <td colspan=2 nowrap>" & hh(1) & "</td>"
            fs1.WriteLine "            <td nowrap>" & hh(0) & "</td>"
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

assad:
hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap><font color=white>Total</td>"
                 
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
              
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1 - z2, "###,###,##0") & "</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((z1 - z2) / z1)), "###,###,##0") & "</td>"
                
                ''Format(Round(((q1 - q2) / q1) * 100, 2), "###,###,##0")

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
                fs.WriteLine "            <td align=center colspan=1 nowrap>PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
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
               
 
 Dim z1 As Double
Dim z2 As Double
 
 Dim jh As String
 Dim hh As New ADODB.Recordset
 Dim bpg As Double
z1 = 0: z2 = 0:
 If hh.State Then hh.Close
 
 hh.Open "select DISTINCT(c.bd_projectkey),p.proj_title from cost c,projectmaster p where c.bd_projectkey=p.proj_key and c.bd_year='" & cbo_year.Text & "' and p.status <> 'InActive'   order by bd_projectkey", Cn, 3, 2
 While Not hh.EOF
 Dim kl As String
                kl = Mid(hh(0), 1, 3)
                If jh = kl Then GoTo assad
                jh = kl

 

 
 Dim blb As New ADODB.Recordset
      If blb.State Then blb.Close
      blb.Open "select SUM(revn),SUM(cost) from baseline where proj_key = '" & hh(0) & "' ", Cn, 3, 2
                        
            If Not blb.EOF Then
                                        
            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            fs.WriteLine "            <td colspan=2 nowrap>" & hh(1) & "</td>"
            fs.WriteLine "            <td nowrap>" & hh(0) & "</td>"
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

assad:
hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap><font color=white>Total</td>"
                 
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1, "###,###,##0") & "</td>"
              
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(z1 - z2, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round(((z1 - z2) / z1)), "###,###,##0") & "</td>"
                
                ''Format(Round(((q1 - q2) / q1) * 100, 2), "###,###,##0")

                fs.WriteLine "        </tr>"
                
        fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
 
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub


Public Sub rephtmlmain(boolSaveAsExcel As Boolean)
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
   'Set fs = fso.CreateTextFile(App.Path & "\rep.html")
If boolSaveAsExcel = True Then
Set fs = fso.CreateTextFile("C:\PCIS-Reports\" & filpat, True)
Else
Set fs = fso.CreateTextFile(App.Path & "\rep.html")
End If
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
                fs.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap> (PART-A)</td>"
                fs.WriteLine "            <td align=center colspan=7 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
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
                fs.WriteLine "            <td ><font color=white>ProjKey</td>"
           
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
Dim bpg1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bpg1 = 0
'-------------------------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
Dim cobpg1 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0: cobpg1 = 0


'----------------------------
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "'  order by bd_projectkey", Cn, 3, 2
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
Dim bpg As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bpg = 0
'--------------------------CO
Dim coz1 As Double
Dim coz2 As Double
Dim coz3 As Double
Dim coz4 As Double
Dim coz5 As Double
Dim coz6 As Double
Dim coz7 As Double
Dim coz8 As Double
Dim coz9 As Double
Dim coz10 As Double
Dim coz11 As Double
Dim coz12 As Double
Dim coz13 As Double
Dim cobpg As Double
coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0: cobpg = 0


'---------------------------
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key = '" & hh(0) & "' order by proj_key", Cn, 3, 2
 While Not pl.EOF
 
      
      
                ' main
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        
               ' co
                        Dim cobdg As Double
                        Dim cobcw As Double
                        Dim coacw As Double
                        Dim coect As Double
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close
abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = IIf(IsNull(abc(0)), 0, abc(0))
bcw = IIf(IsNull(abc(1)), 0, abc(1))
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = IIf(IsNull(ct1(0)), 0, ct1(0))
ect = IIf(IsNull(ct1(1)), 0, ct1(1))
End If
' co
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = IIf(IsNull(coabc(0)), 0, coabc(0))
cobcw = IIf(IsNull(coabc(1)), 0, coabc(1))
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = IIf(IsNull(coct1(0)), 0, coct1(0))
coect = IIf(IsNull(coct1(1)), 0, coct1(1))
End If
'co
                
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
    
    Dim coa1 As Double
    Dim coa2 As Double
    Dim coa3 As Double
    Dim coa4 As Double
    Dim coa5 As Double
    Dim cobvo As Double
    coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0: cobvo = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''----------------------main
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = IIf(IsNull(rv(0)), 0, rv(0))
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = IIf(IsNull(rv1(0)), 0, rv1(0))
   End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = IIf(IsNull(rv2(0)), 0, rv2(0))
   End If
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
   If Not rv3.EOF Then
   a4 = IIf(IsNull(rv3(0)), 0, rv3(0))
    End If
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not rv4.EOF Then
   bvo = IIf(IsNull(rv4(0)), 0, rv4(0))
   End If
'''
'----------------------------------------------------
''''''''----------------------CO
   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not corv.EOF Then
   coa1 = IIf(IsNull(corv(0)), 0, corv(0))
   End If
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not corv1.EOF Then
   coa2 = IIf(IsNull(corv1(0)), 0, corv1(0))
   End If
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not corv2.EOF Then
   coa3 = IIf(IsNull(corv2(0)), 0, corv2(0))
   End If
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
   If Not corv3.EOF Then
   coa4 = IIf(IsNull(corv3(0)), 0, corv3(0))
    End If
        
   Dim corv4 As New ADODB.Recordset
   If corv4.State Then corv4.Close
   corv4.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not corv4.EOF Then
   cobvo = IIf(IsNull(corv4(0)), 0, corv4(0))
   End If

'''
'----------------------------------------------------main

            Dim asam As Double
            Dim esam As Double
            
            asam = 0: esam = 0
            
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(IIf(IsNull(sam(0)), 0, sam(0)), "###,###,###,##0")
                          esam = Format(IIf(IsNull(sam(1)), 0, sam(1)), "###,###,###,##0")
 End If
                          'CO
                          Dim COsam As New ADODB.Recordset
                          If COsam.State Then COsam.Close
                          COsam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='CO' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          COasam = Format(IIf(IsNull(COsam(0)), 0, COsam(0)), "###,###,###,##0")
                          COesam = Format(IIf(IsNull(COsam(1)), 0, COsam(1)), "###,###,###,##0")
                                    
                          End If
        If a1 = Null Then a1 = 0
        If a2 = Null Then a2 = 0
        If a3 = Null Then a3 = 0
        
        If IsNull(a1) Then a1 = 0
        If IsNull(a2) Then a2 = 0
        If IsNull(a3) Then a3 = 0
        'check if the Jobcharge and Cost values are not null and there by asam <> 0
 If CDbl(asam) <> 0 Then
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
End If
 If CDbl(COasam) <> 0 Then
   coa5 = (CDbl(COasam) / (CDbl(COasam) + CDbl(COesam))) * (coa1 + coa2 + coa3)
End If
'-------------------------------------CO

'------------------------------
    Dim av3 As Double
    Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency,r.rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency ='" & jn(1) & "'", Cn, 3, 2
    rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_id='" & jn(2) & "'", Cn, 3, 2
   av2 = 0
   If Not rvv1.EOF Then
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
   
    
                Dim bgv As Double
                bgv = 0
                Dim cobgv As Double
                cobgv = 0
                bgv = CDbl(a1) + CDbl(bvo)
                cobgv = CDbl(coa1) + CDbl(cobvo)
                Dim StrFC As String
                StrFC = pl(0) & " - " & "MAIN"
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrFC & "</td>"
 
                fs.WriteLine "            <td nowrap align=right>" & Format(bgv, "###,###,##0") & "</td>"
                z1 = z1 + (bgv)
                fs.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs.WriteLine "            <td nowrap align=right>" & Format((bgv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bgv - bdg)
                If a1 = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((bgv) - bdg) / (bgv)) * 100), "###,###,##0") & "</td>"
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
                fs.WriteLine "            <td nowrap align=right>" & Format((a5) - a4, "###,###,##0") & "</td>"
                z8 = z8 + ((a5) - a4)
                fs.WriteLine "            <td nowrap align=right>" & Format(((a5)), "###,###,##0") & "</td>"
                z9 = z9 + ((a5))
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
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5)) - acw), "###,###,##0") & "</td>"
                z13 = z13 + (((a5)) - acw)
                If (((a5))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - acw) / ((a5))) * 100), "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
                
    '----------------------CO
    Dim StrCF As String
    StrCF = pl(0) & " - " & "CO"
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrCF & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(cobgv, "###,###,##0") & "</td>"
                coz1 = coz1 + (cobgv)
                fs.WriteLine "            <td nowrap align=right>" & Format(cobdg, "###,###,##0") & "</td>"
                coz2 = coz2 + cobdg
                fs.WriteLine "            <td nowrap align=right>" & Format((cobgv - cobdg), "###,###,##0") & "</td>"
                coz3 = coz3 + (cobgv - cobdg)
                If a1 = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((cobgv) - cobdg) / (cobgv)) * 100), "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((coa1 + coa2 + coa3), "###,###,##0") & "</td>"
                coz4 = coz4 + (coa1 + coa2 + coa3)
                fs.WriteLine "            <td nowrap align=right>" & Format((coacw + coect), "###,###,##0") & "</td>"
                coz5 = coz5 + (coacw + coect)
                fs.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - (coacw + coect)), "###,###,##0") & "</td>"
                coz6 = coz6 + ((coa1 + coa2 + coa3) - (coacw + coect))
                If (coa1 + coa2 + coa3) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - (coacw + coect)) / (coa1 + coa2 + coa3)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(coa4, "###,###,##0") & "</td>"
                coz7 = coz7 + coa4
                fs.WriteLine "            <td nowrap align=right>" & Format((av3) - coa4, "###,###,##0") & "</td>"
                coz8 = coz8 + ((av3) - coa4)
                fs.WriteLine "            <td nowrap align=right>" & Format(((av3)), "###,###,##0") & "</td>"
                coz9 = coz9 + ((av3))
                If (coacw + coect) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((coacw) / (coacw + coect)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(cobcw, "###,###,##0") & "</td>"
                coz10 = coz10 + cobcw
                fs.WriteLine "            <td nowrap align=right>" & Format(coacw, "###,###,##0") & "</td>"
                coz11 = coz11 + coacw
                fs.WriteLine "            <td nowrap align=right>" & Format((cobcw - coacw), "###,###,##0") & "</td>"
                coz12 = coz12 + (cobcw - coacw)
                fs.WriteLine "            <td nowrap align=right>" & Format((((av3)) - coacw), "###,###,##0") & "</td>"
                coz13 = coz13 + (((av3)) - coacw)
                'If (((coa5 + av3))) = 0 Then
                If (coa5 = 0) Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - coacw) / ((coa5))) * 100), "###,###,##0") & "</td>"
                End If
                
'--------------CO
                
pl.MoveNext
Wend
               
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                'CO
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
 
assad:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
               
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1 + coq1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2 + coq2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3 + coq3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5 + coq5, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6 + coq6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q6 + coq6) / (q4 + coq4)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7 + coq7, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8 + coq8, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                ''wrk
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q11 + coq11) / (q5 + coq5)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10 + coq10, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11 + coq11, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q9 + coq9)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "        </tr>"
                
                
                Dim d1, d2, d3, d4, d5, d6, d7, dinc, dinc1 As Integer
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0: dinc = 0: dinc1 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                'oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                'Code change the other transactions by income first and then expenditure and not as tranx order
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.exin desc", Cn, 3, 2
                While Not oi.EOF
                
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
'               fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!bdgt, "###,###,##0") & "</td>"
                d1 = d1 + oi!bdgt
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!bdgt * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!bdgt * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                ' Check the expense type for calculation
                If oi!exin = "Expenditure" Then
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                d3 = d3 + oi!eac
                fs.WriteLine "            <td nowrap align=right>" & Format(((oi!eac) * -1), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!bcwp, "###,###,##0") & "</td>"
                d5 = d5 + oi!bcwp
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                d6 = d6 + oi!acwp
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!bcwp - oi!acwp), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!bcwp - oi!acwp)
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                dinc = dinc + oi!eac
                fs.WriteLine "            <td nowrap align=right>" & Format(((oi!eac)), "###,###,##0") & "</td>"
                Incomed4 = Incomed4 + (oi!eac)
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!bcwp, "###,###,##0") & "</td>"
                Incomed5 = Incomed5 + oi!bcwp
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Incomed6 = Incomed6 + oi!acwp
                If (oi!bcwp - oi!acwp) > 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!bcwp - oi!acwp), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!acwp), "###,###,##0") & "</td>"
                End If
                'd7 = d7 + (oi!bcwp - oi!acwp)
                'fs.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                ' Computation for Cummulative Profit and GP%
                If (oi!bcwp - oi!acwp) <> 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                
                End If
                fs.WriteLine "        </tr>"
                End If
                oi.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp + Nett O/H Recovery</td>"
'                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(((d1) * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(dinc, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc - d3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Incomed4, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Incomed5, "###,###,##0") & "</td>"
                'IncomeTot = Incomed5 - Incomed6
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Incomed6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d5, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5 - d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6 * -1), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                
                fs.WriteLine "        </tr>"
                
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
'                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                ' Estimate @ Completion for Estimated profit calculated as Total Estimate @ Completion + Total Inc/Exp & O/H Recovery
                fs.WriteLine "            <td nowrap align=right>" & Format((((dinc - d3)) + (q6 + coq6)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                
                
                
                
                'potential items
                Dim p1 As Double
                Dim p2 As Double
                Dim p3 As Double
                p1 = 0: p2 = 0: p3 = 0
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                
                 ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
'                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(pti(0), "###,###,##0") & "</td>"
               p1 = p1 + pti(0)
                fs.WriteLine "            <td nowrap align=right>" & Format(pti(1), "###,###,##0") & "</td>"
               p2 = p2 + pti(1)
                fs.WriteLine "            <td nowrap align=right>" & Format((pti(0) - pti(1)), "###,###,##0") & "</td>"
                p3 = p3 + (pti(0) - pti(1))
                If pti(0) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((pti(0) - pti(1)) / pti(0)) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                 fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
'                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p2, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(p3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                
                 'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
'                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((d3) * -1) + (q6 + coq6)) + (p3 + cop3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
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
Public Sub repPartB(boolSaveAsExcel As Boolean)
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
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
                fs.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap> (PART-B)</td>"
                fs.WriteLine "            <td align=center colspan=8 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate To Complete</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Yr TODate LastMonthEnd</td>"
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
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
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
'-----------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0
'-------------
Dim jh As String
Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
'----------------------CO
Dim coz1 As Double
Dim coz2 As Double
Dim coz3 As Double
Dim coz4 As Double
Dim coz5 As Double
Dim coz6 As Double
Dim coz7 As Double
Dim coz8 As Double
Dim coz9 As Double
Dim coz10 As Double
Dim coz11 As Double
Dim coz12 As Double
Dim coz13 As Double

coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0

'---------------------
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        '-------------CO
                        Dim cobdg As Double
                        Dim cobcw As Double
                        Dim coacw As Double
                        Dim coect As Double
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
                        '---------------
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = IIf(IsNull(abc(0)), 0, abc(0))
bcw = IIf(IsNull(abc(1)), 0, abc(1))
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = IIf(IsNull(ct1(0)), 0, ct1(0))
ect = IIf(IsNull(ct1(1)), 0, ct1(1))
End If
'---------------CO
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = IIf(IsNull(coabc(0)), 0, coabc(0))
cobcw = IIf(IsNull(coabc(1)), 0, coabc(1))
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = IIf(IsNull(coct1(0)), 0, coct1(0))
coect = IIf(IsNull(coct1(1)), 0, coct1(1))
End If

'---------------
                
   Dim a1 As Double
   Dim a2 As Double
   Dim a3 As Double
   Dim a4 As Double
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''CO
   Dim coa1 As Double
   Dim coa2 As Double
   Dim coa3 As Double
   Dim coa4 As Double
   Dim coa5 As Double
   coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''
   '----------
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + IIf(IsNull(rv(0)), 0, rv(0))
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + IIf(IsNull(rv1(0)), 0, rv1(0))
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + IIf(IsNull(rv2(0)), 0, rv2(0))
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not rv3.EOF
    a4 = a4 + IIf(IsNull(rv3(0)), 0, rv3(0))
    rv3.MoveNext
    Wend
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend


'-------------CO

   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not corv.EOF
   coa1 = coa1 + IIf(IsNull(corv(0)), 0, corv(0))
   corv.MoveNext
   Wend
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not corv1.EOF
   coa2 = coa2 + IIf(IsNull(corv1(0)), 0, corv1(0))
   corv1.MoveNext
   Wend
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not corv2.EOF
   coa3 = coa3 + IIf(IsNull(corv2(0)), 0, corv2(0))
   corv2.MoveNext
   Wend
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not corv3.EOF
    coa4 = coa4 + IIf(IsNull(corv3(0)), 0, corv3(0))
    corv3.MoveNext
    Wend



'------------
            

            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(IIf(IsNull(sam(0)), 0, sam(0)), "###,###,###,##0")
                          esam = Format(IIf(IsNull(sam(1)), 0, sam(1)), "###,###,###,##0")
                                    
                          End If
                          'CO
                          Dim COsam As New ADODB.Recordset
                          If COsam.State Then COsam.Close
                          COsam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='CO' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          COasam = Format(IIf(IsNull(COsam(0)), 0, COsam(0)), "###,###,###,##0")
                          COesam = Format(IIf(IsNull(COsam(1)), 0, COsam(1)), "###,###,###,##0")
                                    
                          End If
        If a1 = Null Then a1 = 0
        If a2 = Null Then a2 = 0
        If a3 = Null Then a3 = 0
        
        If IsNull(a1) Then a1 = 0
        If IsNull(a2) Then a2 = 0
        If IsNull(a3) Then a3 = 0
        'check if the Jobcharge and Cost values are not null and there by asam <> 0
 If CDbl(asam) <> 0 Then
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
End If
 If CDbl(COasam) <> 0 Then
   coa5 = (CDbl(COasam) / (CDbl(COasam) + CDbl(COesam))) * (coa1 + coa2 + coa3)
End If
Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency, rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   rvv1.Open "select rev_totamount, perc from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "' and r.rev_jobno='" & jn(0) & "'", Cn, 3, 2
    'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_id='" & jn(2) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   'rvv1.MoveNext
   'Wend
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
                    
                    
                    
                    Dim bpdl As Double
                    Dim bydl As Double
                    Dim updl As Double
                    Dim uydl As Double
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='MAIN'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                    
                    'CO
                    
                    Dim cobpdl As Double
                    Dim cobydl As Double
                    Dim coupdl As Double
                    Dim couydl As Double
                    cobpdl = 0: cobydl = 0: coupdl = 0: couydl = 0
                    Dim copt As New ADODB.Recordset
                    If copt.State Then copt.Close
                    copt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='CO'", Cn, 3, 2
                    While Not copt.EOF
                        cobpdl = cobpdl + copt!ptd_lye_revn
                        cobydl = cobydl + copt!ytd_lme_revn
                        coupdl = coupdl + copt!ptd_lye_revn1
                        couydl = couydl + copt!ytd_lme_revn1
                    copt.MoveNext
                    Wend
                    '------------------
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='MAIN' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = IIf(IsNull(ctr(0)), 0, ctr(0))
                        ptd = IIf(IsNull(ctr(1)), 0, ctr(1))
                        End If
                                 'co
                        Dim coytd As Double
                        Dim coptd As Double
                        coytd = 0: coptd = 0
                        
                        Dim coctr As New ADODB.Recordset
                        If coctr.State Then coctr.Close
                        coctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='CO' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not coctr.EOF Then
                        coytd = IIf(IsNull(coctr(0)), 0, coctr(0))
                        coptd = IIf(IsNull(coctr(1)), 0, coctr(1))
                        End If
                        '-------------
                                        
                 Dim StrFC As String
                 StrFC = pl(0) & " - " & "MAIN"
                                        
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrFC & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5)))
                fs.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5))) - ect)
                If ((a1 + a2 + a3) - ((a5))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5))) - ect) / ((a1 + a2 + a3) - ((a5)))) * 100, "###,###,##0") & "</td>"
                End If
                'lme ytd main
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                z5 = z5 + ptd
                fs.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                z7 = z7 + ytd
                ' Profit and GP % added on 06/03/2007
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)), "###,###,##0") & "</td>"
                If ytd <> 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)) / (ytd) * 100, "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5)) - (bpdl + updl))
                fs.WriteLine "            <td nowrap align=right>" & Format(((acw) - ptd), "###,###,##0") & "</td>"
                z9 = z9 + ((acw) - ptd)
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - ((acw) - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5)) - (bpdl + updl)) - ((acw) - ptd))
                If (((a5)) - (bpdl + updl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - ((acw) - ptd)) / (((a5)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5)) - (bpdl + updl)) - (bydl + uydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((((acw) - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + (((acw) - ptd) - ytd)
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd))
                If (((a5) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>&nbsp;" & Format(((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)) / (((a5) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
                
                '-------------------CO
                 Dim StrCF As String
                 StrCF = pl(0) & " - " & "CO"
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrCF & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - ((coa5))), "###,###,##0") & "</td>"
                'fs.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3)), "###,###,##0") & "</td>"
                

                coz1 = coz1 + ((coa1 + coa2 + coa3) - (coa5))
                fs.WriteLine "            <td nowrap align=right>" & Format(coect, "###,###,##0") & "</td>"
                coz2 = coz2 + coect
                fs.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - coa5) - (coect)), "###,###,##0") & "</td>"
                coz3 = coz3 + (((coa1 + coa2 + coa3) - coa5) - (coect))
                If ((coa1 + coa2 + coa3) - coa5) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((coa1 + coa2 + coa3) - coa5) - (coect)) / ((coa1 + coa2 + coa3) - coa5)) * 100, "###,###,##0") & "</td>"
                End If
                
                'Last ME YTD
                
                fs.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                coz4 = coz4 + ((cobpdl + coupdl))
                'fs.WriteLine "            <td nowrap align=right>" & Format((coptd), "###,###,##0") & "</td>"
                
                coz5 = coz5 + ((coptd))
                'fs.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                
                coz6 = coz6 + ((cobydl + couydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((coytd), "###,###,##0") & "</td>"
                coz7 = coz7 + ((coytd))
                'Profit & GP% added on 06/03/2007
                
                fs.WriteLine "            <td nowrap align=right>" & Format(((cobydl + couydl) - (coytd)), "###,###,##0") & "</td>"
                If coytd <> 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format((((cobydl + couydl) - (coytd)) / (coytd)) * 100, "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((((coa5)) - (cobpdl + coupdl)), "###,###,##0") & "</td>"
                coz8 = coz8 + ((((coa5)) - (cobpdl + coupdl)))
                fs.WriteLine "            <td nowrap align=right>" & Format(((coacw) - coptd), "###,###,##0") & "</td>"
                coz9 = coz9 + (((coacw) - coptd))
                fs.WriteLine "            <td nowrap align=right>" & Format(((((coa5)) - (cobpdl + coupdl)) - ((coacw) - coptd)), "###,###,##0") & "</td>"
                coz10 = coz10 + (((((coa5)) - (cobpdl + coupdl)) - ((coacw) - coptd)))
                
                If (((coa5)) - (cobpdl + coupdl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((coa5)) - (cobpdl + coupdl)) - ((coacw) - coptd)) / (((av3)) - (cobpdl + coupdl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "            <td nowrap align=right>" & Format(((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)), "###,###,##0") & "</td>"
                coz11 = coz11 + (((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)))
                fs.WriteLine "            <td nowrap align=right>" & Format((((coacw) - coptd) - coytd), "###,###,##0") & "</td>"
                coz12 = coz12 + ((((coacw) - coptd) - coytd))
                fs.WriteLine "            <td nowrap align=right>" & Format((((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)), "###,###,##0") & "</td>"
                coz13 = coz13 + ((((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)))
                If ((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((((coa5)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)) / ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        </tr>"
pl.MoveNext
Wend
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                '--------------------CO
                
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
                
                
                '--------------------
assad1:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1 + coq1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2 + coq2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3 + coq3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                'fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6 + coq6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7 + coq7), "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6 + coq6) - (q7 + coq7), "###,###,##0") & "</td>"
                If (q7 + coq7) <> 0 Then
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(((((q6 + coq6) - (q7 + coq7)) / (q7 + coq7)) * 100), "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right><font color=white>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8 + coq8), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10 + coq10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q10 + coq10) / (q8 + coq8)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11 + coq11), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q11 + coq11)) * 100), 2), "###,###,##0") & "</td>"
                
                fs.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                'oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.exin desc", Cn, 3, 2
                While Not oi.EOF
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>" & Format(0 - oi!ytd, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>0</td>"
                If oi!exin = "Expenditure" Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + oi!ctd
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd * -1)
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>0</td>"
                dinc = dinc + oi!ctd
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd)
                End If
                fs.WriteLine "            <td nowrap align=right>0</td>"
                If oi!exin = "Expenditure" Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                d6 = d6 + (oi!chng * -1)
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                dCiCMinc = dCiCMinc + oi!chng
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                End If
                'd7 = d7 + (oi!chng * -1)
                fs.WriteLine "            <td nowrap align=right>0</td>"
                fs.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(0 - d3, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dCiCMinc + d6), "###,###,##0") & "</td>"
                d7 = -d6
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                'Last Year End Data
                'fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                'fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dCiCMinc + d6) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                'potential items
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                pti.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((q6 + coq6) - (q7 + coq7)) + (0 - d3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dCiCMinc + d6) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                ' Notes
                
                Dim rsNotes As New ADODB.Recordset
                If rsNotes.State Then rsNotes.Close
                rsNotes.Open "select * from tblL0Notes", Cn, 3, 2
                While Not rsNotes.EOF
                
                fs.WriteLine " <tr><td align='left' valign='top' colspan=22>"
                fs.WriteLine "   <table border=1 class=TableFont width=100% cellspacing=0 BORDERCOLOR=GRAY>"
                fs.WriteLine " <tr >"
                fs.WriteLine "     <td width='50%' valign='top' rowspan='6'>" & Replace(rsNotes(1), vbNewLine, "<br>") & "</td>"
                fs.WriteLine "    <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "   </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "     <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Prepared By:" & rsNotes(2) & "</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "       <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Reviewed By:" & rsNotes(3) & "</td>"
                fs.WriteLine " </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "         <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine " <td width='25%'  valign='top'>Approved By:" & rsNotes(4) & "</td>"
                rsNotes.MoveNext
                Wend
                fs.WriteLine " </tr>"
                fs.WriteLine " </table>"
                fs.WriteLine "      </td>  </tr>"
                If rsNotes.State Then rsNotes.Close
                fs.WriteLine " </table>"
   'WebBrowser.Navigate App.Path & "\rep.html"
If boolSaveAsExcel = True Then
    WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
Else
    WebBrowser.Navigate App.Path & "\rep.html"
End If
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
End Sub
Function CalculatePartBRevenuetoDateMain(acwp As Long, ECAC As Long, TRAC As Long)
decRevenueToDateMain = (acwp / ECAC) * TRAC
End Function
Function CalculatePartBRevenuetoDateCO(TRAC As Long, PercentageWorkCompleted As Long)
decRevenueToDateCO = TRAC * (PercentageWorkCompleted / 100)
End Function
Public Sub repPartBDetail(boolSaveAsExcel As Boolean)
On Error Resume Next
Me.Top = 10
Me.Left = 10
 Dim fso As New FileSystemObject
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
                fs.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=8 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap> (PART-B)</td>"
                fs.WriteLine "            <td align=center colspan=10 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate To Complete</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap><font color=white>Yr TODate LastMonthEnd</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap><font color=white>Current Yr ToDate</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap><font color=white>Changes in Current Month</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                 fs.WriteLine "            <td colspan=3 nowrap><font color=white>Proj Key Description</td>"
                fs.WriteLine "            <td ><font color=white>ProjKey</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Billed</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Un-Billed</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Billed</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Un-Billed</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Revn</td>"
                fs.WriteLine "            <td ><font color=white>Cost</td>"
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Billed</td>"
                fs.WriteLine "            <td nowrap ><font color=white>Un-Billed</td>"
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
'-----------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0
'-------------
Dim jh As String
Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
'----------------------CO
Dim coz1 As Double
Dim coz2 As Double
Dim coz3 As Double
Dim coz4 As Double
Dim coz5 As Double
Dim coz6 As Double
Dim coz7 As Double
Dim coz8 As Double
Dim coz9 As Double
Dim coz10 As Double
Dim coz11 As Double
Dim coz12 As Double
Dim coz13 As Double

coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0

'---------------------
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        '-------------CO
                        Dim cobdg As Double
                        Dim cobcw As Double
                        Dim coacw As Double
                        Dim coect As Double
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
                        '---------------
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = IIf(IsNull(abc(0)), 0, abc(0))
bcw = IIf(IsNull(abc(1)), 0, abc(1))
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = IIf(IsNull(ct1(0)), 0, ct1(0))
ect = IIf(IsNull(ct1(1)), 0, ct1(1))
End If
'---------------CO
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = IIf(IsNull(coabc(0)), 0, coabc(0))
cobcw = IIf(IsNull(coabc(1)), 0, coabc(1))
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = IIf(IsNull(coct1(0)), 0, coct1(0))
coect = IIf(IsNull(coct1(1)), 0, coct1(1))
End If

'---------------
                
   Dim a1 As Double
   Dim a2 As Double
   Dim a3 As Double
   Dim a4 As Double
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''CO
   Dim coa1 As Double
   Dim coa2 As Double
   Dim coa3 As Double
   Dim coa4 As Double
   Dim coa5 As Double
   coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''
   '----------
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + IIf(IsNull(rv(0)), 0, rv(0))
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + IIf(IsNull(rv1(0)), 0, rv1(0))
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + IIf(IsNull(rv2(0)), 0, rv2(0))
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not rv3.EOF
    a4 = a4 + IIf(IsNull(rv3(0)), 0, rv3(0))
    rv3.MoveNext
    Wend
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend


'-------------CO

   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not corv.EOF
   coa1 = coa1 + IIf(IsNull(corv(0)), 0, corv(0))
   corv.MoveNext
   Wend
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not corv1.EOF
   coa2 = coa2 + IIf(IsNull(corv1(0)), 0, corv1(0))
   corv1.MoveNext
   Wend
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not corv2.EOF
   coa3 = coa3 + IIf(IsNull(corv2(0)), 0, corv2(0))
   corv2.MoveNext
   Wend
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not corv3.EOF
    coa4 = coa4 + IIf(IsNull(corv3(0)), 0, corv3(0))
    corv3.MoveNext
    Wend



'------------
            

            Dim asam As Double
            Dim esam As Double
'            Dim aa1, aa2, aa3 As Double
           
            asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(IIf(IsNull(sam(0)), 0, sam(0)), "###,###,###,##0")
                          esam = Format(IIf(IsNull(sam(1)), 0, sam(1)), "###,###,###,##0")
                                    
                          End If
        If a1 = Null Then a1 = 0
        If a2 = Null Then a2 = 0
        If a3 = Null Then a3 = 0
        
        If IsNull(a1) Then a1 = 0
        If IsNull(a2) Then a2 = 0
        If IsNull(a3) Then a3 = 0
        'check if the Jobcharge and Cost values are not null and there by asam <> 0
 If CDbl(asam) <> 0 Then
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
End If
Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency,rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
    rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_id='" & jn(2) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
                    
                    
                    
                    Dim bpdl As Double
                    Dim bydl As Double
                    Dim updl As Double
                    Dim uydl As Double
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='MAIN'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                    
                    'CO
                    
                    Dim cobpdl As Double
                    Dim cobydl As Double
                    Dim coupdl As Double
                    Dim couydl As Double
                    cobpdl = 0: cobydl = 0: coupdl = 0: couydl = 0
                    Dim copt As New ADODB.Recordset
                    If copt.State Then copt.Close
                    copt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='CO'", Cn, 3, 2
                    While Not copt.EOF
                        cobpdl = cobpdl + copt!ptd_lye_revn
                        cobydl = cobydl + copt!ytd_lme_revn
                        coupdl = coupdl + copt!ptd_lye_revn1
                        couydl = couydl + copt!ytd_lme_revn1
                    copt.MoveNext
                    Wend
                    '------------------
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='MAIN' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = IIf(IsNull(ctr(0)), 0, ctr(0))
                        ptd = IIf(IsNull(ctr(1)), 0, ctr(1))
                        End If
                                 'co
                        Dim coytd As Double
                        Dim coptd As Double
                        coytd = 0: coptd = 0
                        
                        Dim coctr As New ADODB.Recordset
                        If coctr.State Then coctr.Close
                        coctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='CO' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not coctr.EOF Then
                        coytd = IIf(IsNull(coctr(0)), 0, coctr(0))
                        coptd = IIf(IsNull(coctr(1)), 0, coctr(1))
                        End If
                        '-------------
                                        
                 Dim StrFC As String
                 StrFC = pl(0) & " - " & "MAIN"
                                        
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrFC & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5)))
                fs.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5))) - ect)
                If ((a1 + a2 + a3) - ((a5))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5))) - ect) / ((a1 + a2 + a3) - ((a5)))) * 100, "###,###,##0") & "</td>"
                End If
                'lme ytd main
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((uydl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                z5 = z5 + ptd
                fs.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                z7 = z7 + ytd
                ' Profit and GP % added on 06/03/2007
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)), "###,###,##0") & "</td>"
                If ytd <> 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)) / (ytd) * 100, "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(((bpdl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((updl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5)) - (bpdl + updl))
                fs.WriteLine "            <td nowrap align=right>" & Format(((acw) - ptd), "###,###,##0") & "</td>"
                z9 = z9 + ((acw) - ptd)
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - ((acw) - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5)) - (bpdl + updl)) - ((acw) - ptd))
                If (((a5)) - (bpdl + updl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - ((acw) - ptd)) / (((a5)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (bydl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (uydl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5)) - (bpdl + updl)) - (bydl + uydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((((acw) - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + (((acw) - ptd) - ytd)
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd))
                If (((a5) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>&nbsp;" & Format(((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)) / (((a5) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
                
                '-------------------CO
                 Dim StrCF As String
                 StrCF = pl(0) & " - " & "CO"
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrCF & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - ((av3))), "###,###,##0") & "</td>"
                coz1 = coz1 + ((coa1 + coa2 + coa3) - ((av3)))
                fs.WriteLine "            <td nowrap align=right>" & Format(coect, "###,###,##0") & "</td>"
                coz2 = coz2 + coect
                fs.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - ((av3))) - coect), "###,###,##0") & "</td>"
                coz3 = coz3 + (((coa1 + coa2 + coa3) - ((av3))) - coect)
                If ((coa1 + coa2 + coa3) - ((av3))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((coa1 + coa2 + coa3) - ((av3))) - coect) / ((coa1 + coa2 + coa3) - ((av3)))) * 100, "###,###,##0") & "</td>"
                End If
                
                'Last ME YTD
                fs.WriteLine "            <td nowrap align=right>" & Format((cobydl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((couydl), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                coz4 = coz4 + ((cobpdl + coupdl))
                'fs.WriteLine "            <td nowrap align=right>" & Format((coptd), "###,###,##0") & "</td>"
                
                coz5 = coz5 + ((coptd))
                'fs.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                coz6 = coz6 + ((cobydl + couydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((coytd), "###,###,##0") & "</td>"
                coz7 = coz7 + ((coytd))
                'Profit & GP% added on 06/03/2007
                fs.WriteLine "            <td nowrap align=right>" & Format(((cobydl + couydl) - (coytd)), "###,###,##0") & "</td>"
                If coytd <> 0 Then
                fs.WriteLine "            <td nowrap align=right>" & Format((((cobydl + couydl) - (coytd)) / (coytd)) * 100, "###,###,##0") & "</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((((av3)) - (cobpdl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((av3)) - (coupdl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((av3)) - (cobpdl + coupdl)), "###,###,##0") & "</td>"
                coz8 = coz8 + ((((av3)) - (cobpdl + coupdl)))
                fs.WriteLine "            <td nowrap align=right>" & Format(((coacw) - coptd), "###,###,##0") & "</td>"
                coz9 = coz9 + (((coacw) - coptd))
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)), "###,###,##0") & "</td>"
                coz10 = coz10 + (((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)))
                
                If (((av3)) - (cobpdl + coupdl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)) / (((av3)) - (cobpdl + coupdl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - (cobydl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - (couydl)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)), "###,###,##0") & "</td>"
                coz11 = coz11 + (((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)))
                fs.WriteLine "            <td nowrap align=right>" & Format((((coacw) - coptd) - coytd), "###,###,##0") & "</td>"
                coz12 = coz12 + ((((coacw) - coptd) - coytd))
                fs.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)), "###,###,##0") & "</td>"
                coz13 = coz13 + ((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)))
                If ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)) / ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        </tr>"
pl.MoveNext
Wend
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                '--------------------CO
                
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
                
                
                '--------------------
assad1:

hh.MoveNext
Wend
                                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1 + coq1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2 + coq2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3 + coq3), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q5 + coq5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6 + coq6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7 + coq7), "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6 + coq6) - (q7 + coq7), "###,###,##0") & "</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>0</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>0</td>"
                If (q7 + coq7) <> 0 Then
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(((((q6 + coq6) - (q7 + coq7)) / (q7 + coq7)) * 100), "###,###,##0") & "</td>"
                Else
                fs.WriteLine "<td nowrap align=right><font color=white>0</td>"
                End If
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8 + coq8), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10 + coq10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q10 + coq10) / (q8 + coq8)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>0</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>0</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11 + coq11), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q11 + coq11)) * 100), 2), "###,###,##0") & "</td>"
                
                fs.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                'oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.exin desc", Cn, 3, 2
                While Not oi.EOF
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                 
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "<td>NA</td>"
                If oi!exin = "Expenditure" Then
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + (oi!ctd * -1)
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                dinc = dinc + oi!ctd
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                d5 = d5 + (oi!ctd * -1)
                End If
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                If oi!exin = "Expenditure" Then
                fs.WriteLine "<td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                d6 = d6 + (oi!chng * -1)
                ElseIf oi!exin = "Income" Then
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                fs.WriteLine "<td>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                dCiCMinc = dCiCMinc + oi!chng
                
                End If
                'd7 = d7 + (oi!chng * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc + d4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                 fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "<td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dCiCMinc), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((dCiCMinc + d6), "###,###,##0") & "</td>"
                d7 = -d6
                 fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dinc + d4) + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dCiCMinc + d6) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                'potential items
                
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                               
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                pti.MoveNext
                Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                    'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dinc + d4) + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((dCiCMinc + d6) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                ' Notes Column at the bottom
                
                Dim rsNotes As New ADODB.Recordset
                If rsNotes.State Then rsNotes.Close
                rsNotes.Open "select * from tblL0Notes", Cn, 3, 2
                While Not rsNotes.EOF
                fs.WriteLine " <tr><td align='left' valign='top' colspan=26>"
                fs.WriteLine "   <table border=1 class=TableFont width=100% cellspacing=0 BORDERCOLOR=GRAY>"
                fs.WriteLine " <tr >"
                fs.WriteLine "     <td width='50%' valign='top' rowspan='6'>" & Replace(rsNotes(1), vbNewLine, "<br>") & "</td>"
                fs.WriteLine "    <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "   </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "     <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Prepared By:" & rsNotes(2) & "</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "       <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Reviewed By:" & rsNotes(3) & "</td>"
                fs.WriteLine " </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "         <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine " <td width='25%'  valign='top'>Approved By:" & rsNotes(4) & "</td>"
                rsNotes.MoveNext
                Wend
                fs.WriteLine " </tr>"
                fs.WriteLine " </table>"
                fs.WriteLine "      </td>  </tr>"
                If rsNotes.State Then rsNotes.Close
                
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


Public Sub rephtml1main()
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
                fs.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap> (PART-B)</td>"
                fs.WriteLine "            <td align=center colspan=8 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
                fs.WriteLine "        </tr>"
    
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Estimate To Complete</td>"
                fs.WriteLine "            <td align=center colspan=2 nowrap><font color=white>Proj Todate Last YrEnd </td>"
                fs.WriteLine "            <td align=center colspan=4 nowrap><font color=white>Yr TODate LastMonthEnd</td>"
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
                fs.WriteLine "            <td ><font color=white>Profit</td>"
                fs.WriteLine "            <td><font color=white>GP%</td>"
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
'-----------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0
'-------------
                
Dim jh As String

Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
'----------------------CO
Dim coz1 As Double
Dim coz2 As Double

Dim coz3 As Double
Dim coz4 As Double

Dim coz5 As Double
Dim coz6 As Double

Dim coz7 As Double
Dim coz8 As Double

Dim coz9 As Double
Dim coz10 As Double

Dim coz11 As Double
Dim coz12 As Double

Dim coz13 As Double

coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0

'---------------------
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        
                        Dim acw As Double
                        Dim ect As Double
                        
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        '-------------CO
                        Dim cobdg As Double
                        Dim cobcw As Double
                        
                        Dim coacw As Double
                        Dim coect As Double
                        
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
                        '---------------
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If
'---------------CO
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = coabc(0)
cobcw = coabc(1)
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = coct1(0)
coect = coct1(1)
End If

'---------------
                
   Dim a1 As Double
   Dim a2 As Double
   
   Dim a3 As Double
   Dim a4 As Double
   
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''CO
   Dim coa1 As Double
   Dim coa2 As Double
   
   Dim coa3 As Double
   Dim coa4 As Double
   
   Dim coa5 As Double
   coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''
 
   
   '----------
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
   While Not rv3.EOF
   a4 = a4 + rv3(0)
   rv3.MoveNext
   Wend
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend


'-------------CO

   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not corv.EOF
   coa1 = coa1 + corv(0)
   corv.MoveNext
   Wend
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not corv1.EOF
   coa2 = coa2 + corv1(0)
   corv1.MoveNext
   Wend
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not corv2.EOF
   coa3 = coa3 + corv2(0)
   corv2.MoveNext
   Wend
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not corv3.EOF
    coa4 = coa4 + corv3(0)
    corv3.MoveNext
    Wend



'------------
            

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
        If a1 = Null Then a1 = 0
        If a2 = Null Then a2 = 0
        If a3 = Null Then a3 = 0
        
        If IsNull(a1) Then a1 = 0
        If IsNull(a2) Then a2 = 0
        If IsNull(a3) Then a3 = 0
        'check if the Jobcharge and Cost values are not null and there by asam <> 0
 If CDbl(asam) <> 0 Then
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
End If
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
                    Dim bpdl As Double
                    Dim bydl As Double
                    
                    Dim updl As Double
                    Dim uydl As Double
                    
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='MAIN'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                    
                    'CO
                    
                    Dim cobpdl As Double
                    Dim cobydl As Double
                    
                    Dim coupdl As Double
                    Dim couydl As Double
                    cobpdl = 0: cobydl = 0: coupdl = 0: couydl = 0
                    Dim copt As New ADODB.Recordset
                    If copt.State Then copt.Close
                    copt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='CO'", Cn, 3, 2
                    While Not copt.EOF
                        cobpdl = cobpdl + copt!ptd_lye_revn
                        cobydl = cobydl + copt!ytd_lme_revn
                        
                        coupdl = coupdl + copt!ptd_lye_revn1
                        couydl = couydl + copt!ytd_lme_revn1
                    copt.MoveNext
                    Wend
                    '------------------
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='MAIN' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = ctr(0)
                        ptd = ctr(1)
                        End If
                                        
                                 'co
                        Dim coytd As Double
                        Dim coptd As Double
                        coytd = 0: coptd = 0
                        
                        Dim coctr As New ADODB.Recordset
                        If coctr.State Then coctr.Close
                        coctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='CO' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not coctr.EOF Then
                        coytd = coctr(0)
                        coptd = coctr(1)
                        End If
                        '-------------
                                        
                 Dim StrFC As String
                 StrFC = pl(0) & " - " & "MAIN"
                                        
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                
                fs.WriteLine "            <td nowrap>" & StrFC & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5)))
                
                fs.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5))) - ect)
                If ((a1 + a2 + a3) - ((a5))) = 0 Then
                
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5))) - ect) / ((a1 + a2 + a3) - ((a5)))) * 100, "###,###,##0") & "</td>"
                End If
                
                
                fs.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                
                
                fs.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                
                ' Profit and GP % added on 06/03/2007
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((bydl + uydl) - (ytd)) / (ytd) * 100, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right>" & Format((((a5)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5)) - (bpdl + updl))
                fs.WriteLine "            <td nowrap align=right>" & Format(((acw) - ptd), "###,###,##0") & "</td>"
                z9 = z9 + ((acw) - ptd)
                
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - ((acw) - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5)) - (bpdl + updl)) - ((acw) - ptd))
                
                If (((a5)) - (bpdl + updl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - ((acw) - ptd)) / (((a5)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5)) - (bpdl + updl)) - (bydl + uydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((((acw) - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + (((acw) - ptd) - ytd)
                
                fs.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd))
                If ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)) / ((((a5)) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "        </tr>"
                
                
                '-------------------CO
                 Dim StrCF As String
                 StrCF = pl(0) & " - " & "CO"
                   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs.WriteLine "            <td nowrap>" & StrCF & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - ((av3))), "###,###,##0") & "</td>"
                coz1 = coz1 + ((coa1 + coa2 + coa3) - ((av3)))
                fs.WriteLine "            <td nowrap align=right>" & Format(coect, "###,###,##0") & "</td>"
                coz2 = coz2 + coect
                fs.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - ((av3))) - coect), "###,###,##0") & "</td>"
                coz3 = coz3 + (((coa1 + coa2 + coa3) - ((av3))) - coect)
                If ((coa1 + coa2 + coa3) - ((av3))) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((coa1 + coa2 + coa3) - ((av3))) - coect) / ((coa1 + coa2 + coa3) - ((av3)))) * 100, "###,###,##0") & "</td>"
                End If
                fs.WriteLine "            <td nowrap align=right>" & Format((cobpdl + coupdl), "###,###,##0") & "</td>"
                coz4 = coz4 + ((cobpdl + coupdl))
                fs.WriteLine "            <td nowrap align=right>" & Format((coptd), "###,###,##0") & "</td>"
                coz5 = coz5 + ((coptd))
                fs.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                coz6 = coz6 + ((cobydl + couydl))
                fs.WriteLine "            <td nowrap align=right>" & Format((coytd), "###,###,##0") & "</td>"
                coz7 = coz7 + ((coytd))
                'Profit & GP% added on 06/03/2007
                fs.WriteLine "            <td nowrap align=right>" & Format(((cobydl + couydl) - (coytd)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((((cobydl + couydl) - (coytd)) / (coytd)) * 100, "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right>" & Format((((av3)) - (cobpdl + coupdl)), "###,###,##0") & "</td>"
                coz8 = coz8 + ((((av3)) - (cobpdl + coupdl)))
                fs.WriteLine "            <td nowrap align=right>" & Format(((coacw) - coptd), "###,###,##0") & "</td>"
                coz9 = coz9 + (((coacw) - coptd))
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)), "###,###,##0") & "</td>"
                coz10 = coz10 + (((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)))
                
                If (((av3)) - (cobpdl + coupdl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)) / (((av3)) - (cobpdl + coupdl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)), "###,###,##0") & "</td>"
                coz11 = coz11 + (((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)))
                fs.WriteLine "            <td nowrap align=right>" & Format((((coacw) - coptd) - coytd), "###,###,##0") & "</td>"
                coz12 = coz12 + ((((coacw) - coptd) - coytd))
                fs.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)), "###,###,##0") & "</td>"
                coz13 = coz13 + ((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)))
                
                If ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) = 0 Then
                fs.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs.WriteLine "            <td nowrap align=right>" & Format(((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)) / ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        </tr>"
pl.MoveNext
Wend
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                '--------------------CO
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
                
                
                '--------------------
assad1:

hh.MoveNext
Wend
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1 + coq1), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2 + coq2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3 + coq3), "###,###,##0") & "</td>"
                 fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q5 + coq5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6 + coq6, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7 + coq7), "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6 + coq6) - (q7 + coq7), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(((((q6 + coq6) - (q7 + coq7)) / (q7 + coq7)) * 100), "###,###,##0") & "</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8 + coq8), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10 + coq10), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q10 + coq10) / (q8 + coq8)) * 100), 2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11 + coq11), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q11 + coq11)) * 100), 2), "###,###,##0") & "</td>"
                
                fs.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                'oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.exin desc", Cn, 3, 2
                While Not oi.EOF
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + oi!ctd
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                d6 = d6 + oi!chng
                fs.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!chng * -1)
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>" & Format((-d6), "###,###,##0") & "</td>"
                d7 = -d6
                 fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                'Profit & GP%
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d7 + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                
                
                
                'potential items
                
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                               
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "            <td nowrap align=right>NA</td>"
                fs.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                
                fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                    
                fs.WriteLine "        </tr>"
                
                
                    'estimated profit before tax
                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "            <td nowrap align=right>" & Format((d7 + (q13 + coq13)), "###,###,##0") & "</td>"
                fs.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs.WriteLine "        </tr>"
                ' Notes Column at the bottom
                
                Dim rsNotes As New ADODB.Recordset
                If rsNotes.State Then rsNotes.Close
                rsNotes.Open "select * from tblL0Notes", Cn, 3, 2
                While Not rsNotes.EOF
                
                fs.WriteLine " <tr><td align='left' valign='top' colspan=22>"
                fs.WriteLine "   <table border=1 class=TableFont width=100% cellspacing=0 BORDERCOLOR=GRAY>"
                fs.WriteLine " <tr >"
                fs.WriteLine "     <td width='50%' valign='top' rowspan='6'>" & Replace(rsNotes(1), vbNewLine, "<br>") & "</td>"
                fs.WriteLine "    <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "   </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "     <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Prepared By:" & rsNotes(2) & "</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "       <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine "   <td width='25%'  valign='top'>Reviewed By:" & rsNotes(3) & "</td>"
                fs.WriteLine " </tr>"
                fs.WriteLine "  <tr>"
                fs.WriteLine "         <td width='50%' colspan='2'>&nbsp;<br>&nbsp;</td>"
                fs.WriteLine "  </tr>"
                fs.WriteLine " <tr>"
                fs.WriteLine "  <td width='25%'  valign='top'>Date:</td>"
                fs.WriteLine " <td width='25%'  valign='top'>Approved By:" & rsNotes(4) & "</td>"
                rsNotes.MoveNext
                Wend
                fs.WriteLine " </tr>"
                fs.WriteLine " </table>"
                fs.WriteLine "      </td>  </tr>"
                If rsNotes.State Then rsNotes.Close
                
                fs.WriteLine " </table>"
    
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Public Sub rephtmlmainsave()

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
                fs1.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap> (PART-A)</td>"
                fs1.WriteLine "            <td align=center colspan=11 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
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
                fs1.WriteLine "            <td ><font color=white>ProjKey</td>"
           
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
Dim bpg1 As Double
q1 = 0: q2 = 0: q3 = 0: q4 = 0: q5 = 0: q6 = 0: q7 = 0: q8 = 0: q9 = 0: q10 = 0: q11 = 0: q12 = 0: q13 = 0: bpg1 = 0
'-------------------------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
Dim cobpg1 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0: cobpg1 = 0


'----------------------------
 Dim jh As String
 Dim hh As New ADODB.Recordset
 If hh.State Then hh.Close
 hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "'  order by bd_projectkey", Cn, 3, 2
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
Dim bpg As Double
z1 = 0: z2 = 0: z3 = 0: z4 = 0: z5 = 0: z6 = 0: z7 = 0: z8 = 0: z9 = 0: z10 = 0: z11 = 0: z12 = 0: z13 = 0: bpg = 0
'--------------------------CO
Dim coz1 As Double
Dim coz2 As Double
Dim coz3 As Double
Dim coz4 As Double
Dim coz5 As Double
Dim coz6 As Double
Dim coz7 As Double
Dim coz8 As Double
Dim coz9 As Double
Dim coz10 As Double
Dim coz11 As Double
Dim coz12 As Double
Dim coz13 As Double
Dim cobpg As Double
coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0: cobpg = 0


'---------------------------
 Dim pl As New ADODB.Recordset
 If pl.State Then pl.Close
 pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key = '" & hh(0) & "' order by proj_key", Cn, 3, 2
 While Not pl.EOF
 
      
      
                ' main
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        
               ' co
                        Dim cobdg As Double
                        Dim cobcw As Double
                        Dim coacw As Double
                        Dim coect As Double
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close
abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If
' co
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = coabc(0)
cobcw = coabc(1)
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = coct1(0)
coect = coct1(1)
End If
'co
                
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
    
    Dim coa1 As Double
    Dim coa2 As Double
    Dim coa3 As Double
    Dim coa4 As Double
    Dim coa5 As Double
    Dim cobvo As Double
    coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0: cobvo = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''----------------------main
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
   End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
   If Not rv3.EOF Then
   a4 = rv3(0)
    End If
        
   Dim rv4 As New ADODB.Recordset
   If rv4.State Then rv4.Close
   rv4.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not rv4.EOF Then
   bvo = rv4(0)
   End If
'''
'----------------------------------------------------
''''''''----------------------CO
   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not corv.EOF Then
   coa1 = corv(0)
   End If
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not corv1.EOF Then
   coa2 = corv1(0)
   End If
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not corv2.EOF Then
   coa3 = corv2(0)
   End If
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
   If Not corv3.EOF Then
   coa4 = corv3(0)
    End If
        
   Dim corv4 As New ADODB.Recordset
   If corv4.State Then corv4.Close
   corv4.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT VO' ", Cn, 3, 2
   If Not corv4.EOF Then
   cobvo = corv4(0)
   End If

'''
'----------------------------------------------------main

            Dim asam As Double
            Dim esam As Double
            
            asam = 0: esam = 0
            
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pl(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
            If a1 = Null Then a1 = 0
            If a2 = Null Then a2 = 0
            If a3 = Null Then a3 = 0
            
            If IsNull(a1) Then a1 = 0
            If IsNull(a2) Then a2 = 0
            If IsNull(a3) Then a3 = 0
            If IsNull(a4) Then a4 = 0
            
            a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
'-------------------------------------CO

'------------------------------

Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency ='" & jn(1) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   '-----------------------------------------------------------
   
   
    
                Dim bgv As Double
                bgv = 0
                Dim cobgv As Double
                cobgv = 0
                bgv = CDbl(a1) + CDbl(bvo)
                cobgv = CDbl(coa1) + CDbl(cobvo)
                Dim StrFC As String
                StrFC = pl(0) & " - " & "MAIN"
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & StrFC & "</td>"
 
                fs1.WriteLine "            <td nowrap align=right>" & Format(bgv, "###,###,##0") & "</td>"
                z1 = z1 + (bgv)
                fs1.WriteLine "            <td nowrap align=right>" & Format(bdg, "###,###,##0") & "</td>"
                z2 = z2 + bdg
                fs1.WriteLine "            <td nowrap align=right>" & Format((bgv - bdg), "###,###,##0") & "</td>"
                z3 = z3 + (bgv - bdg)
                If a1 = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((bgv) - bdg) / (bgv)) * 100), "###,###,##0") & "</td>"
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
                fs1.WriteLine "            <td nowrap align=right>" & Format((a5) - a4, "###,###,##0") & "</td>"
                z8 = z8 + ((a5) - a4)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a5)), "###,###,##0") & "</td>"
                z9 = z9 + ((a5))
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
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a5)) - acw), "###,###,##0") & "</td>"
                z13 = z13 + (((a5)) - acw)
                If (((a5 + av3))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - acw) / ((a5))) * 100), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
                
    '----------------------CO
    Dim StrCF As String
    StrCF = pl(0) & " - " & "CO"
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & StrCF & "</td>"
 
                fs1.WriteLine "            <td nowrap align=right>" & Format(cobgv, "###,###,##0") & "</td>"
                coz1 = coz1 + (cobgv)
                fs1.WriteLine "            <td nowrap align=right>" & Format(cobdg, "###,###,##0") & "</td>"
                coz2 = coz2 + cobdg
                fs1.WriteLine "            <td nowrap align=right>" & Format((cobgv - cobdg), "###,###,##0") & "</td>"
                coz3 = coz3 + (cobgv - cobdg)
                If coa1 = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((cobgv) - cobdg) / (cobgv)) * 100), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format((coa1 + coa2 + coa3), "###,###,##0") & "</td>"
                coz4 = coz4 + (coa1 + coa2 + coa3)
                fs1.WriteLine "            <td nowrap align=right>" & Format((coacw + coect), "###,###,##0") & "</td>"
                coz5 = coz5 + (coacw + coect)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - (coacw + coect)), "###,###,##0") & "</td>"
                coz6 = coz6 + ((coa1 + coa2 + coa3) - (coacw + coect))
                If (coa1 + coa2 + coa3) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - (coacw + coect)) / (coa1 + coa2 + coa3)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(coa4, "###,###,##0") & "</td>"
                coz7 = coz7 + coa4
                fs1.WriteLine "            <td nowrap align=right>" & Format((av3) - coa4, "###,###,##0") & "</td>"
                coz8 = coz8 + ((av3) - coa4)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((av3)), "###,###,##0") & "</td>"
                coz9 = coz9 + ((av3))
                If (coacw + coect) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((coacw) / (coacw + coect)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(cobcw, "###,###,##0") & "</td>"
                coz10 = coz10 + cobcw
                fs1.WriteLine "            <td nowrap align=right>" & Format(coacw, "###,###,##0") & "</td>"
                coz11 = coz11 + coacw
                fs1.WriteLine "            <td nowrap align=right>" & Format((cobcw - coacw), "###,###,##0") & "</td>"
                coz12 = coz12 + (cobcw - coacw)
                fs1.WriteLine "            <td nowrap align=right>" & Format((((av3)) - coacw), "###,###,##0") & "</td>"
                coz13 = coz13 + (((av3)) - coacw)
                If (((coa5 + av3))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - coacw) / ((av3))) * 100), "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
'--------------CO
                
pl.MoveNext
Wend
               
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                'CO
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
 
assad:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
               
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q1 + coq1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q2 + coq2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q3 + coq3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q5 + coq5, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q6 + coq6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q6 + coq6) / (q4 + coq4)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q7 + coq7, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q8 + coq8, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                ''wrk
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q11 + coq11) / (q5 + coq5)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q10 + coq10, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q11 + coq11, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q9 + coq9)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "        </tr>"
                
                
                Dim d1, d2, d3, d4, d5, d6, d7, dinc, dinc1 As Integer
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0: dinc = 0: dinc1 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                
                 fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
'                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!bdgt, "###,###,##0") & "</td>"
                d1 = d1 + oi!bdgt
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!bdgt * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!bdgt * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                If oi!exin = "Expenditure" Then
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                d3 = d3 + oi!eac
                fs1.WriteLine "            <td nowrap align=right>" & Format(((oi!eac) * -1), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac * -1)
                ElseIf oi!exin = "Income" Then
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!eac, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                dinc = dinc + oi!eac
                fs1.WriteLine "            <td nowrap align=right>" & Format(((oi!eac)), "###,###,##0") & "</td>"
                d4 = d4 + (oi!eac)
                End If

                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!bcwp, "###,###,##0") & "</td>"
                d5 = d5 + oi!bcwp
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!acwp, "###,###,##0") & "</td>"
                d6 = d6 + oi!acwp
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!bcwp - oi!acwp), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!bcwp - oi!acwp)
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!acwp * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                  
                oi.MoveNext
                Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp + Nett O/H Recovery</td>"
'                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(((d1) * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(dinc, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((dinc - d3), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d5, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d6, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5 - d6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6 * -1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
'                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + (q3 + coq3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d3) * -1) + (q6 + coq6)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                
                
                'potential items
                Dim p1 As Double
                Dim p2 As Double
                Dim p3 As Double
                p1 = 0: p2 = 0: p3 = 0
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                
                 ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
'                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(pti(0), "###,###,##0") & "</td>"
               p1 = p1 + pti(0)
                fs1.WriteLine "            <td nowrap align=right>" & Format(pti(1), "###,###,##0") & "</td>"
               p2 = p2 + pti(1)
                fs1.WriteLine "            <td nowrap align=right>" & Format((pti(0) - pti(1)), "###,###,##0") & "</td>"
                p3 = p3 + (pti(0) - pti(1))
                If pti(0) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((pti(0) - pti(1)) / pti(0)) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                 fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
'                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p2, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(p3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                 'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
'                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d1) * -1) + (q3 + coq3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((d3) * -1) + (q6 + coq6)) + (p3 + cop3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((((d6) * -1) + (q13 + coq13)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                
        fs1.WriteLine " </table>"
    
   
  WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"


End Sub

Public Sub rephtml1mainsave()
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
                fs1.WriteLine "            <td colspan=6> " & GetCompanyName & "</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap> PROJECT REVENUE & COST REPORT - L0 COMPANY LEVEL</td>"
                fs1.WriteLine "            <td align=center colspan=2 nowrap> (PART-B)</td>"
                fs1.WriteLine "            <td align=center colspan=6 nowrap> CuttOffDate: " & main.DTPcutdate1.Value & "</td>"
               
                fs1.WriteLine "        </tr>"
    
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4><font color=white> Date :" & Format(Date, "dd/MM/yyyy") & "</td>"
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
'-----------CO
Dim coq1 As Double
Dim coq2 As Double
Dim coq3 As Double
Dim coq4 As Double
Dim coq5 As Double
Dim coq6 As Double
Dim coq7 As Double
Dim coq8 As Double
Dim coq9 As Double
Dim coq10 As Double
Dim coq11 As Double
Dim coq12 As Double
Dim coq13 As Double
coq1 = 0: coq2 = 0: coq3 = 0: coq4 = 0: coq5 = 0: coq6 = 0: coq7 = 0: coq8 = 0: coq9 = 0: coq10 = 0: coq11 = 0: coq12 = 0: coq13 = 0
'-------------
                
Dim jh As String

Dim hh As New ADODB.Recordset
If hh.State Then hh.Close
hh.Open "select DISTINCT(bd_projectkey) from cost where bd_year='" & cbo_year.Text & "' order by bd_projectkey", Cn, 3, 2
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
'----------------------CO
Dim coz1 As Double
Dim coz2 As Double
Dim coz3 As Double
Dim coz4 As Double
Dim coz5 As Double
Dim coz6 As Double
Dim coz7 As Double
Dim coz8 As Double
Dim coz9 As Double
Dim coz10 As Double
Dim coz11 As Double
Dim coz12 As Double
Dim coz13 As Double

coz1 = 0: coz2 = 0: coz3 = 0: coz4 = 0: coz5 = 0: coz6 = 0: coz7 = 0: coz8 = 0: coz9 = 0: coz10 = 0: coz11 = 0: coz12 = 0: coz13 = 0

'---------------------
Dim pl As New ADODB.Recordset
If pl.State Then pl.Close
pl.Open "select DISTINCT(proj_key),proj_title from projectmaster where proj_key ='" & hh(0) & "' order by proj_key", Cn, 3, 2
While Not pl.EOF
                        Dim bdg As Double
                        Dim bcw As Double
                        Dim acw As Double
                        Dim ect As Double
                        Dim eac As Double
                        eac = 0: bdg = 0: bcw = 0: acw = 0: ect = 0
                        '-------------CO
                        Dim cobdg As Double
                        Dim cobcw As Double
                        Dim coacw As Double
                        Dim coect As Double
                        Dim coeac As Double
                        coeac = 0: cobdg = 0: cobcw = 0: coacw = 0: coect = 0
                        
                        '---------------
Dim abc As New ADODB.Recordset
If abc.State Then abc.Close

abc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='MAIN' and c.bd_costtype='B' ", Cn, 3, 2
If Not abc.EOF Then
bdg = abc(0)
bcw = abc(1)
End If
                          
Dim ct1 As New ADODB.Recordset
If ct1.State Then ct1.Close
ct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='MAIN' and c.bd_costtype='E' ", Cn, 3, 2
If Not ct1.EOF Then
acw = ct1(0)
ect = ct1(1)
End If
'---------------CO
Dim coabc As New ADODB.Recordset
If coabc.State Then coabc.Close
coabc.Open "select SUM(c.bd_extdamt),SUM(c.bd_bcwpamt)  from  cost c ,jobcharge j , jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "'  and jn.type='CO' and c.bd_costtype='B' ", Cn, 3, 2
If Not coabc.EOF Then
cobdg = coabc(0)
cobcw = coabc(1)
End If
                          
Dim coct1 As New ADODB.Recordset
If coct1.State Then coct1.Close
coct1.Open "select SUM(c.bd_extdamt),SUM(c.bd_e_extdamt)  from  cost c ,jobcharge j, jobno jn  where c.bd_jobcharge=j.job_code and c.bd_projectkey=j.job_proj_key and j.jobno=jn.jobno_code and c.bd_projectkey='" & pl(0) & "' and jn.type='CO' and c.bd_costtype='E' ", Cn, 3, 2
If Not coct1.EOF Then
coacw = coct1(0)
coect = coct1(1)
End If

'---------------
                
   Dim a1 As Double
   Dim a2 As Double
   Dim a3 As Double
   Dim a4 As Double
   Dim a5 As Double
   a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Dim revt1 As Double
    Dim revt2 As Double
    revt1 = 0: revt2 = 0
   ''''''''CO
   Dim coa1 As Double
   Dim coa2 As Double
   Dim coa3 As Double
   Dim coa4 As Double
   Dim coa5 As Double
   coa1 = 0: coa2 = 0: coa3 = 0: coa4 = 0: coa5 = 0
    Dim corevt1 As Double
    Dim corevt2 As Double
    corevt1 = 0: corevt2 = 0
   ''''''''
 
   
   '----------
   Dim rv As New ADODB.Recordset
   If rv.State Then rv.Close
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not rv.EOF
   a1 = a1 + rv(0)
   rv.MoveNext
   Wend
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not rv1.EOF
   a2 = a2 + rv1(0)
   rv1.MoveNext
   Wend
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not rv2.EOF
   a3 = a3 + rv2(0)
   rv2.MoveNext
   Wend
   
   Dim rv3 As New ADODB.Recordset
   If rv3.State Then rv3.Close
   rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not rv3.EOF
    a4 = a4 + rv3(0)
    rv3.MoveNext
    Wend
        
'''   Dim rv4 As New ADODB.Recordset
'''   If rv4.State Then rv4.Close
'''   rv4.Open "select rev_totamount from revenue where rev_projcode='" & pl(0) & "'  and rev_type='UBL' ", Cn, 3, 2
'''   While Not rv4.EOF
'''   a5 = a5 + rv4(0)
'''   rv4.MoveNext
'''   Wend


'-------------CO

   Dim corv As New ADODB.Recordset
   If corv.State Then corv.Close
   corv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   While Not corv.EOF
   coa1 = coa1 + corv(0)
   corv.MoveNext
   Wend
   
   Dim corv1 As New ADODB.Recordset
   If corv1.State Then corv1.Close
   corv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   While Not corv1.EOF
   coa2 = coa2 + corv1(0)
   corv1.MoveNext
   Wend
   
   Dim corv2 As New ADODB.Recordset
   If corv2.State Then corv2.Close
   corv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   While Not corv2.EOF
   coa3 = coa3 + corv2(0)
   corv2.MoveNext
   Wend
   
   Dim corv3 As New ADODB.Recordset
   If corv3.State Then corv3.Close
   corv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pl(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    While Not corv3.EOF
    coa4 = coa4 + corv3(0)
    corv3.MoveNext
    Wend



'------------
            

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
        If a1 = Null Then a1 = 0
        If a2 = Null Then a2 = 0
        If a3 = Null Then a3 = 0
        
        If IsNull(a1) Then a1 = 0
        If IsNull(a2) Then a2 = 0
        If IsNull(a3) Then a3 = 0
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)

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
                    
                    
                    
                    
                    Dim bpdl As Double
                    Dim bydl As Double
                    Dim updl As Double
                    Dim uydl As Double
                    bpdl = 0: bydl = 0: updl = 0: uydl = 0
                    Dim pt As New ADODB.Recordset
                    If pt.State Then pt.Close
                    pt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='MAIN'", Cn, 3, 2
                    While Not pt.EOF
                        bpdl = bpdl + pt!ptd_lye_revn
                        bydl = bydl + pt!ytd_lme_revn
                        updl = updl + pt!ptd_lye_revn1
                        uydl = uydl + pt!ytd_lme_revn1
                    pt.MoveNext
                    Wend
                    
                    'CO
                    
                    Dim cobpdl As Double
                    Dim cobydl As Double
                    Dim coupdl As Double
                    Dim couydl As Double
                    cobpdl = 0: cobydl = 0: coupdl = 0: couydl = 0
                    Dim copt As New ADODB.Recordset
                    If copt.State Then copt.Close
                    copt.Open "select * from projecttransaction where pk_projkey='" & pl(0) & "' and notes='CO'", Cn, 3, 2
                    While Not copt.EOF
                        cobpdl = cobpdl + copt!ptd_lye_revn
                        cobydl = cobydl + copt!ytd_lme_revn
                        coupdl = coupdl + copt!ptd_lye_revn1
                        couydl = couydl + copt!ytd_lme_revn1
                    copt.MoveNext
                    Wend
                    '------------------
                        Dim ytd As Double
                        Dim ptd As Double
                        ytd = 0: ptd = 0
                        
                        Dim ctr As New ADODB.Recordset
                        If ctr.State Then ctr.Close
                        ctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='MAIN' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not ctr.EOF Then
                        ytd = ctr(0)
                        ptd = ctr(1)
                        End If
                                        
                                 'co
                        Dim coytd As Double
                        Dim coptd As Double
                        coytd = 0: coptd = 0
                        
                        Dim coctr As New ADODB.Recordset
                        If coctr.State Then coctr.Close
                        coctr.Open "select SUM(ytd_lme_cost),SUM(ptd_lye_cost) from transaction1 t, jobno j where t.jobno=j.jobno_code and j.type='CO' and projkey='" & pl(0) & "'", Cn, 3, 2
                        If Not coctr.EOF Then
                        coytd = coctr(0)
                        coptd = coctr(1)
                        End If
                        '-------------
                                        
                 Dim StrFC As String
                 StrFC = pl(0) & " - " & "MAIN"
                                        
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & StrFC & "</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(((a1 + a2 + a3) - ((a5))), "###,###,##0") & "</td>"
                z1 = z1 + ((a1 + a2 + a3) - ((a5)))
                fs1.WriteLine "            <td nowrap align=right>" & Format(ect, "###,###,##0") & "</td>"
                z2 = z2 + ect
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a1 + a2 + a3) - ((a5))) - ect), "###,###,##0") & "</td>"
                z3 = z3 + (((a1 + a2 + a3) - ((a5))) - ect)
                If ((a1 + a2 + a3) - ((a5))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a1 + a2 + a3) - ((a5))) - ect) / ((a1 + a2 + a3) - ((a5)))) * 100, "###,###,##0") & "</td>"
                End If
                
                
                fs1.WriteLine "            <td nowrap align=right>" & Format((bpdl + updl), "###,###,##0") & "</td>"
                z4 = z4 + (bpdl + updl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ptd), "###,###,##0") & "</td>"
                z5 = z5 + ptd
                
                
                fs1.WriteLine "            <td nowrap align=right>" & Format((bydl + uydl), "###,###,##0") & "</td>"
                z6 = z6 + (bydl + uydl)
                fs1.WriteLine "            <td nowrap align=right>" & Format((ytd), "###,###,##0") & "</td>"
                z7 = z7 + ytd
                fs1.WriteLine "            <td nowrap align=right>" & Format((((a5)) - (bpdl + updl)), "###,###,##0") & "</td>"
                z8 = z8 + (((a5)) - (bpdl + updl))
                fs1.WriteLine "            <td nowrap align=right>" & Format(((acw) - ptd), "###,###,##0") & "</td>"
                z9 = z9 + ((acw) - ptd)
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - ((acw) - ptd)), "###,###,##0") & "</td>"
                z10 = z10 + ((((a5)) - (bpdl + updl)) - ((acw) - ptd))
                If (((a5)) - (bpdl + updl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - ((acw) - ptd)) / (((a5)) - (bpdl + updl))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((a5)) - (bpdl + updl)) - (bydl + uydl)), "###,###,##0") & "</td>"
                z11 = z11 + ((((a5)) - (bpdl + updl)) - (bydl + uydl))
                fs1.WriteLine "            <td nowrap align=right>" & Format((((acw) - ptd) - ytd), "###,###,##0") & "</td>"
                z12 = z12 + (((acw) - ptd) - ytd)
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)), "###,###,##0") & "</td>"
                z13 = z13 + (((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd))
                If ((((a5 + av3)) - (bpdl + updl)) - (bydl + uydl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((((a5)) - (bpdl + updl)) - (bydl + uydl)) - (((acw) - ptd) - ytd)) / ((((a5)) - (bpdl + updl)) - (bydl + uydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
                
                
                '-------------------CO
                 Dim StrCF As String
                 StrCF = pl(0) & " - " & "CO"
                   fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=3 nowrap>" & pl(1) & "</td>"
                fs1.WriteLine "            <td nowrap>" & StrCF & "</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(((coa1 + coa2 + coa3) - ((av3))), "###,###,##0") & "</td>"
                coz1 = coz1 + ((coa1 + coa2 + coa3) - ((av3)))
                fs1.WriteLine "            <td nowrap align=right>" & Format(coect, "###,###,##0") & "</td>"
                coz2 = coz2 + coect
                fs1.WriteLine "            <td nowrap align=right>" & Format((((coa1 + coa2 + coa3) - ((av3))) - coect), "###,###,##0") & "</td>"
                coz3 = coz3 + (((coa1 + coa2 + coa3) - ((av3))) - coect)
                If ((coa1 + coa2 + coa3) - ((av3))) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((coa1 + coa2 + coa3) - ((av3))) - coect) / ((coa1 + coa2 + coa3) - ((av3)))) * 100, "###,###,##0") & "</td>"
                
                End If
                fs1.WriteLine "            <td nowrap align=right>" & Format((cobpdl + coupdl), "###,###,##0") & "</td>"
                coz4 = coz4 + ((cobpdl + coupdl))
                fs1.WriteLine "            <td nowrap align=right>" & Format((coptd), "###,###,##0") & "</td>"
                coz5 = coz5 + ((coptd))
                fs1.WriteLine "            <td nowrap align=right>" & Format((cobydl + couydl), "###,###,##0") & "</td>"
                coz6 = coz6 + ((cobydl + couydl))
                fs1.WriteLine "            <td nowrap align=right>" & Format((coytd), "###,###,##0") & "</td>"
                coz7 = coz7 + ((coytd))
                fs1.WriteLine "            <td nowrap align=right>" & Format((((av3)) - (cobpdl + coupdl)), "###,###,##0") & "</td>"
                coz8 = coz8 + ((((av3)) - (cobpdl + coupdl)))
                fs1.WriteLine "            <td nowrap align=right>" & Format(((coacw) - coptd), "###,###,##0") & "</td>"
                coz9 = coz9 + (((coacw) - coptd))
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)), "###,###,##0") & "</td>"
                coz10 = coz10 + (((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)))
                
                If (((av3)) - (cobpdl + coupdl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - ((coacw) - coptd)) / (((av3)) - (cobpdl + coupdl))) * 100, "###,###,##0") & "</td>"
                End If
                
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)), "###,###,##0") & "</td>"
                coz11 = coz11 + (((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)))
                fs1.WriteLine "            <td nowrap align=right>" & Format((((coacw) - coptd) - coytd), "###,###,##0") & "</td>"
                coz12 = coz12 + ((((coacw) - coptd) - coytd))
                fs1.WriteLine "            <td nowrap align=right>" & Format((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)), "###,###,##0") & "</td>"
                coz13 = coz13 + ((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)))
                If ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) = 0 Then
                fs1.WriteLine "            <td nowrap align=right>0</td>"
                Else
                fs1.WriteLine "            <td nowrap align=right>" & Format(((((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl)) - (((coacw) - coptd) - coytd)) / ((((av3)) - (cobpdl + coupdl)) - (cobydl + couydl))) * 100, "###,###,##0") & "</td>"
                End If
                fs1.WriteLine "        </tr>"
                
                fs1.WriteLine "        </tr>"
                                          
                                        
pl.MoveNext
Wend
                q1 = q1 + z1
                q2 = q2 + z2
                q3 = q3 + z3
                q4 = q4 + z4
                q5 = q5 + z5
                q6 = q6 + z6
                q7 = q7 + z7
                q8 = q8 + z8
                q9 = q9 + z9
                q10 = q10 + z10
                q11 = q11 + z11
                q12 = q12 + z12
                q13 = q13 + z13
                '--------------------CO
                coq1 = coq1 + coz1
                coq2 = coq2 + coz2
                coq3 = coq3 + coz3
                coq4 = coq4 + coz4
                coq5 = coq5 + coz5
                coq6 = coq6 + coz6
                coq7 = coq7 + coz7
                coq8 = coq8 + coz8
                coq9 = coq9 + coz9
                coq10 = coq10 + coz10
                coq11 = coq11 + coz11
                coq12 = coq12 + coz12
                coq13 = coq13 + coz13
                
                
                '--------------------
assad1:

hh.MoveNext
Wend
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total</td>"
                
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q1 + coq1), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q2 + coq2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q3 + coq3), "###,###,##0") & "</td>"
                 fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q3 + coq3) / (q1 + coq1)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q4 + coq4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q5 + coq5), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(q6 + coq6, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q7 + coq7), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q8 + coq8), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q9 + coq9), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q10 + coq10), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q10 + coq10) / (q8 + coq8)) * 100), 2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q11 + coq11), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q12 + coq12), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((q13 + coq13), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(Round((((q13 + coq13) / (q11 + coq11)) * 100), 2), "###,###,##0") & "</td>"
                
                fs1.WriteLine "        </tr>"
                
                
                
                Dim d1, d2, d3, d4, d5, d6, d7 As Double
                d1 = 0: d2 = 0: d3 = 0: d4 = 0: d5 = 0: d6 = 0: d7 = 0
                Dim oi As New ADODB.Recordset
                If oi.State Then oi.Close
                oi.Open "select * from oitranx ot, othertransaction ott where ot.tranx=ott.ot_desc and ot.oi_year='" & cbo_year.Text & "' order by ott.ot_tranx", Cn, 3, 2
                While Not oi.EOF
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & oi!tranx & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!etc, "###,###,##0") & "</td>"
                d1 = d1 + oi!etc
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!etc * -1), "###,###,##0") & "</td>"
                d2 = d2 + (oi!etc * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format(oi!ytd, "###,###,##0") & "</td>"
                d3 = d3 + oi!ytd
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!ctd), "###,###,##0") & "</td>"
                d4 = d4 + oi!ctd
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!ctd * -1), "###,###,##0") & "</td>"
                d5 = d5 + (oi!ctd * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!chng), "###,###,##0") & "</td>"
                d6 = d6 + oi!chng
                fs1.WriteLine "            <td nowrap align=right>" & Format((oi!chng * -1), "###,###,##0") & "</td>"
                'd7 = d7 + (oi!chng * -1)
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                oi.MoveNext
                Wend
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Oth Inc/Exp+Nett O/M Recovery</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d1, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d2), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format(d3, "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d4), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d5), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((d6), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>" & Format((-d6), "###,###,##0") & "</td>"
                d7 = -d6
                 fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Estimated Profit Before Tax</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d7 + (q13 + coq13)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
                
                'potential items
                
                Dim pti As New ADODB.Recordset
                If pti.State Then pti.Close
                pti.Open "select SUM(p_revn),SUM(p_cost),p_item from potentialitem group by p_item", Cn, 3, 2
                While Not pti.EOF
                ju = Split(pti(2), "  -  ", Len(pti(2)), vbTextCompare)
                               
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>" & ju(1) & "</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "            <td nowrap align=right>NA</td>"
                fs1.WriteLine "        </tr>"
                
                
                pti.MoveNext
                Wend
                
                fs1.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap><font color=white>Total PotentialItems</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                fs1.WriteLine "            <td nowrap align=right><font color=white>NA</td>"
                    
                fs1.WriteLine "        </tr>"
                
                
                    'estimated profit before tax
                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=4 nowrap>Est.Profit B4 TAX(INC PI)</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d2 + (q3 + coq3)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d5 + (q10 + coq10)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "            <td nowrap align=right>" & Format((d7 + (q13 + coq13)), "###,###,##0") & "</td>"
                fs1.WriteLine "            <td nowrap align=right>&nbsp;</td>"
                fs1.WriteLine "        </tr>"
                
                
        fs1.WriteLine " </table>"
    
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"

End Sub
