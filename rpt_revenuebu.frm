VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_revenuebu 
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
      Height          =   6975
      Left            =   120
      TabIndex        =   12
      Top             =   1920
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
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton cmd_co 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CO"
         Height          =   480
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to View"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   7680
         Picture         =   "rpt_revenuebu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Click to Exit"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Main"
         Height          =   480
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to View"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   6840
         Picture         =   "rpt_revenuebu.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to Print"
         Top             =   720
         Width           =   735
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   720
         Width           =   5175
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   5175
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
         TabIndex        =   7
         Top             =   720
         Width           =   585
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
         TabIndex        =   6
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   6615
      End
   End
End
Attribute VB_Name = "rpt_revenuebu"
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
         Option1.Value = 0
         Option2.Value = 0

End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_co_Click()
If cbo_proj.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor1
Unload frmBusy
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
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
   nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
        fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
        fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
        fs.WriteLine "            <td colspan=3><b>" & GetCompanyName & "</td>"
        fs.WriteLine "           <td  >Project key</td>"
        fs.WriteLine "           <td  >" & nm(0) & "</td>"
        fs.WriteLine "           <td  >JobNo</td>"
        fs.WriteLine "           <td  >SeeEndOfReport</td>"
        fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=5><b>REVENUE - BILLED & UNBILLED</td>"
                fs.WriteLine "           <td  >Cutt-off Date</td>"
                fs.WriteLine "           <td  >" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
    fs.WriteLine "            <td colspan=7><font color=white>&nbsp;</td>"
    fs.WriteLine "        </tr>"
        fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
        fs.WriteLine "            <td colspan=7><font color=white>Revenue Type</td>"
        fs.WriteLine "        </tr>"
            
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap  ><font color=white>Notes</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>xRate</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount(RM)</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Inv No</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Inv Date</td>"
   'fs.WriteLine "            <td width=200 align=center><font color=white>Notes</td>"
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

                    Dim pnh As String
                    pn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
                    pnh = Mid(pn(0), 1, 3)
  Dim jk As Double
  jk = 0
  Dim i As Integer
  i = 0
  For i = 0 To List2.ListCount - 1
  If List2.Selected(i) = True Then
  ii = Split(List2.List(i), "  -  ", Len(List2.List(i)), vbTextCompare)
  bd = 0
                    
 
                    
                    fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=7><font color=black>" & List2.List(i) & "</td>"
                    fs.WriteLine "        </tr>"
                    Dim fldata As New ADODB.Recordset
                    If fldata.State Then fldata.Close
                    fldata.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_type='BLD' and r.rev_projcode='" & pn(0) & "' and r.rev_jobno='" & ii(0) & "' order by r.rev_invoice ", Cn, 3, 2
                 
                    While Not fldata.EOF
                     fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                       If fldata!rev_tranxnotes = "" Then
                      fs.WriteLine "            <td  >&nbsp;</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldata!rev_tranxnotes & "</td>"
                      End If
                      fs.WriteLine "            <td  align=center>" & fldata!rev_Currency & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_amount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_exchange, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_totamount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=center>" & fldata!rev_invoice & "</td>"
                      fs.WriteLine "            <td  align=center>" & fldata!rev_invoicedate & "</td>"
                      bd = bd + fldata!rev_totamount
                     
                      fs.WriteLine "        </tr>"
                    fldata.MoveNext
                    Wend
                    
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><B> Billed Sub Total - " & ii(0) & "</td>"
                    fs.WriteLine "            <td  align=right><B> " & Format(bd, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                    jk = jk + bd
                    
                    
   
  End If
  Next i
  '------------------------------

  On Error Resume Next
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
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='MAIN' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
Dim asam As Double
        Dim esam As Double
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
'-----------------
  
                    
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><b>Total billed - " & pn(0) & "</td>"
                    fs.WriteLine "            <td  align=right><b> " & Format(jk, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><b>  Unbilled  Total- " & pn(0) & "</td>"
                    fs.WriteLine "            <td  align=right><b> " & Format(a5 - a4, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                   
                    fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><font color=white><b>REPORT TOTAL</td>"
                    fs.WriteLine "            <td  align=right><font color=white><b>" & Format(a5, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
  
    
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub
Public Sub nocolor1()
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
   nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
        fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
        fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
        fs.WriteLine "            <td colspan=3><b>" & GetCompanyName & "</td>"
        fs.WriteLine "           <td  >Project key</td>"
        fs.WriteLine "           <td  >" & nm(0) & "</td>"
        fs.WriteLine "           <td  >JobNo</td>"
        fs.WriteLine "           <td  >SeeEndOfReport</td>"
        fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=5><b>REVENUE - BILLED & UNBILLED (CO)</td>"
                fs.WriteLine "           <td  >Cutt-off Date</td>"
                fs.WriteLine "           <td  >" & main.DTPcutdate1.Value & "</td>"
                fs.WriteLine "        </tr>"
    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
    fs.WriteLine "            <td colspan=7><font color=white>&nbsp;</td>"
    fs.WriteLine "        </tr>"
        fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
        fs.WriteLine "            <td colspan=7><font color=white>Revenue Type</td>"
        fs.WriteLine "        </tr>"
            
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap  ><font color=white>Notes</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>xRate</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Amount(RM)</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Inv No</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Inv Date</td>"
   'fs.WriteLine "            <td width=200 align=center><font color=white>Notes</td>"
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

                    Dim pnh As String
                    pn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
                    pnh = Mid(pn(0), 1, 3)
  Dim jk As Double
  jk = 0
  Dim i As Integer
  i = 0
  For i = 0 To List2.ListCount - 1
  If List2.Selected(i) = True Then
  ii = Split(List2.List(i), "  -  ", Len(List2.List(i)), vbTextCompare)
  bd = 0
                    
 
                    
                    fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=7><font color=black>" & List2.List(i) & "</td>"
                    fs.WriteLine "        </tr>"
                    Dim fldata As New ADODB.Recordset
                    If fldata.State Then fldata.Close
                    fldata.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_type='BLD' and r.rev_projcode='" & pn(0) & "' and r.rev_jobno='" & ii(0) & "' order by r.rev_invoice ", Cn, 3, 2
                 
                    While Not fldata.EOF
                     fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                       If fldata!rev_tranxnotes = "" Then
                      fs.WriteLine "            <td  >&nbsp;</td>"
                      Else
                      fs.WriteLine "            <td  >" & fldata!rev_tranxnotes & "</td>"
                      End If
                      fs.WriteLine "            <td  align=center>" & fldata!rev_Currency & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_amount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_exchange, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=right>" & Format(fldata!rev_totamount, "###,###,##0.00") & "</td>"
                      fs.WriteLine "            <td  align=center>" & fldata!rev_invoice & "</td>"
                      fs.WriteLine "            <td  align=center>" & fldata!rev_invoicedate & "</td>"
                      bd = bd + fldata!rev_totamount
                     
                      fs.WriteLine "        </tr>"
                    fldata.MoveNext
                    Wend
                    
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><B> Billed Sub Total - " & ii(0) & "</td>"
                    fs.WriteLine "            <td  align=right><B> " & Format(bd, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                    jk = jk + bd
                    
                    
   
  End If
  Next i
  '------------------------------

  On Error Resume Next
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
   rv.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='BGT' ", Cn, 3, 2
   If Not rv.EOF Then
   a1 = rv(0)
   End If
   
   Dim rv1 As New ADODB.Recordset
   If rv1.State Then rv1.Close
   rv1.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   If Not rv1.EOF Then
   a2 = rv1(0)
    End If
   Dim av3 As Double
   Dim av2 As Double
   
   Dim jn As New ADODB.Recordset
   If jn.State Then jn.Close
   jn.Open "select (r.rev_jobno),r.rev_currency, rev_id from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(+)' ", Cn, 3, 2
   av3 = 0
   While Not jn.EOF
    Dim rvv1 As New ADODB.Recordset
   If rvv1.State Then rvv1.Close
   'rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_currency='" & jn(1) & "'", Cn, 3, 2
   rvv1.Open "select * from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(+)' and r.rev_jobno='" & jn(0) & "' and r.rev_id='" & jn(2) & "'", Cn, 3, 2
   If Not rvv1.EOF Then
   av2 = 0
   av2 = CDbl(rvv1!rev_totamount) * (CDbl(rvv1!perc) / 100)
   End If
   av3 = av3 + av2
   
   jn.MoveNext
   Wend
   Dim rv2 As New ADODB.Recordset
   If rv2.State Then rv2.Close
   rv2.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='VO(-)' ", Cn, 3, 2
   If Not rv2.EOF Then
   a3 = rv2(0)
   End If
   
    Dim rv3 As New ADODB.Recordset
    If rv3.State Then rv3.Close
    rv3.Open "select SUM(r.rev_totamount) from revenue r , jobno j where r.rev_jobno=j.jobno_code and j.type='CO' and r.rev_projcode='" & pn(0) & "'  and r.rev_type='BLD' ", Cn, 3, 2
    If Not rv3.EOF Then
    a4 = rv3(0)
    End If
        
Dim asam As Double
        Dim esam As Double
        asam = 0: esam = 0
        
                          Dim sam As New ADODB.Recordset
                          If sam.State Then sam.Close
                          sam.Open "select SUM(bd_extdamt),SUM(bd_e_extdamt) from jobcharge j, cost c ,jobno jn where j.job_code=c.bd_jobcharge and jn.jobno_code=j.jobno and jn.type='MAIN' and j.job_proj_key='" & pn(0) & "' and c.bd_costtype='E'  ", Cn, 3, 2
                          If Not sam.EOF Then
                          asam = Format(sam(0), "###,###,###,##0")
                          esam = Format(sam(1), "###,###,###,##0")
                                    
                          End If
        
 
   a5 = (CDbl(asam) / (CDbl(asam) + CDbl(esam))) * (a1 + a2 + a3)
'-----------------
  
                    
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><b>Total billed - " & pn(0) & "</td>"
                    fs.WriteLine "            <td  align=right><b> " & Format(jk, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><b>  Unbilled  Total- " & pn(0) & "</td>"
                    fs.WriteLine "            <td  align=right><b> " & Format(av3 - a4, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
                   
                    fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
                    fs.WriteLine "            <td  colspan=4><font color=white><b>REPORT TOTAL</td>"
                    fs.WriteLine "            <td  align=right><font color=white><b>" & Format(av3, "###,###,##0.00") & "</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "            <td  >&nbsp;</td>"
                    fs.WriteLine "        </tr>"
  
    
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "REVENUE BILLED/UNBILLED"
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
 
            Option1.Value = False
            Option2.Value = True
            
Me.Width = 11415
Me.Height = 9750

         
End Sub
Private Sub Option1_Click()
If Option1.Value = True Then
Dim f As Integer
f = 0
For f = 0 To List2.ListCount - 1
List2.Selected(f) = True
Next f
 
End If
 
End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List2.ListCount - 1
List2.Selected(g) = False
Next g
 
End If
 
End Sub
 


