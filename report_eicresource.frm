VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form report_eicresource 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EIC by RESOURCE/PROJECT"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   13575
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6375
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   11245
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
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00DC7E5A&
      Height          =   480
      Left            =   12480
      Picture         =   "report_eicresource.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click to Print"
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13575
      Begin VB.CommandButton cmd_clear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   5160
         TabIndex        =   0
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_search 
         Caption         =   "Search"
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txt_search 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   12480
         Picture         =   "report_eicresource.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Click to Exit"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton command2 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   12480
         Picture         =   "report_eicresource.frx":0B72
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Click to View"
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All Projects By Date"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton opt_all 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton opt_nonspread 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Non Spread"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opt_spread 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spread"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10920
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cbo_year 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.ListBox lst_prj 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   450
         Width           =   4455
      End
      Begin VB.ListBox lst_resc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   10800
         X2              =   10800
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resource"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Project"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "report_eicresource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbo_year_Change()
lst_prj.Clear
Dim f As Integer
f = 0
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
lst_prj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
End Sub

Private Sub cbo_year_Click()
lst_prj.Clear
Dim f As Integer
f = 0
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(rd.dresc_proj),p.proj_desc  from resourcedetails rd,projectmaster p,userproject u where rd.dresc_proj=p.proj_key and p.proj_key=u.project and rd.dresc_year='" & cbo_year.Text & "' and u.username ='" & main.Label2.Caption & "'  order by rd.dresc_proj", Cn, 3, 2
While Not pr.EOF
lst_prj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
End Sub
Private Sub cmd_clear_Click()
Dim Slsc As Double
Slsc = 0
For Slsc = 0 To lst_resc.ListCount - 1
lst_resc.Selected(Slsc) = False
Next Slsc
End Sub
Private Sub Check1_Click()
Dim a As Integer
If Check1.Value = 1 Then


a = 0
For a = 0 To lst_prj.ListCount - 1
lst_prj.Selected(a) = True
Next a
lst_prj.Enabled = False

Else
a = 0
For a = 0 To lst_prj.ListCount - 1
lst_prj.Selected(a) = False
Next a
lst_prj.Enabled = True

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

Private Sub cmd_search_Click()
Dim Sls As Double
Sls = 0
For Sls = 0 To lst_resc.ListCount - 1
If InStr(lst_resc.List(Sls), txt_search.Text) Then
lst_resc.Selected(Sls) = True
End If

Next Sls

End Sub

Private Sub command2_Click()
frmBusy.Show
SetParent frmBusy.hwnd, report_eicresource.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call flex_dataallreport
Unload frmBusy

End Sub
Public Sub flex_dataallreport()
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

 i = 0
'With flex_grid
'        .Rows = 1
For i = 0 To lst_resc.ListCount - 1
If lst_resc.Selected(i) = True Then
nmm = Split(lst_resc.List(i), "  -  ", Len(lst_resc.List(i)), vbTextCompare)

'''Dim j As Integer
''' j = 0
''' For j = 0 To lst_prj.ListCount - 1
''' If lst_prj.Selected(j) = True Then
''' nmd = Split(lst_prj.List(j), "  -  ", Len(lst_prj.List(j)), vbTextCompare)
 
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
Dim jc As New ADODB.Recordset
        
        If jc.State Then jc.Close
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
If opt_spread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "'  and bd_costtype='E' and bd_spread <> 'NA'   order by bd_sdate,bd_edate", Cn, 3, 2
'fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey='" & nmd(0) & "' and bd_costtype='E' and bd_spread <> 'NA'   order by bd_sdate,bd_edate", Cn, 3, 2
   While Not fldata.EOF
   
        
        cnt = cnt + 1 '********************************
                If cnt >= 53 Then
                fs.WriteLine "</table><P></P>"
                RPTHEADING fs
                cnt = 0
                End If
        fs.WriteLine "         <tr bgcolor=white height=15 class=TableFont>"
        fs.WriteLine "            <td >" & fldata!bd_sdate & "</td>"
        fs.WriteLine "            <td >" & fldata!bd_edate & "</td>"
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & "  -  " & jc(0) & "</td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & " </td>"
        End If
        jc.Close
        fs.WriteLine "            <td >" & fldata!bd_curr & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_xchg, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_qty, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_unitrate, "###,###,##0.00") & " </td>"
            
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_spread & "  -  " & spr(0) & " </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_spread & "  </td>"
        End If
        spr.Close
        fs.WriteLine "            <td >" & fldata!bd_type & "  </td>"
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_costcode & "  -  " & cs(0) & "  </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_costcode & "   </td>"
        End If
        cs.Close
        fs.WriteLine "            <td >" & fldata!bd_notes & "  </td>"
        fs.WriteLine "        </tr>"
        fldata.MoveNext
    Wend


ElseIf opt_nonspread.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "' and bd_costtype='E' and bd_spread ='NA'  order by bd_jobcharge,bd_spread", Cn, 3, 2
'fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey='" & nmd(0) & "' and bd_costtype='E' and bd_spread = 'NA'   order by bd_sdate,bd_edate", Cn, 3, 2
    
 While Not fldata.EOF
       
        
        cnt = cnt + 1 '********************************
                If cnt >= 53 Then
                fs.WriteLine "</table><P></P>"
                RPTHEADING fs
                cnt = 0
                End If
        fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        fs.WriteLine "            <td >" & fldata!bd_sdate & "</td>"
        fs.WriteLine "            <td >" & fldata!bd_edate & "</td>"
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & "  -  " & jc(0) & "</td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & " </td>"
        End If
        jc.Close
        fs.WriteLine "            <td >" & fldata!bd_curr & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_xchg, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_qty, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_unitrate, "###,###,##0.00") & " </td>"
            
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_spread & "  -  " & spr(0) & " </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_spread & "  </td>"
        End If
        spr.Close
        fs.WriteLine "            <td >" & fldata!bd_type & "  </td>"
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_costcode & "  -  " & cs(0) & "  </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_costcode & "   </td>"
        End If
        cs.Close
        fs.WriteLine "            <td >" & fldata!bd_notes & "  </td>"
        fs.WriteLine "        </tr>"
        
        fldata.MoveNext
    Wend



ElseIf opt_all.Value = True Then
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "' and bd_costtype='E'   order by bd_sdate,bd_edate", Cn, 3, 2
'fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey='" & nmd(0) & "' and bd_costtype='E'   order by bd_sdate,bd_edate", Cn, 3, 2
    
   While Not fldata.EOF
     
        
        cnt = cnt + 1 '********************************
                If cnt >= 53 Then
                fs.WriteLine "</table><P></P>"
                RPTHEADING fs
                cnt = 0
                End If
        fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        fs.WriteLine "            <td >" & fldata!bd_sdate & "</td>"
        fs.WriteLine "            <td >" & fldata!bd_edate & "</td>"
        
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & "  -  " & jc(0) & "</td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_JobCharge & " </td>"
        End If
        jc.Close
        fs.WriteLine "            <td >" & fldata!bd_curr & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_xchg, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_qty, "###,###,##0.00") & " </td>"
        fs.WriteLine "            <td >" & Format(fldata!bd_unitrate, "###,###,##0.00") & " </td>"
            
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_spread & "  -  " & spr(0) & " </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_spread & "  </td>"
        End If
        spr.Close
        fs.WriteLine "            <td >" & fldata!bd_type & "  </td>"
        
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        fs.WriteLine "            <td >" & fldata!bd_costcode & "  -  " & cs(0) & "  </td>"
        Else
        fs.WriteLine "            <td >" & fldata!bd_costcode & "   </td>"
        End If
        cs.Close
        fs.WriteLine "            <td >" & fldata!bd_notes & "  </td>"
        fs.WriteLine "        </tr>"
        
        fldata.MoveNext
    Wend



Else
End If


    
'''End If
'''Next j
End If
Next i
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"



'End With
 
End Sub




Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"


            
                fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
                fs.WriteLine "            <td colspan=11><font color=white><b>" & GetCompanyName & "</td>"
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
                
                fs.WriteLine "            <td colspan=6><font color=white><b>EIC BY RESOURCE</td>"
                fs.WriteLine "           <td colspan=2><font color=white><b>PrintDate</td>"
                fs.WriteLine "           <td colspan=3><font color=white>" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "        </tr>"
                


                fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                fs.WriteLine "            <td Nowrap><font color=white>StartDate</td>"
                fs.WriteLine "            <td Nowrap><font color=white>EndDate</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Jobcharge</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Curr</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Xchg</td>"
                
                 fs.WriteLine "            <td Nowrap><font color=white>Qty</td>"
                fs.WriteLine "            <td Nowrap><font color=white>UnitRate</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Spread</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Type</td>"
                fs.WriteLine "            <td Nowrap><font color=white>CostCode</td>"
                fs.WriteLine "            <td Nowrap><font color=white>Notes</td>"
                fs.WriteLine "        </tr>"
    
End Sub



Private Sub Form_Load()
main.lbltitle.Caption = "EIC By Resource/Project"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim rs2 As New ADODB.Recordset
If rs2.State Then rs2.Close
rs2.Open "select DISTINCT(bd_resccode) from cost where bd_costtype='E'  order by bd_resccode", Cn, 3, 2
While Not rs2.EOF
Dim ki As New ADODB.Recordset
If ki.State Then ki.Close
ki.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rs2(0) & "' ", Cn, 3, 2
If Not ki.EOF Then
lst_resc.AddItem rs2(0) & "  -  " & ki(0)
Else
lst_resc.AddItem rs2(0)
End If
rs2.MoveNext
Wend
rs2.Close
cbo_year.Text = Year(Date)
Dim i As Integer
i = 0
For i = 2004 To 2050
cbo_year.AddItem i
Next i

End Sub
