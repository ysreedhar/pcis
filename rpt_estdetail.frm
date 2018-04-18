VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_estdetail 
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
      Height          =   6495
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   11535
      ExtentX         =   20346
      ExtentY         =   11456
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
      Width           =   11175
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9000
         Picture         =   "rpt_estdetail.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8160
         Picture         =   "rpt_estdetail.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9840
         Picture         =   "rpt_estdetail.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACWP"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ECTC"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EAC"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TranX Dates"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Left            =   4680
         Top             =   240
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5520
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   16449539
         CurrentDate     =   38099
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   720
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
         Left            =   6810
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   4290
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1200
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
         Caption         =   "JobCharge"
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
         Left            =   5880
         TabIndex        =   11
         Top             =   720
         Width           =   930
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5745
         Top             =   120
         Width           =   5415
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
         Width           =   5535
      End
   End
End
Attribute VB_Name = "rpt_estdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hgg As Integer
Private Sub cbo_proj_Click()
spp = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
List2.Clear
List1.Clear
Dim lst As String
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(jobno_code),jobno_desc from jobno where job_key='" & spp(0) & "' order by jobno_code", Cn, 3, 2
While Not rs1.EOF
List2.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
 
 Check1.Value = 1
  Check2.Value = 1
   Check3.Value = 1
   
   
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
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call estdetails
Unload frmBusy

End Sub
Public Sub estdetails()
Dim fs As Object
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
   
    


   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

Dim ddtot As Double
Dim ddtot1 As Double
Dim ddwtot2 As Double
  
 Dim cnt As Integer
 RPTHEADINGESTDETAILS fs
 cnt = 0
 
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
    ddtot = 0
    ddtot1 = 0
    ddwtot2 = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 
        
        Dim stot As Double
        Dim tot As Double
        Dim tot1 As Double
        Dim dtot As Double
        Dim atot As Double
        Dim ktot As Double
        Dim wtot As Double
        Dim wtot1 As Double
        Dim wtot2 As Double
         wtot2 = 0
        
        tot = 0:  tot1 = 0
        
        Dim l As Integer
        l = 0
        For l = 0 To List1.ListCount - 1
        If List1.Selected(l) = True Then
         nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
         ju = Split(nm(0), "-", Len(nm(0)), vbTextCompare)
         If gy(0) = ju(0) Then
                              
                    dtot = 0
                    ktot = 0
                    wtot1 = 0
                    Dim yre As String
                    Dim fl As New ADODB.Recordset
                    If fl.State Then fl.Close
                    fl.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j  where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and j.job_code='" & nm(0) & "'  and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_costtype='E' ", Cn, 3, 2
                   
                            While Not fl.EOF
                            yre = fl(0)
                            
                                  stot = 0
                                  atot = 0
                                  wtot = 0
                                  
                                                Dim fldata1 As New ADODB.Recordset
                                                If fldata1.State Then fldata1.Close
                                                fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='E' and j.job_code='" & nm(0) & "'  and j.job_desc='" & nm(1) & "'  and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
                                               
                                                While Not fldata1.EOF
         
                                                
                                                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_resccode & "</td>"
                                                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_rescname & "</td>" ''''''
                                                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_respcode & "</td>" ''''''
                                                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_vendor & "</td>" ''''''
                                                fs.WriteLine "            <td Nowrap  >" & fldata1!bd_tranx & "</td>" ''''''
                                                fs.WriteLine "            <td Nowrap align=center> " & fldata1!bd_costcode & " </td>"
                                                fs.WriteLine "            <td Nowrap align=center> " & fldata1!bd_spread & " </td>"
                                                fs.WriteLine "            <td Nowrap> " & fldata1!bd_JobCharge & " </td>"
                                                If Check4.Value = 1 Then
                                                fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_sdate, "dd/MM/yyyy") & " </td>"
                                                fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_edate, "dd/MM/yyyy") & " </td>"
                                                
                                                End If
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_qty, "###,###,##0.00") & " </td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_days, "###,###,##0.00") & " </td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_tqty, "###,###,##0.00") & " </td>"
                                                fs.WriteLine "            <td Nowrap align=center > " & fldata1!bd_uom & "</td>"
                                                fs.WriteLine "            <td Nowrap align=center> " & fldata1!bd_curr & "</td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_unitrate, "###,###,##0.00") & "  </td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_xchg, "###,###,##0.00") & " </td>"
                                                'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
                                                'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
                                                 If Check1.Value = 1 Then
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_extdamt, "###,###,##0.00") & " </td>"
                                                stot = stot + fldata1!bd_extdamt
                                                End If
                                                 If Check2.Value = 1 Then
                                               
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_days, "###,###,##0.00") & " </td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_tqty, "###,###,##0.00") & " </td>"
                                                fs.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "  </td>"
                                                atot = atot + fldata1!bd_e_extdamt
                                                End If
                                                 If Check3.Value = 1 Then
                                                fs.WriteLine "            <td Nowrap align=right> " & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "  </td>"
                                                wtot = wtot + (fldata1!bd_extdamt) + (fldata1!bd_e_extdamt)
                                                End If
               If fldata1!bd_notes <> "" Then
                                Dim jh11 As String
                                jh11 = Mid(fldata1!bd_notes, 1, 15)
                                fs.WriteLine "            <td ><b> " & jh11 & "</td>"
                                Else
                                Dim cd1 As New ADODB.Recordset
                                If cd1.State Then cd1.Close
                                cd1.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
                                If Not cd1.EOF Then
                                Dim jh1 As String
                                jh1 = Mid(cd1(0), 1, 15)
                                fs.WriteLine "            <td Nowrap> " & jh1 & "</td>"
                                End If
                                End If
                                                fs.WriteLine "       </tr>"
                                                fldata1.MoveNext
                                                Wend
                                                
                                                
                                                
                                                
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = rscd(0)
End If
                            
                            dtot = dtot + stot
                            ktot = ktot + atot
                            wtot1 = wtot1 + wtot
                            fl.MoveNext
                            Wend
                           
                                
                     tot = tot + dtot
                     tot1 = tot1 + ktot
                     wtot2 = wtot2 + wtot1
                                         
            End If
            End If
            Next l
          
        ddtot = ddtot + tot
        ddtot1 = ddtot1 + tot1
        ddwtot2 = ddwtot2 + wtot2
   End If
   Next w
 
                               
fs.WriteLine " </table>"
   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
fs.WriteLine " </table>"
   
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
  
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub

Public Sub RPTHEADINGESTDETAILS(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
 
ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
fs.WriteLine "        <tr bgcolor=white  height=25 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=17>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=15>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13>" & GetCompanyName & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=14>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13>" & GetCompanyName & "</td>"

Else
fs.WriteLine "            <td colspan=12>" & GetCompanyName & "</td>"
End If

                fs.WriteLine "           <td colspan=2><b>ProjectKey</td>"
                fs.WriteLine "           <td colspan=2 align=center>" & ff(0) & "</td>"
                fs.WriteLine "           <td><b>JobCharge</td>"
                            If Option4.Value = True Then
                            fs.WriteLine "           <td align=center>All</td>"
                            Else
                            fs.WriteLine "           <td align=center>SeeEndOfReport</td>"
                            End If
                fs.WriteLine "        </tr>"
                
 fs.WriteLine "        <tr bgcolor=white  height=25 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=17>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=15>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=14>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13>INCURRED BY JOBCHARGE(L3)</td>"

Else
fs.WriteLine "            <td colspan=12>INCURRED BY JOBCHARGE(L3)</td>"
End If
                    fs.WriteLine "           <td colspan=2><b>JobNo.</td>"
                                If Option1.Value = True Then
                                fs.WriteLine "           <td colspan=2 align=center>All</td>"
                                Else
                                fs.WriteLine "           <td colspan=2 align=center>SeeEndOfReport</td>"
                                End If
                    fs.WriteLine "           <td><b>Cut-OffDate</td>"
                    fs.WriteLine "           <td align=center>" & main.DTPcutdate1.Value & "</td>"
                    fs.WriteLine "        </tr>"
                  
fs.WriteLine "     <tr bgcolor=white  height=8 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=26><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=24><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=24><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=24><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=22><font color=white>&nbsp;</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=21><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=22><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=22><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=22><font color=white>&nbsp;</td>"
Else
fs.WriteLine "            <td colspan=20><font color=white>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"
 
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap ><font color=white> RescCde  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>RescName</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>RescResp</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>Vendor</td>"
            fs.WriteLine "            <td Nowrap align=center><font color=white>TranX</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> CostCde  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> SprdCde </td>"
    fs.WriteLine "            <td Nowrap><font color=white>JobCharge</td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=white> StartDate  </td>"
   fs.WriteLine "            <td Nowrap><font color=white> EndDate  </td>"
   End If
   fs.WriteLine "            <td Nowrap align=right><font color=white> Qty  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white> Days </td>"
   
   fs.WriteLine "            <td Nowrap align=right><font color=white> TotalQty  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> UOM  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> Curcy  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white> UnitRate  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white> xRate </td>"
'   fs.WriteLine "            <td Nowrap>DT</td>"
'   fs.WriteLine "            <td Nowrap>Escl</td>"
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>ACWP Amt(RM)</font> </td>"
   End If
   If Check2.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white> Days  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>TotQty </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>ECTC Amt(RM)  </td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>EAC Amt(RM)  </td>"
   End If
   fs.WriteLine "            <td ><font color=white>Notes/CostCde Desc </td>"
   fs.WriteLine "        </tr>"
    
End Sub


Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "TABLES - EIC COST DETAILS"
Me.Top = 10
Me.Left = 10
Me.Width = 11415
Me.Height = 9750

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
DTPicker1.Value = main.DTPcutdate1.Value
End Sub



Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub List2_Click()
List1.Clear
Option1.Value = False
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim h As Integer
h = 0
For h = 0 To List2.ListCount - 1
If List2.Selected(h) = True Then
ju = Split(List2.List(h), "  -  ", Len(List2.List(h)), vbTextCompare)
            Dim rc As New ADODB.Recordset
            If rc.State Then rc.Close
            rc.Open "select DISTINCT(c.bd_jobcharge),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and c.bd_projectkey = '" & nn(0) & "' and j.jobno='" & ju(0) & "' and c.bd_costtype='B' order by c.bd_jobcharge", Cn, 3, 2
            While Not rc.EOF
            List1.AddItem rc(0) & "  -  " & rc(1)
            rc.MoveNext
            Wend
            rc.Close
 End If
 Next h
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
 
End Sub
Private Sub Option3_Click()
If Option3.Value = True Then
Dim g As Integer
g = 0
For g = 0 To List1.ListCount - 1
List1.Selected(g) = False
Next g
 
End If
 
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
