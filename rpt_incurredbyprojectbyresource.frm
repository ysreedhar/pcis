VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_incurredbyprojectbyresource 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6855
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   11055
      ExtentX         =   19500
      ExtentY         =   12091
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
      TabIndex        =   7
      Top             =   0
      Width           =   11655
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   705
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   720
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   1045
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6690
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   240
         Width           =   4290
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1205
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   9
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
         TabIndex        =   19
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Resource"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   720
         Width           =   825
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5640
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
         TabIndex        =   17
         Top             =   240
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10320
         Picture         =   "rpt_incurredbyprojectbyresource.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8640
         Picture         =   "rpt_incurredbyprojectbyresource.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9480
         Picture         =   "rpt_incurredbyprojectbyresource.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FF8080&
         Caption         =   "Calculate"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6600
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FF8080&
         Caption         =   "L3"
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FF8080&
         Caption         =   "L2"
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACWP"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ECTC"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EAC"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Transaction Dates"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   48758787
         CurrentDate     =   38099
      End
   End
End
Attribute VB_Name = "rpt_incurredbyprojectbyresource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rg As New ADODB.Recordset
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
       
            
            Dim rc As New ADODB.Recordset
            If rc.State Then rc.Close
            rc.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j where c.bd_jobcharge=j.job_code  and  bd_costtype='E' and c.bd_projectkey='" & spp(0) & "' order by c.bd_resccode", Cn, 3, 2
            While Not rc.EOF
            Dim rcd As New ADODB.Recordset
            If rcd.State Then rcd.Close
            rcd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rc(0) & "' ", Cn, 3, 2
                   If Not rcd.EOF Then
                   List1.AddItem rc(0) & "  -  " & rcd(0)
                   Else
                   List1.AddItem rc(0)
                   End If
            
            rc.MoveNext
            Wend
            rc.Close

  
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


If Check6.Value = 1 Then
If Check8.Value = 1 Then
Call cuttoffdatechange
End If
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy

ElseIf Check5.Value = 1 Then
Check4.Value = 0
If Check8.Value = 1 Then
Call cuttoffdatechange
End If
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call l2rep
Unload frmBusy

End If
End Sub
 
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "EIC BY PROJECT BY RESOURCE"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pk.EOF
cbo_proj.AddItem pk(0) & "  -  " & pk(1)
pk.MoveNext
Wend
pk.Close
Check1.Value = 1
DTPicker1.Value = main.DTPcutdate1.Value
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
            
    Me.Width = 11415
    Me.Height = 9750

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
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
  
            Dim cnt As Integer
            RPTHEADING fs
            cnt = 0
      
Dim stot As Double
Dim tot As Double
Dim dtot As Double
stot = 0: tot = 0: dtot = 0
nl = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)

If rg.State Then rg.Close
rg.Open "select * from resourcemaster r,resourcedetails d where r.resc_code=d.dresc_code and r.resc_code='" & nm(0) & "' and dresc_proj='" & nl(0) & "' order by r.resc_code", Cn, 3, 2
If Not rg.EOF Then

cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=7 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=4 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
Else
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=3 ><font color=black>" & rg!resc_vendorcode & "</td>"
End If
 
fs.WriteLine "        </tr>"
End If

 
dtot = 0
ktot = 0
wtot1 = 0
 


Dim Y As Integer
Y = 0
For Y = 0 To List2.ListCount - 1
If List2.Selected(Y) = True Then
fl = Split(List2.List(Y), "  -  ", Len(List2.List(Y)), vbTextCompare)
stot = 0
atot = 0
wtot = 0

                                    Dim fldata1 As New ADODB.Recordset
                                    If fldata1.State Then fldata1.Close
                                    fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and c.bd_costtype='E' and c.bd_resccode='" & nm(0) & "' and c.bd_projectkey='" & nl(0) & "' and  j.jobno='" & fl(0) & "' order by j.jobno,c.bd_jobcharge", Cn, 3, 2
                                    stot = 0
                                    While Not fldata1.EOF
                                    
                                    
                                    cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                    
                                    fs.WriteLine "            <td Nowrap>" & fldata1!bd_JobCharge & "</td>"
                                    fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_costcode & "</td>"
                                    fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_spread & "</td>"
                                   
                                    If Check4.Value = 1 Then
                                    fs.WriteLine "            <td Nowrap>" & Format(fldata1!bd_sdate, "dd/MM/yyyy") & "</td>"
                                    fs.WriteLine "            <td Nowrap>" & Format(fldata1!bd_edate, "dd/MM/yyyy") & "</td>"
                                    
                                    End If
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</td>"
                                    fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_uom & "</td>"
                                    fs.WriteLine "            <td Nowrap align=center>" & fldata1!bd_curr & "</td>"
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</td>"
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</td>"
                                     
                                    If Check1.Value = 1 Then
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</td>"
                                    stot = stot + fldata1!bd_extdamt
                                    End If
                                     If Check2.Value = 1 Then
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_e_tqty, "###,###,##0.00") & "</td>"
                                    fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "</td>"
                                    atot = atot + fldata1!bd_e_extdamt
                                    End If
                                     If Check3.Value = 1 Then
                                    fs.WriteLine "            <td Nowrap align=right>" & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "</td>"
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
                                    fs.WriteLine "        </tr>"
                                    fldata1.MoveNext
                                    Wend
                                    
                                    
 cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
Dim sttt As String
sttt = Mid(List2.List(Y), 1, 35)

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=9><b>SubTotal  - " & sttt & "</td>"
Else
fs.WriteLine "            <td  colspan=7><b>SubTotal  - " & sttt & "</td>"
End If
 If Check1.Value = 1 Then
fs.WriteLine "            <td align=right ><b>" & Format(stot, "###,###,##0.00") & "</td>"
End If
 If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right ><b>" & Format(atot, "###,###,##0.00") & "</td>"
End If
 If Check3.Value = 1 Then
fs.WriteLine "            <td align=right ><b>" & Format(wtot, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
dtot = dtot + stot
ktot = ktot + atot
wtot1 = wtot1 + wtot
End If
Next Y

cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
     Dim stttt As String
stttt = Mid(List1.List(l), 1, 40)

                                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                    fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                                                        If Check4.Value = 1 Then
                                                        fs.WriteLine "            <td  colspan=9><b>Total - " & stttt & "</td>"
                                                        Else
                                                        fs.WriteLine "            <td  colspan=7><b>Total - " & stttt & "</td>"
                                                        End If
                                                        If Check1.Value = 1 Then
                                                        fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
                                                        End If
                                                        If Check2.Value = 1 Then
                                                        fs.WriteLine "            <td  align=right>&nbsp;</td>"
                                                        fs.WriteLine "            <td align=right ><b>" & Format(ktot, "###,###,##0.00") & "</td>"
                                                        End If
                                                        If Check3.Value = 1 Then
                                                        fs.WriteLine "            <td align=right ><b>" & Format(wtot1, "###,###,##0.00") & "</td>"
                                                        End If
                                    fs.WriteLine "            <td align=right >&nbsp;</td>"
                                    fs.WriteLine "        </tr>"
dtot2 = dtot2 + dtot
ktot2 = ktot2 + ktot
wtot2 = wtot2 + wtot1

End If
Next l


cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
Else
fs.WriteLine "            <td  colspan=8><font color=white>REPORT TOTAL</td>"
End If
If Check1.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(dtot2, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(ktot2, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(wtot2, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
fs.WriteLine "        </tr>"
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
fs.WriteLine "            <td > <b>Resource</td>"
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
Public Sub RPTHEADING(fs As Object)

fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
 
ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
fs.WriteLine "        <tr bgcolor=white  height=25 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=9>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"

ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td colspan=5>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=10>" & GetCompanyName & "</td>"
Else
fs.WriteLine "            <td colspan=4>" & GetCompanyName & "</td>"
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
fs.WriteLine "            <td colspan=9>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td colspan=5>INCURRED BY RESOURCE(L3)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=10>INCURRED BY RESOURCE(L3)</td>"
Else
fs.WriteLine "            <td colspan=4>INCURRED BY RESOURCE(L3)</td>"
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
fs.WriteLine "            <td colspan=18><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=17><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>&nbsp;</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>&nbsp;</td>"
Else
fs.WriteLine "            <td colspan=12><font color=white>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"

 
fs.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=7 ><font color=white>Vendor Name</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor Name</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor Name</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor Name</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor Name</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor Name</td>"
Else
fs.WriteLine "            <td Nowrap><font color=white>Resc Cde</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Resource Code Description</td>"
fs.WriteLine "            <td Nowrap ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=3 ><font color=white>Vendor Name</td>"
End If
fs.WriteLine "        </tr>"

   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap ><font color=white>JobCharge</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>CostCde</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>SprdCde</td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap ><font color=white>StartDate</td>"
   fs.WriteLine "            <td Nowrap  ><font color=white>EndDate</td>"
   End If
   fs.WriteLine "            <td Nowrap align=right><font color=white>TotalQty</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>UOM</td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white>Curcy</td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>UnitRate</td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>xRate</td>"
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>ACWP Amt(RM)</td>"
   End If
   If Check2.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>TotQty</td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>ECTC Amt(RM)</td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>EAC Amt(RM)</td>"
   End If
   fs.WriteLine "            <td ><font color=white>Notes/CostCdeDesc</td>"
   fs.WriteLine "        </tr>"

End Sub

Public Sub cuttoffdatechange()
Dim j As Integer
j = 0
For j = 0 To List1.ListCount - 1
If List1.Selected(j) = True Then
xk = Split(List1.List(j), "  -  ", Len(List1.List(j)), vbTextCompare)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where bd_resccode='" & xk(0) & "' and bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2


    While Not fldata.EOF

     iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_JobCharge, "  -  ", Len(fldata!bd_JobCharge), vbTextCompare)


Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If

Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
fldata2.Open "select * from cost where    bd_jobcharge='" & fldata!bd_JobCharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

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

End If
Next j


Dim jj As Integer
jj = 0
For jj = 0 To List1.ListCount - 1
If List1.Selected(jj) = True Then
xk = Split(List1.List(jj), "  -  ", Len(List1.List(jj)), vbTextCompare)
Dim cid As Double
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where  bd_resccode='" & xk(0) & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
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
End If
Next jj

End Sub

Public Sub l2rep()
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
 cnt = 0
RPTHEADINGL2 fs

Dim stot As Double
Dim tot As Double
Dim dtot As Double
stot = 0: tot = 0: dtot = 0
nl = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
 nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)

If rg.State Then rg.Close
rg.Open "select * from resourcemaster r,resourcedetails d where r.resc_code=d.dresc_code and r.resc_code='" & nm(0) & "' and dresc_proj='" & nl(0) & "' order by r.resc_code", Cn, 3, 2
If Not rg.EOF Then
cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=7 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=4 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=4 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=4 ><font color=black>" & rg!resc_vendorcode & "</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=5 ><font color=black>" & rg!resc_vendorcode & "</td>"
Else
fs.WriteLine "            <td ><font color=black>" & nm(0) & "</td>"
fs.WriteLine "            <td colspan=6 ><font color=black>" & nm(1) & "</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=black>" & kj(0) & "</td>"
fs.WriteLine "            <td colspan=2 ><font color=black>" & rg!resc_vendorcode & "</td>"
End If
 
fs.WriteLine "        </tr>"
End If

 
dtot = 0
ktot = 0
wtot1 = 0
 


Dim Y As Integer
Y = 0
For Y = 0 To List2.ListCount - 1
If List2.Selected(Y) = True Then
fl = Split(List2.List(Y), "  -  ", Len(List2.List(Y)), vbTextCompare)
stot = 0
atot = 0
wtot = 0

                                    Dim fldata1 As New ADODB.Recordset
                                    If fldata1.State Then fldata1.Close
                                    fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and c.bd_costtype='E' and c.bd_resccode='" & nm(0) & "' and c.bd_projectkey='" & nl(0) & "' and  j.jobno='" & fl(0) & "' order by j.jobno,c.bd_jobcharge", Cn, 3, 2
                                    stot = 0
                                    While Not fldata1.EOF
                                    
                                    
                                   
                                     
                                    If Check1.Value = 1 Then
                                    stot = stot + fldata1!bd_extdamt
                                    End If
                                    If Check2.Value = 1 Then
                                    atot = atot + fldata1!bd_e_extdamt
                                    End If
                                    If Check3.Value = 1 Then
                                    wtot = wtot + (fldata1!bd_extdamt) + (fldata1!bd_e_extdamt)
                                    End If
                                    
                                    fldata1.MoveNext
                                    Wend
                                    
                                    
 
cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=9>SubTotal  - " & List2.List(Y) & "</td>"
Else
fs.WriteLine "            <td  colspan=7>SubTotal  - " & List2.List(Y) & "</td>"
End If
 If Check1.Value = 1 Then
fs.WriteLine "            <td align=right > " & Format(stot, "###,###,##0.00") & "</td>"
End If
 If Check2.Value = 1 Then

fs.WriteLine "            <td align=right > " & Format(atot, "###,###,##0.00") & "</td>"
End If
 If Check3.Value = 1 Then
fs.WriteLine "            <td align=right > " & Format(wtot, "###,###,##0.00") & "</td>"
End If

fs.WriteLine "        </tr>"
dtot = dtot + stot
ktot = ktot + atot
wtot1 = wtot1 + wtot
End If
Next Y
cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If

                                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                    fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                                                        If Check4.Value = 1 Then
                                                        fs.WriteLine "            <td  colspan=9>Total - " & List1.List(l) & "</td>"
                                                        Else
                                                        fs.WriteLine "            <td  colspan=7>Total - " & List1.List(l) & "</td>"
                                                        End If
                                                        If Check1.Value = 1 Then
                                                        fs.WriteLine "            <td align=right ><b>" & Format(dtot, "###,###,##0.00") & "</td>"
                                                        End If
                                                        If Check2.Value = 1 Then
                                                     
                                                        fs.WriteLine "            <td align=right ><b>" & Format(ktot, "###,###,##0.00") & "</td>"
                                                        End If
                                                        If Check3.Value = 1 Then
                                                        fs.WriteLine "            <td align=right ><b>" & Format(wtot1, "###,###,##0.00") & "</td>"
                                                        End If
                                 
                                    fs.WriteLine "        </tr>"
dtot2 = dtot2 + dtot
ktot2 = ktot2 + ktot
wtot2 = wtot2 + wtot1

End If
Next l


 cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
Else
fs.WriteLine "            <td  colspan=8><font color=white>REPORT TOTAL</td>"
End If
If Check1.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(dtot2, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then

fs.WriteLine "            <td  align=right><font color=white>" & Format(ktot2, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(wtot2, "###,###,##0.00") & "</td>"
End If

fs.WriteLine "        </tr>"
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
fs.WriteLine "            <td > <b>Resource</td>"
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
Public Sub RPTHEADINGL2(fs As Object)

fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
 
ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
fs.WriteLine "        <tr bgcolor=white  height=25 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=9>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=5>" & GetCompanyName & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"

ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=4>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=3>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=3>" & GetCompanyName & "</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td colspan=3>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=10>" & GetCompanyName & "</td>"
Else
fs.WriteLine "            <td colspan=3>" & GetCompanyName & "</td>"
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
fs.WriteLine "            <td colspan=9>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=4>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=3>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=4>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td colspan=5>INCURRED BY RESOURCE(L2)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=10>INCURRED BY RESOURCE(L2)</td>"
Else
fs.WriteLine "            <td colspan=3>INCURRED BY RESOURCE(L2)</td>"
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
fs.WriteLine "            <td colspan=11><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=11><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=17><font color=white>&nbsp;</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>&nbsp;</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=10><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=10><font color=white>&nbsp;</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>&nbsp;</td>"
Else
fs.WriteLine "            <td colspan=9><font color=white>&nbsp;</td>"
End If
fs.WriteLine "        </tr>"





fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=7 ><font color=white>Vendor</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Vendor</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
kj = Split(rg!resc_type, "  -  ", Len(rg!resc_type), vbTextCompare)
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=4 ><font color=white>Vendor</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=4 ><font color=white>Vendor</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=4 ><font color=white>Vendor</td>"
ElseIf Check4.Value = 1 Then
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=5 ><font color=white>Vendor</td>"
Else
fs.WriteLine "            <td ><font color=white>Resource</td>"
fs.WriteLine "            <td colspan=6 ><font color=white>Description</td>"
fs.WriteLine "            <td  ><font color=white>Resc Type</td>"
fs.WriteLine "            <td colspan=2 ><font color=white>Vendor</td>"
End If
 
fs.WriteLine "        </tr>"

 
        fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
        'fs.WriteLine "            <td Nowrap colspan=2><font color=white>&nbsp;</td>"
        fs.WriteLine "            <td Nowrap colspan=8><font color=white>&nbsp;</td>"
        If Check1.Value = 1 Then
        fs.WriteLine "            <td Nowrap align=right><font color=white>ACWP Amt(RM)</font> </td>"
        End If
        If Check2.Value = 1 Then
        fs.WriteLine "            <td Nowrap align=right><font color=white>ECTC Amt(RM)</td>"
        End If
        If Check3.Value = 1 Then
        fs.WriteLine "            <td Nowrap align=right><font color=white>EAC Amt(RM)</td>"
        End If
        fs.WriteLine "        </tr>"


End Sub
