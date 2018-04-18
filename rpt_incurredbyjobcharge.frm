VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_incurredbyjobcharge 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   11250
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6375
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   11175
      ExtentX         =   19711
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1200
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1155
         Left            =   6810
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   4290
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5760
         TabIndex        =   10
         Top             =   960
         Width           =   1040
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   975
         End
      End
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   75
         Top             =   120
         Width           =   5535
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
         TabIndex        =   20
         Top             =   240
         Width           =   1185
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
         TabIndex        =   19
         Top             =   720
         Width           =   930
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
         TabIndex        =   18
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1380
      Width           =   11175
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8640
         Picture         =   "rpt_incurredbyjobcharge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to Save"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10320
         Picture         =   "rpt_incurredbyjobcharge.frx":057F
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Click to Exit"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   7800
         Picture         =   "rpt_incurredbyjobcharge.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Click to View"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9480
         Picture         =   "rpt_incurredbyjobcharge.frx":1199
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Click to Print"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FF8080&
         Caption         =   "Calculate"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6240
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FF8080&
         Caption         =   "L3"
         Height          =   195
         Left            =   5640
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Left            =   4680
         Top             =   240
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FF8080&
         Caption         =   "L2"
         Height          =   195
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TranX Dates"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EAC"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ECTC"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACWP"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy H:mm:ss"
         Format          =   16384003
         CurrentDate     =   38099
      End
   End
End
Attribute VB_Name = "rpt_incurredbyjobcharge"
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

Private Sub Check5_Click()
If Check5.Value = 1 Then
Call appcolor
Else
Call nocolor
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

Private Sub cmd_save_Click()
Load filepatheicjob
End Sub

Private Sub cmd_show_Click()
If cbo_proj.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If


    If Check7.Value = 1 Then
    
    If Check5.Value = 1 Then
    Call appcolor
    Else
    
    Load frmBusy
    frmBusy.Show
    frmBusy.lblBusyString = "Please Wait Report Under Process......"
    If Check8.Value = 1 Then
    Call cuttoffdatechange
    End If
    Call nocolor
    Unload frmBusy
    
    End If
    ElseIf Check6.Value = 1 Then
    Check4.Value = 0
    
    Load frmBusy
    frmBusy.Show
    frmBusy.lblBusyString = "Please Wait Report Under Process......"
   If Check8.Value = 1 Then
    Call cuttoffdatechange
    End If
    Call l2rep
    Unload frmBusy
    
    End If
 
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "EIC BY JOBCHARGE"
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
Me.Width = 11415
Me.Height = 9750

End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub List2_Click()
List1.Clear
 
nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim h As Integer
h = 0
For h = 0 To List2.ListCount - 1
If List2.Selected(h) = True Then
ju = Split(List2.List(h), "  -  ", Len(List2.List(h)), vbTextCompare)
            Dim rc As New ADODB.Recordset
            If rc.State Then rc.Close
            rc.Open "select DISTINCT(c.bd_jobcharge),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and c.bd_projectkey = '" & nn(0) & "' and j.jobno='" & ju(0) & "'  order by c.bd_jobcharge", Cn, 3, 2
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



Public Sub appcolor()
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
   
   
   fs.WriteLine "           <font size=2.5 COLOR= BLUE face=Arial Narrow>" & GetCompanyName & "</font></font><br> "
   fs.WriteLine "        <font COLOR= BLUE size=2>ESTIMATED INCURRED COST BY JOBCHARGE</font>"
   
  
 

fs.WriteLine "    <table border=1 cellspacing=1 bgcolor=blue width=95%>"
fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge</font></td>"

If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>Description</font></td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</font></td>"

Else
fs.WriteLine "            <td colspan=10><font color=white>Description</font></td>"
End If

fs.WriteLine "        </tr>"


   fs.WriteLine "        <tr bgcolor =white height=15 class=TableFont>"
    
   fs.WriteLine "            <td Nowrap><font color=blue>Resc Cde</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>CostCode</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>SprdCde</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>TrnxType</font> </td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=blue>Start Date</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>End Date</font> </td>"
   End If
   fs.WriteLine "            <td Nowrap><font color=blue>Total Qty</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>UOM</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Curcy</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>UnitRate</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>Xrate</font> </td>"
'   fs.WriteLine "            <td Nowrap>DT</td>"
'   fs.WriteLine "            <td Nowrap>Escl</td>"
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=blue>ACWP Amt(RM)</font> </td>"
   End If
   If Check2.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=blue>Tot Qty</font> </td>"
   fs.WriteLine "            <td Nowrap><font color=blue>ECTC Amt(RM)</font> </td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=blue>EAC Amt(RM)</font> </td>"
   End If
   fs.WriteLine "            <td ><font color=blue>Notes</font> </td>"
   fs.WriteLine "        </tr>"
    
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

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

fs.WriteLine "        <tr bgcolor=blue  class=TableFont>"
fs.WriteLine "            <td colspan=2><font color=white>" & nm(0) & "</td>"
 
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=brown>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>" & nm(1) & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>" & nm(1) & "</td>"

Else
fs.WriteLine "            <td colspan=10><font color=white>" & nm(1) & "</td>"
End If

nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim yre As String
Dim fl As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select DISTINCT(bd_resccode) from cost  where bd_jobcharge='" & nm(0) & "' and bd_projectkey ='" & nn(0) & "' and bd_costtype='E' ", Cn, 3, 2

 dtot = 0
ktot = 0
wtot1 = 0
While Not fl.EOF
yre = fl(0)
stot = 0
atot = 0
wtot = 0
                                            Dim fldata1 As New ADODB.Recordset
                                            If fldata1.State Then fldata1.Close
                                            fldata1.Open "select * from cost  where bd_costtype='E' and bd_jobcharge='" & nm(0) & "'   and bd_projectkey ='" & nn(0) & "' and bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2


                                            While Not fldata1.EOF
                                            fs.WriteLine "        <tr bgcolor=white class=TableFont>"
                                             
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_resccode & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_costcode & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_spread & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & fldata1!bd_tranx & "</font> </td>"
                                            If Check4.Value = 1 Then
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & Format(fldata1!bd_sdate, "dd/MM/yyyy") & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap><font color=blue>" & Format(fldata1!bd_edate, "dd/MM/yyyy") & "</font> </td>"
                                            
                                            End If
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_tqty, "###,###,##0.00") & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap  ><font color=blue>" & fldata1!bd_uom & "</td>"
                                            fs.WriteLine "            <td Nowrap ><font color=blue>" & fldata1!bd_curr & "</td>"
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_unitrate, "###,###,##0.00") & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_xchg, "###,###,##0.00") & "</font> </td>"
                                            'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
                                            'fs.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
                                             If Check1.Value = 1 Then
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_extdamt, "###,###,##0.00") & "</font> </td>"
                                            stot = stot + fldata1!bd_extdamt
                                            End If
                                             If Check2.Value = 1 Then
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_e_tqty, "###,###,##0.00") & "</font> </td>"
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "</font> </td>"
                                            atot = atot + fldata1!bd_e_extdamt
                                            End If
                                             If Check3.Value = 1 Then
                                            fs.WriteLine "            <td Nowrap align=right><font color=blue>" & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "</font> </td>"
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

fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
''fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal for the Resource - " & yre & "</font></td>"
Else
fs.WriteLine "            <td  colspan=9><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SubTotal for the Resource- " & yre & "</font></td>"
End If
 If Check1.Value = 1 Then
fs.WriteLine "            <td align=right ><font color=brown>" & Format(stot, "###,###,##0.00") & "</td>"
End If
 If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td align=right ><font color=brown>" & Format(atot, "###,###,##0.00") & "</td>"
End If
 If Check3.Value = 1 Then
fs.WriteLine "            <td align=right ><font color=brown>" & Format(wtot, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
dtot = dtot + stot
ktot = ktot + atot
wtot1 = wtot1 + wtot
fl.MoveNext
Wend



fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
    If Check4.Value = 1 Then
    fs.WriteLine "            <td  colspan=11><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total for the Job - " & List1.List(l) & "</font></td>"
    Else
    fs.WriteLine "            <td  colspan=9><font color=brown>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Total for the Job - " & List1.List(l) & "</font></td>"
    End If
         If Check1.Value = 1 Then
        fs.WriteLine "            <td align=right ><font color=brown>" & Format(dtot, "###,###,##0.00") & "</font></td>"
        End If
             If Check2.Value = 1 Then
            fs.WriteLine "            <td  align=right>&nbsp;</td>"
            fs.WriteLine "            <td align=right ><font color=brown>" & Format(ktot, "###,###,##0.00") & "</font></td>"
            End If
                 If Check3.Value = 1 Then
                fs.WriteLine "            <td align=right ><font color=brown>" & Format(wtot1, "###,###,##0.00") & "</font></td>"
                End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
 tot = tot + dtot
 tot1 = tot1 + ktot
 wtot2 = wtot2 + wtot1
End If

Next l
fs.WriteLine "        <tr bgcolor=yellow height=15 class=TableFont>"
If Check4.Value = 1 Then
fs.WriteLine "            <td  colspan=11>NET TOTAL</td>"
Else
fs.WriteLine "            <td  colspan=9>NET TOTAL</td>"
End If
If Check1.Value = 1 Then
fs.WriteLine "            <td  align=right>" & Format(tot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right>&nbsp;</td>"
fs.WriteLine "            <td  align=right>" & Format(tot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right>" & Format(wtot2, "###,###,##0.00") & "</td>"
End If
fs.WriteLine "            <td align=right >&nbsp;</td>"
fs.WriteLine "        </tr>"
fs.WriteLine " </table>"
    
   
        WebBrowser.Navigate App.Path & "\rep.html"
        fs.WriteLine "    </table><br>"
        fs.WriteLine "    </body>"
        fs.WriteLine "    <html>"


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
   fs.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
   
    


   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

Dim ddtot As Double
Dim ddtot1 As Double
Dim ddwtot2 As Double
  
 Dim cnt As Integer
 RPTHEADING fs
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
 

  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td colspan=17 ><font color=black><b>" & gy(1) & "</td>"
                Else
                fs.WriteLine "            <td colspan=16 ><font color=black><b>" & gy(1) & "</td>"
                End If
                'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
                fs.WriteLine "        </tr>"
        
        
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
                cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
         
         
                    fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                    fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
                     
                    If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=16><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
                    fs.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=15><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
                    fs.WriteLine "            <td colspan=12><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check3.Value = 1 Then
                    fs.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    
                    Else
                    fs.WriteLine "            <td colspan=10><font color=black><font color=black>" & nm(1) & "</td>"
                    End If
                     
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
                            cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                            fs.WriteLine "</table><P></P>"
                            RPTHEADING fs
                            cnt = 0
                            End If
                                                
                                                
                                                fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                                 
                                                fs.WriteLine "            <td Nowrap> " & fldata1!bd_resccode & " </td>"
                                                fs.WriteLine "            <td Nowrap align=center> " & fldata1!bd_costcode & " </td>"
                                                fs.WriteLine "            <td Nowrap align=center> " & fldata1!bd_spread & " </td>"
                                                'fs.WriteLine "            <td Nowrap> " & fldata1!bd_tranx & " </td>"
                                                If Check4.Value = 1 Then
                                                fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_sdate, "dd/MM/yyyy") & " </td>"
                                                fs.WriteLine "            <td Nowrap> " & Format(fldata1!bd_edate, "dd/MM/yyyy") & " </td>"
                                                
                                                End If
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
                            
                            
                   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                            
Dim sttt As String
sttt = Mid(yre & "  -  " & assk, 1, 35)
                            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                            If Check4.Value = 1 Then
                            fs.WriteLine "            <td  colspan=9>  SubTotal   " & sttt & "</td>"
                            Else
                            fs.WriteLine "            <td  colspan=7>  SubTotal   " & sttt & "</td>"
                            End If
                             If Check1.Value = 1 Then
                            fs.WriteLine "            <td align=right ><b> " & Format(stot, "###,###,##0.00") & "</td>"
                            End If
                             If Check2.Value = 1 Then
                            fs.WriteLine "            <td  align=right>&nbsp;</td>"
                            fs.WriteLine "            <td align=right ><b> " & Format(atot, "###,###,##0.00") & "</td>"
                            End If
                             If Check3.Value = 1 Then
                            fs.WriteLine "            <td align=right ><b> " & Format(wtot, "###,###,##0.00") & "</td>"
                            End If
                            fs.WriteLine "            <td align=right >&nbsp;</td>"
                            fs.WriteLine "        </tr>"
                            dtot = dtot + stot
                            ktot = ktot + atot
                            wtot1 = wtot1 + wtot
                            fl.MoveNext
                            Wend
                    
                    
                     cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
Dim stt As String
stt = Mid(List1.List(l), 1, 50)
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    ' fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                        If Check4.Value = 1 Then
                        fs.WriteLine "            <td  colspan=10>  Total     " & stt & " </td>"
                        Else
                        fs.WriteLine "            <td  colspan=8>  Total   " & stt & " </td>"
                        End If
                             If Check1.Value = 1 Then
                            fs.WriteLine "            <td align=right ><b> " & Format(dtot, "###,###,##0.00") & " </td>"
                            End If
                                 If Check2.Value = 1 Then
                                fs.WriteLine "            <td  align=right>&nbsp;</td>"
                                fs.WriteLine "            <td align=right ><b> " & Format(ktot, "###,###,##0.00") & " </td>"
                                End If
                                     If Check3.Value = 1 Then
                                    fs.WriteLine "            <td align=right ><b> " & Format(wtot1, "###,###,##0.00") & " </td>"
                                    End If
                        fs.WriteLine "            <td align=right >&nbsp;</td>"
                        fs.WriteLine "        </tr>"
                        tot = tot + dtot
                        tot1 = tot1 + ktot
                        wtot2 = wtot2 + wtot1
                     
                     
            End If
            End If
            Next l
         cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
                                
Dim sttr As String
sttr = Mid(List2.List(w), 1, 50)
                                
        fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        If Check4.Value = 1 Then
        fs.WriteLine "            <td  colspan=10><b>Total     " & sttr & "</td>"
        Else
        fs.WriteLine "            <td  colspan=8><b> Total     " & sttr & "</td>"
        End If
        If Check1.Value = 1 Then
        fs.WriteLine "            <td  align=right><b> " & Format(tot, "###,###,##0.00") & "</td>"
        End If
        If Check2.Value = 1 Then
        fs.WriteLine "            <td  align=right><b> &nbsp;</td>"
        fs.WriteLine "            <td  align=right><b> " & Format(tot1, "###,###,##0.00") & "</td>"
        End If
        If Check3.Value = 1 Then
        fs.WriteLine "            <td  align=right><b> " & Format(wtot2, "###,###,##0.00") & "</td>"
        End If
        fs.WriteLine "            <td align=right ><b> &nbsp;</td>"
        fs.WriteLine "        </tr>"
        
        ddtot = ddtot + tot
        ddtot1 = ddtot1 + tot1
        ddwtot2 = ddwtot2 + wtot2
   End If
   Next w

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
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddtot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddtot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddwtot2, "###,###,##0.00") & "</td>"
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

Public Sub nocolor1()
Dim fs1 As Object
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
   'fs1.WriteLine "      BORDER-BOTTOM: Wheat 1px solid;"
   'fs1.WriteLine "      BORDER-LEFT: Wheat 1px solid;"
   'fs1.WriteLine "      BORDER-RIGHT: Wheat 1px solid;"
   'fs1.WriteLine "      BORDER-TOP: Wheat 1px solid"
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
   fs1.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
   fs1.WriteLine "<body scroll=auto>"
   fs1.WriteLine "    <center>"
   
    


   'fs1.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "

Dim ddtot As Double
Dim ddtot1 As Double
Dim ddwtot2 As Double
  
 Dim cnt As Integer
 RPTHEADING fs1
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
 

  cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
                fs1.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
                fs1.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
                If Check2.Value = 1 Then
                fs1.WriteLine "            <td colspan=17 ><font color=black><b>" & gy(1) & "</td>"
                Else
                fs1.WriteLine "            <td colspan=16 ><font color=black><b>" & gy(1) & "</td>"
                End If
                'fs1.WriteLine "            <td colspan=9 >&nbsp;</td>"
                fs1.WriteLine "        </tr>"
        
        
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
                cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
         
         
                    fs1.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                    fs1.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
                     
                    If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=16><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
                    fs1.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=15><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check3.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
                    fs1.WriteLine "            <td colspan=12><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check3.Value = 1 Then
                    fs1.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check4.Value = 1 Then
                    fs1.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    
                    Else
                    fs1.WriteLine "            <td colspan=10><font color=black><font color=black>" & nm(1) & "</td>"
                    End If
                     
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
                            cnt = cnt + 1 '********************************
                            If cnt >= 52 Then
                            fs1.WriteLine "</table><P></P>"
                            RPTHEADING fs1
                            cnt = 0
                            End If
                                                
                                                
                                                fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                                                 
                                                fs1.WriteLine "            <td Nowrap> " & fldata1!bd_resccode & " </td>"
                                                fs1.WriteLine "            <td Nowrap align=center> " & fldata1!bd_costcode & " </td>"
                                                fs1.WriteLine "            <td Nowrap align=center> " & fldata1!bd_spread & " </td>"
                                                'fs1.WriteLine "            <td Nowrap> " & fldata1!bd_tranx & " </td>"
                                                If Check4.Value = 1 Then
                                                fs1.WriteLine "            <td Nowrap> " & Format(fldata1!bd_sdate, "dd/MM/yyyy") & " </td>"
                                                fs1.WriteLine "            <td Nowrap> " & Format(fldata1!bd_edate, "dd/MM/yyyy") & " </td>"
                                                
                                                End If
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_tqty, "###,###,##0.00") & " </td>"
                                                fs1.WriteLine "            <td Nowrap align=center > " & fldata1!bd_uom & "</td>"
                                                fs1.WriteLine "            <td Nowrap align=center> " & fldata1!bd_curr & "</td>"
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_unitrate, "###,###,##0.00") & "  </td>"
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_xchg, "###,###,##0.00") & " </td>"
                                                'fs1.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_downtime, "###,###,##0.00") & "</td>"
                                                'fs1.WriteLine "            <td Nowrap align=right>" & Format(fldata1!bd_escl, "###,###,##0.00") & "</td>"
                                                 If Check1.Value = 1 Then
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_extdamt, "###,###,##0.00") & " </td>"
                                                stot = stot + fldata1!bd_extdamt
                                                End If
                                                 If Check2.Value = 1 Then
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_tqty, "###,###,##0.00") & " </td>"
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format(fldata1!bd_e_extdamt, "###,###,##0.00") & "  </td>"
                                                atot = atot + fldata1!bd_e_extdamt
                                                End If
                                                 If Check3.Value = 1 Then
                                                fs1.WriteLine "            <td Nowrap align=right> " & Format((fldata1!bd_extdamt) + (fldata1!bd_e_extdamt), "###,###,##0.00") & "  </td>"
                                                wtot = wtot + (fldata1!bd_extdamt) + (fldata1!bd_e_extdamt)
                                                End If
               If fldata1!bd_notes <> "" Then
                                Dim jh11 As String
                                jh11 = Mid(fldata1!bd_notes, 1, 15)
                                fs1.WriteLine "            <td ><b> " & jh11 & "</td>"
                                Else
                                Dim cd1 As New ADODB.Recordset
                                If cd1.State Then cd1.Close
                                cd1.Open "select cc_desc from costcode where cc_code='" & fldata1!bd_costcode & "'", Cn, 3, 2
                                If Not cd1.EOF Then
                                Dim jh1 As String
                                jh1 = Mid(cd1(0), 1, 15)
                                fs1.WriteLine "            <td Nowrap> " & jh1 & "</td>"
                                End If
                                End If
                                                fs1.WriteLine "       </tr>"
                                                fldata1.MoveNext
                                                Wend
                                                
                                                
                                                
                                                
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = rscd(0)
End If
                            
                            
                   cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
                            
Dim sttt As String
sttt = Mid(yre & "  -  " & assk, 1, 35)
                            fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                            fs1.WriteLine "            <td  colspan=1>&nbsp;</td>"
                            If Check4.Value = 1 Then
                            fs1.WriteLine "            <td  colspan=9>  SubTotal   " & sttt & "</td>"
                            Else
                            fs1.WriteLine "            <td  colspan=7>  SubTotal   " & sttt & "</td>"
                            End If
                             If Check1.Value = 1 Then
                            fs1.WriteLine "            <td align=right ><b> " & Format(stot, "###,###,##0.00") & "</td>"
                            End If
                             If Check2.Value = 1 Then
                            fs1.WriteLine "            <td  align=right>&nbsp;</td>"
                            fs1.WriteLine "            <td align=right ><b> " & Format(atot, "###,###,##0.00") & "</td>"
                            End If
                             If Check3.Value = 1 Then
                            fs1.WriteLine "            <td align=right ><b> " & Format(wtot, "###,###,##0.00") & "</td>"
                            End If
                            fs1.WriteLine "            <td align=right >&nbsp;</td>"
                            fs1.WriteLine "        </tr>"
                            dtot = dtot + stot
                            ktot = ktot + atot
                            wtot1 = wtot1 + wtot
                            fl.MoveNext
                            Wend
                    
                    
                     cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
                                
Dim stt As String
stt = Mid(List1.List(l), 1, 50)
                    fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                    ' fs1.WriteLine "            <td  colspan=1>&nbsp;</td>"
                        If Check4.Value = 1 Then
                        fs1.WriteLine "            <td  colspan=10>  Total     " & stt & " </td>"
                        Else
                        fs1.WriteLine "            <td  colspan=8>  Total   " & stt & " </td>"
                        End If
                             If Check1.Value = 1 Then
                            fs1.WriteLine "            <td align=right ><b> " & Format(dtot, "###,###,##0.00") & " </td>"
                            End If
                                 If Check2.Value = 1 Then
                                fs1.WriteLine "            <td  align=right>&nbsp;</td>"
                                fs1.WriteLine "            <td align=right ><b> " & Format(ktot, "###,###,##0.00") & " </td>"
                                End If
                                     If Check3.Value = 1 Then
                                    fs1.WriteLine "            <td align=right ><b> " & Format(wtot1, "###,###,##0.00") & " </td>"
                                    End If
                        fs1.WriteLine "            <td align=right >&nbsp;</td>"
                        fs1.WriteLine "        </tr>"
                        tot = tot + dtot
                        tot1 = tot1 + ktot
                        wtot2 = wtot2 + wtot1
                     
                     
            End If
            End If
            Next l
         cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
                                
Dim sttr As String
sttr = Mid(List2.List(w), 1, 50)
                                
        fs1.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        If Check4.Value = 1 Then
        fs1.WriteLine "            <td  colspan=10><b>Total     " & sttr & "</td>"
        Else
        fs1.WriteLine "            <td  colspan=8><b> Total     " & sttr & "</td>"
        End If
        If Check1.Value = 1 Then
        fs1.WriteLine "            <td  align=right><b> " & Format(tot, "###,###,##0.00") & "</td>"
        End If
        If Check2.Value = 1 Then
        fs1.WriteLine "            <td  align=right><b> &nbsp;</td>"
        fs1.WriteLine "            <td  align=right><b> " & Format(tot1, "###,###,##0.00") & "</td>"
        End If
        If Check3.Value = 1 Then
        fs1.WriteLine "            <td  align=right><b> " & Format(wtot2, "###,###,##0.00") & "</td>"
        End If
        fs1.WriteLine "            <td align=right ><b> &nbsp;</td>"
        fs1.WriteLine "        </tr>"
        
        ddtot = ddtot + tot
        ddtot1 = ddtot1 + tot1
        ddwtot2 = ddwtot2 + wtot2
   End If
   Next w

 cnt = cnt + 1 '********************************
                                If cnt >= 52 Then
                                fs1.WriteLine "</table><P></P>"
                                RPTHEADING fs1
                                cnt = 0
                                End If
fs1.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
If Check4.Value = 1 Then
fs1.WriteLine "            <td  colspan=10><font color=white>REPORT TOTAL</td>"
Else
fs1.WriteLine "            <td  colspan=8><font color=white>REPORT TOTAL</td>"
End If
If Check1.Value = 1 Then
fs1.WriteLine "            <td  align=right><font color=white>" & Format(ddtot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
fs1.WriteLine "            <td  align=right><font color=white>&nbsp;</td>"
fs1.WriteLine "            <td  align=right><font color=white>" & Format(ddtot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs1.WriteLine "            <td  align=right><font color=white>" & Format(ddwtot2, "###,###,##0.00") & "</td>"
End If
fs1.WriteLine "            <td align=right ><font color=white>&nbsp;</td>"
fs1.WriteLine "        </tr>"
fs1.WriteLine " </table>"
   fs1.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs1.WriteLine "           <br></br> <td ><b> JobNo.</td>"
For f = 0 To List2.ListCount - 1
If List2.Selected(f) = True Then
hh = Split(List2.List(f), "  -  ", Len(List2.List(f)), vbTextCompare)
fs1.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs1.WriteLine "            <td > " & List2.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs1.WriteLine "            <td > <b>JobCharge</td>"
For r = 0 To List1.ListCount - 1
If List1.Selected(r) = True Then
hh = Split(List1.List(r), "  -  ", Len(List1.List(r)), vbTextCompare)
 fs1.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs1.WriteLine "            <td > " & List1.List(r) & "</td></tr>"
End If
Next r
 
fs1.WriteLine " </table>"
   
   WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
   fs1.WriteLine "    </table><br>"
  
   fs1.WriteLine "    </body>"
   fs1.WriteLine "    <html>"


  
    

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
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"

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
fs.WriteLine "            <td colspan=9>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L3)</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L3)</td>"

Else
fs.WriteLine "            <td colspan=4>INCURRED BY JOBCHARGE(L3)</td>"
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
 
 
 
fs.WriteLine "     <tr bgcolor=black  height=20 class=TableFont>"
fs.WriteLine "     <td Nowrap colspan=2><font color=white>JobCharge</td>"

If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=16><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=15><font color=white>Description</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</td>"
Else
fs.WriteLine "            <td colspan=10><font color=white>Description</td>"
End If

   fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap ><font color=white> RescCde  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> CostCde  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> SprdCde </td>"
   If Check4.Value = 1 Then
   fs.WriteLine "            <td Nowrap><font color=white> StartDate  </td>"
   fs.WriteLine "            <td Nowrap><font color=white> EndDate  </td>"
   End If
   fs.WriteLine "            <td Nowrap align=right><font color=white> TotalQty  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> UOM  </td>"
   fs.WriteLine "            <td Nowrap align=center><font color=white> Curcy  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white> UnitRate  </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white> xRate </td>"
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>ACWP Amt(RM)</font> </td>"
   End If
   If Check2.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>TotQty </td>"
   fs.WriteLine "            <td Nowrap align=right><font color=white>ECTC Amt(RM)  </td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>EAC Amt(RM)  </td>"
   End If
   fs.WriteLine "            <td ><font color=white>Notes/CostCde Desc </td>"
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
fldata.Open "select * from cost where bd_jobcharge='" & xk(0) & "' and bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2


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
cd.Open "select * from cost where  bd_jobcharge='" & xk(0) & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
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

Public Sub l2rep()
Call cuttoffdatechange
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
   fs.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
   fs.WriteLine "<body scroll=auto>"
   fs.WriteLine "    <center>"
   
   
Dim cnt As Integer
RPTHEADINGL2 fs
cnt = 0
            
                 

Dim ddtot As Double
Dim ddtot1 As Double
Dim ddwtot2 As Double
  
 
  
 nn = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
 Dim w As Integer
 w = 0
    ddtot = 0
    ddtot1 = 0
    ddwtot2 = 0
 For w = 0 To List2.ListCount - 1
 If List2.Selected(w) = True Then
 gy = Split(List2.List(w), "  -  ", Len(List2.List(w)), vbTextCompare)
 

   cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
                fs.WriteLine "        <tr bgcolor=#aeaeae  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2><font color=black><b>" & gy(0) & "</td>"
                If Check2.Value = 1 Then
                fs.WriteLine "            <td colspan=15 ><font color=black><b>" & gy(1) & "</td>"
                Else
                fs.WriteLine "            <td colspan=14 ><font color=black><b>" & gy(1) & "</td>"
                End If
                'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
                fs.WriteLine "        </tr>"
        
        
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
 
         
         
   cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
                    fs.WriteLine "        <tr bgcolor=#acacac  height=15 class=TableFont>"
                    fs.WriteLine "            <td colspan=2><font color=black>" & nm(0) & "</td>"
                     
                    If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=14><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
                    fs.WriteLine "            <td colspan=12><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=12><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=13><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check2.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=11><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check3.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=12><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
                    fs.WriteLine "            <td colspan=10><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check3.Value = 1 Then
                    fs.WriteLine "            <td colspan=11><font color=black><font color=black>" & nm(1) & "</td>"
                    ElseIf Check1.Value = 1 And Check4.Value = 1 Then
                    fs.WriteLine "            <td colspan=11><font color=black><font color=black>" & nm(1) & "</td>"
                    
                    Else
                    fs.WriteLine "            <td colspan=8><font color=black><font color=black>" & nm(1) & "</td>"
                    End If
                     
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
                                                fldata1.Open "select * from cost c,jobcharge j where c.bd_jobcharge=j.job_code and j.jobno='" & gy(0) & "' and c.bd_costtype='E' and j.job_code='" & nm(0) & "'  and j.job_desc='" & nm(1) & "' and c.bd_projectkey ='" & nn(0) & "' and c.bd_resccode='" & yre & "' order by bd_resccode", Cn, 3, 2
                                               
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
                                                
                                                
                                                
                                                
Dim assk As String
Dim rscd As New ADODB.Recordset
If rscd.State Then rscd.Close
rscd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & yre & "'", Cn, 3, 2
If Not rscd.EOF Then
assk = rscd(0)
End If
                            
 
                            
   cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
                            
                            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                            fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                            If Check4.Value = 1 Then
                            fs.WriteLine "            <td  colspan=9>  SubTotal   " & yre & "  -  " & assk & "</td>"
                            Else
                            fs.WriteLine "            <td  colspan=7>  SubTotal   " & yre & "  -  " & assk & "</td>"
                            End If
                             If Check1.Value = 1 Then
                            fs.WriteLine "            <td align=right >  " & Format(stot, "###,###,##0.00") & "</td>"
                            End If
                             If Check2.Value = 1 Then
                            'fs.WriteLine "            <td  align=right>&nbsp;</td>"
                            fs.WriteLine "            <td align=right >  " & Format(atot, "###,###,##0.00") & "</td>"
                            End If
                             If Check3.Value = 1 Then
                            fs.WriteLine "            <td align=right >  " & Format(wtot, "###,###,##0.00") & "</td>"
                            End If
                            
                            fs.WriteLine "        </tr>"
                            dtot = dtot + stot
                            ktot = ktot + atot
                            wtot1 = wtot1 + wtot
                            fl.MoveNext
                            Wend
                    
                    
   cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
      
                    fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
                     'fs.WriteLine "            <td  colspan=1>&nbsp;</td>"
                        If Check4.Value = 1 Then
                        fs.WriteLine "            <td  colspan=10>  Total     " & List1.List(l) & " </td>"
                        Else
                        fs.WriteLine "            <td  colspan=8>  Total   " & List1.List(l) & " </td>"
                        End If
                             If Check1.Value = 1 Then
                            fs.WriteLine "            <td align=right ><b> " & Format(dtot, "###,###,##0.00") & " </td>"
                            End If
                                 If Check2.Value = 1 Then
                                'fs.WriteLine "            <td  align=right>&nbsp;</td>"
                                fs.WriteLine "            <td align=right ><b> " & Format(ktot, "###,###,##0.00") & " </td>"
                                End If
                                     If Check3.Value = 1 Then
                                    fs.WriteLine "            <td align=right ><b> " & Format(wtot1, "###,###,##0.00") & " </td>"
                                    End If
                   
                    fs.WriteLine "        </tr>"
                     tot = tot + dtot
                     tot1 = tot1 + ktot
                     wtot2 = wtot2 + wtot1
                     
                     
            End If
            End If
            Next l
     
     
     
   cnt = cnt + 1 '********************************
                                If cnt >= 53 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADINGL2 fs
                                cnt = 0
                                End If
        fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        If Check4.Value = 1 Then
        fs.WriteLine "            <td  colspan=10><b>Total     " & List2.List(w) & "</td>"
        Else
        fs.WriteLine "            <td  colspan=8><b> Total     " & List2.List(w) & "</td>"
        End If
        If Check1.Value = 1 Then
        fs.WriteLine "            <td  align=right><b> " & Format(tot, "###,###,##0.00") & "</td>"
        End If
        If Check2.Value = 1 Then
        'fs.WriteLine "            <td  align=right><b> &nbsp;</td>"
        fs.WriteLine "            <td  align=right><b> " & Format(tot1, "###,###,##0.00") & "</td>"
        End If
        If Check3.Value = 1 Then
        fs.WriteLine "            <td  align=right><b> " & Format(wtot2, "###,###,##0.00") & "</td>"
        End If
         
        fs.WriteLine "        </tr>"
        
        ddtot = ddtot + tot
        ddtot1 = ddtot1 + tot1
        ddwtot2 = ddwtot2 + wtot2
   End If
   Next w

 
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
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddtot, "###,###,##0.00") & "</td>"
End If
If Check2.Value = 1 Then
'fs.WriteLine "            <td  align=right><font color=white>&nbsp;</td>"
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddtot1, "###,###,##0.00") & "</td>"
End If
If Check3.Value = 1 Then
fs.WriteLine "            <td  align=right><font color=white>" & Format(ddwtot2, "###,###,##0.00") & "</td>"
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

Public Sub RPTHEADINGL2(fs As Object)

   fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
 
ff = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
            
fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=9>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>" & GetCompanyName & "</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=4>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=4>" & GetCompanyName & "</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>" & GetCompanyName & "</td>"

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
                
 fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=9>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=5>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=8>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=7>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=4>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=4>INCURRED BY JOBCHARGE(L2)</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=6>INCURRED BY JOBCHARGE(L2)</td>"

Else
fs.WriteLine "            <td colspan=3>INCURRED BY JOBCHARGE(L2)</td>"
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



fs.WriteLine "     <tr bgcolor=black  height=15 class=TableFont>"
fs.WriteLine "     <td Nowrap colspan=2><font color=white>JobNo.</td>"
Check4.Value = 0
If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=14><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</td>"
ElseIf Check2.Value = 1 And Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=13><font color=white>Description</td>"
ElseIf Check2.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=11><font color=white>Description</td>"
ElseIf Check3.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=12><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check2.Value = 1 Then
fs.WriteLine "            <td colspan=10><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check3.Value = 1 Then
fs.WriteLine "            <td colspan=11><font color=white>Description</td>"
ElseIf Check1.Value = 1 And Check4.Value = 1 Then
fs.WriteLine "            <td colspan=11><font color=white>Description</td>"

Else
fs.WriteLine "            <td colspan=8><font color=white>Description</td>"
End If

'fs.WriteLine "            <td colspan=9 >&nbsp;</td>"
fs.WriteLine "        </tr>"
 
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
    
   fs.WriteLine "            <td Nowrap colspan=2><font color=white>JobCharge</td>"
   fs.WriteLine "            <td Nowrap colspan=6><font color=white>Description</td>"
   
   If Check1.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>ACWP Amt(RM)</font> </td>"
   End If
   If Check2.Value = 1 Then
 
   fs.WriteLine "            <td Nowrap align=right><font color=white>ECTC Amt(RM)  </td>"
   End If
   If Check3.Value = 1 Then
   fs.WriteLine "            <td Nowrap align=right><font color=white>EAC Amt(RM)  </td>"
   End If
  
   fs.WriteLine "        </tr>"

End Sub


