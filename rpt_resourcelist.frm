VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_resourcelist 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7095
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   12515
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
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9480
         Picture         =   "rpt_resourcelist.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to Exit"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9000
         Picture         =   "rpt_resourcelist.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to View"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9840
         Picture         =   "rpt_resourcelist.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to Print"
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "rpt_resourcelist"
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
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Dim fso As New FileSystemObject
Dim fs As Object
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
 RPTHEADING fs
 cnt = 0
Dim sn As Integer
sn = 1
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from resourcemaster where resc_code='" & nm(0) & "' order by resc_code", Cn, 3, 2
While Not rs.EOF
  cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
fs.WriteLine "            <td  >" & rs!resc_code & "</td>"
fs.WriteLine "            <td  >" & rs!resc_desc & "</td>"
fs.WriteLine "            <td  >" & rs!resc_type & "</td>"
fs.WriteLine "            <td  colspan=2>" & rs!resc_vendorcode & "</td>"
fs.WriteLine "            <td  >" & rs!resc_uom & "</td>"
fs.WriteLine "            <td  >" & rs!resc_respcode & "</td>"
fs.WriteLine "        </tr>"
   Dim rs1 As New ADODB.Recordset
   If rs1.State Then rs1.Close
   rs1.Open "select * from resourcedetails where dresc_code='" & rs!resc_code & "' order by dresc_year,dresc_ratetype ", Cn, 3, 2
   While Not rs1.EOF
     cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
   fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
   fs.WriteLine "            <td colspan=2>&nbsp;</td>"
   fs.WriteLine "            <td  >" & rs1!dresc_proj & "</td>"
   fs.WriteLine "            <td  >" & rs1!dresc_year & "</td>"
   fs.WriteLine "            <td  >" & rs1!dresc_curcy & "</td>"
   fs.WriteLine "            <td align=right>" & Format(rs1!dresc_rate, "###,###,##0.00") & "</td>"
   fs.WriteLine "            <td  >" & rs1!dresc_ratetype & "</td>"
   fs.WriteLine "            <td  colspan=2>" & rs1!dresc_notes & "</td>"
   fs.WriteLine "        </tr>"
   rs1.MoveNext
   Wend
   rs.MoveNext
   Wend
 End If
 
 sn = sn + 1
 Next l
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
Unload frmBusy
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "RESOURCE CODE LIST"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
 Dim rc As New ADODB.Recordset
 If rc.State Then rc.Close
 rc.Open "select DISTINCT(resc_code),resc_desc from resourcemaster order by resc_code", Cn, 3, 2
 While Not rc.EOF
 List1.AddItem rc(0) & "  -  " & rc(1)
 rc.MoveNext
 Wend
 rc.Close
                Option1.Value = False
                Option2.Value = True
                Me.Width = 11415
                Me.Height = 9750
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
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
Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=GRAY width=95%>"
    fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
    fs.WriteLine "            <td colspan=2><b>" & GetCompanyName & "</td>"
    fs.WriteLine "           <td COLSPAN=3 ><b>RESOURCE LIST</td>"
    fs.WriteLine "           <td COLSPAN=3>Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
    fs.WriteLine "        </tr>"
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "
   fs.WriteLine "        <tr bgcolor=BLACK height=20 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>S.No.</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Resource Code</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Description</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Resource type</td>"
   fs.WriteLine "            <td Nowrap colspan=2><font color=white>Vendor Code</td>"
   fs.WriteLine "            <td Nowrap><font color=white>UOM</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Responsible Person</td>"
   fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor=BLACK height=15 class=TableFont>"
   fs.WriteLine "            <td colspan=2><font color=white>&nbsp;</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Project</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Year</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Currency</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Rate</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Rate type</td>"
   fs.WriteLine "            <td Nowrap colspan=2><font color=white>Resc Desc</td>"
   fs.WriteLine "        </tr>"
End Sub
