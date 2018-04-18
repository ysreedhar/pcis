VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_currencylist 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7095
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   10695
      ExtentX         =   18865
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
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9120
         Picture         =   "rpt_currencylist.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Click to Print"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8280
         Picture         =   "rpt_currencylist.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to View"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9960
         Picture         =   "rpt_currencylist.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to Exit"
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DC7E5A&
         Caption         =   "Apply Color"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   4335
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
Attribute VB_Name = "rpt_currencylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 If Check1.Value = 1 Then
 Call appcolor
 Else
 nocolor
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

Private Sub cmd_show_Click()
 If Check1.Value = 1 Then
 Call appcolor
 Else
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy
 End If
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "CURRENCY CODE LIST"
Me.Top = 10
Me.Left = 10
WebBrowser.Navigate "About:Blank"
Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select DISTINCT(c_name),c_desc from currency  order by c_name", Cn, 3, 2
While Not cr.EOF
List1.AddItem cr(0) & "  -  " & cr(1)
cr.MoveNext
Wend
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
   
   
      fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"

            
                fs.WriteLine "        <tr bgcolor=white  height=15 class=TableFont>"
                fs.WriteLine "            <td colspan=2><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td  colspan=3><b>CURRENCY LIST</td>"
                fs.WriteLine "           <td >Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                          
                fs.WriteLine "        </tr>"
    
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Currency</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Exchange Rate</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Description</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Date</td>"
   fs.WriteLine "            <td width=200><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim l As Integer
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from currencymaster where cur_currency='" & nm(0) & "' order by cur_date", Cn, 3, 2
While Not rs.EOF
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
fs.WriteLine "            <td  >" & rs!cur_currency & "</td>"
fs.WriteLine "            <td  >" & rs!cur_xchgrate & "</td>"
If rs!cur_desc = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!cur_desc & "</td>"
End If
fs.WriteLine "            <td  >" & rs!cur_date & "</td>"
If rs!notes = "" Or IsNull(rs!notes) Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!notes & "</td>"
End If
fs.WriteLine "        </tr>"

 sn = sn + 1
 rs.MoveNext
 Wend
 End If
 Next l
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

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
   
   
   fs.WriteLine "    <table border=0 cellspacing=0 width=95%>"
   fs.WriteLine "        <tr style='color=black;FONT-FAMILY: Arial Narrow;FONT-SIZE: 8pt;'>"
   fs.WriteLine "            <td align=center><font size=2.5 face=Arial Narrow>" & GetCompanyName & "</font></br><br> "
   fs.WriteLine "        </tr><tr>  <td align=center><font size=2>CURRENCY LIST</font></td>"
   fs.WriteLine "        </tr>"
   fs.WriteLine "    </table><br>"
   
  


      fs.WriteLine "    <table border=1 cellspacing=1 Bgcolor=blue width=95%>"

    
   'fs.WriteLine "            <td align=left bgcolor=white colspan=3><font size=3 face=arial><u><i><b>Complaints</font></br><br> "
   fs.WriteLine "        <tr bgcolor=blue height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Currency</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Exchange Rate</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Description</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Date</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Notes</td>"
   fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim l As Integer
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from currencymaster where cur_currency='" & nm(0) & "' order by cur_date", Cn, 3, 2
While Not rs.EOF
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td ><font color=blue>" & sn & "</td>"
fs.WriteLine "            <td ><font color=blue>" & rs!cur_currency & "</td>"
fs.WriteLine "            <td ><font color=blue>" & rs!cur_xchgrate & "</td>"
If rs!cur_desc = "" Then
fs.WriteLine "            <td  ><font color=blue>&nbsp;</td>"
Else
fs.WriteLine "            <td  ><font color=blue>" & rs!cur_desc & "</td>"
End If
fs.WriteLine "            <td  ><font color=blue>" & rs!cur_date & "</td>"
If rs!notes = "" Or IsNull(rs!notes) Then
fs.WriteLine "            <td  ><font color=blue>&nbsp;</td>"
Else
fs.WriteLine "            <td  ><font color=blue>" & rs!notes & "</td>"
End If
fs.WriteLine "        </tr>"

 sn = sn + 1
 rs.MoveNext
 Wend
 End If
 Next l
   fs.WriteLine " </table>"
   WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"

End Sub
