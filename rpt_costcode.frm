VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_costcode 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   15045
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7215
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   12726
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
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10680
         Picture         =   "rpt_costcode.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to Exit"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8880
         Picture         =   "rpt_costcode.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to View"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9780
         Picture         =   "rpt_costcode.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to Print"
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DC7E5A&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   1200
         Width           =   4335
         Begin VB.OptionButton opt_cost 
            BackColor       =   &H00DC7E5A&
            Caption         =   "Only CostCode"
            Height          =   255
            Left            =   2160
            TabIndex        =   8
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton opt_resc 
            BackColor       =   &H00DC7E5A&
            Caption         =   "Detail By Resource"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3720
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   2
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
Attribute VB_Name = "rpt_costcode"
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

If opt_resc.Value = True Then
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
Unload frmBusy
ElseIf opt_cost.Value = True Then
Load frmBusy
frmBusy.Show
frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor1
Unload frmBusy
Else
MsgBox "Select Option"
Exit Sub
End If


End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "COST CODE LIST"
Me.Top = 10
Me.Left = 10
Option1.Value = False
Option2.Value = True
WebBrowser.Navigate "About:Blank"
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select Distinct(cc_code),cc_desc from costcode order by cc_code", Cn, 3, 2
While Not ls.EOF
List1.AddItem ls(0) & "  -  " & ls(1)
ls.MoveNext
Wend
ls.Close
opt_resc.Value = False
opt_cost.Value = False
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
RPTHEADING fs
cnt = 0
Dim sn As Integer
sn = 1
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
   nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
'
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from  costcode where cc_code='" & nm(0) & "' order by cc_code", Cn, 3, 2
While Not rs.EOF
 cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=#acacac height=15 class=TableFont>"
fs.WriteLine "            <td  ><font color=black>" & sn & "</td>"
fs.WriteLine "            <td  ><font color=black>" & rs!cc_code & "</td>"
If rs!cc_desc = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  ><font color=black>" & rs!cc_desc & "</td>"
End If
 
If rs!notes = "" Or IsNull(rs!notes) Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  ><font color=black>" & rs!notes & "</td>"
End If
fs.WriteLine "        </tr>"
 sn = sn + 1
Dim kk As New ADODB.Recordset
If kk.State Then kk.Close
kk.Open "select * from resourcecostcode  where rcc_id=" & rs!cc_id, Cn, 3, 2
While Not kk.EOF
 cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
            fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
            fs.WriteLine "            <td  >&nbsp;</td>"
            fs.WriteLine "            <td  >" & kk!rcc_resource & "</td>"
            Dim rcc As New ADODB.Recordset
            If rcc.State Then rcc.Close
            rcc.Open "select * from resourcemaster where resc_code='" & kk!rcc_resource & "' ", Cn, 3, 2
            If Not rcc.EOF Then
                        If rcc!resc_desc = "" Or IsNull(rcc!resc_desc) Then
                        fs.WriteLine "            <td  >&nbsp;</td>"
                        Else
                        fs.WriteLine "            <td  >" & rcc!resc_desc & "</td>"
                        End If
            End If
            
            If kk!notes = "" Or IsNull(kk!notes) Then
            fs.WriteLine "            <td  >&nbsp;</td>"
            Else
            fs.WriteLine "            <td  >" & kk!notes & "</td>"
            End If
            fs.WriteLine "        </tr>"



kk.MoveNext
Wend


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

 



Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"



fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
fs.WriteLine "            <td colspan=2><b>" & GetCompanyName & "</td>"
fs.WriteLine "           <td  ><b>COST CODE BY RESOURCE</td>"
fs.WriteLine "           <td colspan=1 >Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"

fs.WriteLine "        </tr>"



fs.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
fs.WriteLine "            <td Nowrap><font color=white>CostCode</td>"
fs.WriteLine "            <td Nowrap><font color=white>CostCode Desc</td>"
fs.WriteLine "            <td width=200><font color=white>Notes</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=white>&nbsp;</td>"
fs.WriteLine "            <td Nowrap><font color=white>Resource Code</td>"
fs.WriteLine "            <td Nowrap><font color=white>Resource Desc</td>"
fs.WriteLine "            <td width=200><font color=white>Notes</td>"
fs.WriteLine "        </tr>"

End Sub

Public Sub nocolor1()
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
   
   

Dim cnt As Integer
RPTHEADING1 fs
cnt = 0
Dim sn As Integer
sn = 1
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
   nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
'
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from  costcode where cc_code='" & nm(0) & "' order by cc_code", Cn, 3, 2
While Not rs.EOF
 cnt = cnt + 1 '********************************
                                If cnt >= 58 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING1 fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  > " & sn & "</td>"
fs.WriteLine "            <td  > " & rs!cc_code & "</td>"
If rs!cc_desc = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  > " & rs!cc_desc & "</td>"
End If
 
If rs!notes = "" Or IsNull(rs!notes) Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  > " & rs!notes & "</td>"
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
Public Sub RPTHEADING1(fs As Object)


fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"



fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
fs.WriteLine "            <td colspan=2><b>" & GetCompanyName & "</td>"
fs.WriteLine "           <td  ><b>COST CODE </td>"
fs.WriteLine "           <td colspan=1 >Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"

fs.WriteLine "        </tr>"



fs.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
fs.WriteLine "            <td Nowrap><font color=white>CostCode</td>"
fs.WriteLine "            <td Nowrap><font color=white>CostCode Desc</td>"
fs.WriteLine "            <td width=200><font color=white>Notes</td>"
fs.WriteLine "        </tr>"
 
End Sub
