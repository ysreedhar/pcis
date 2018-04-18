VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_jobno 
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14205
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   11415
      ExtentX         =   20135
      ExtentY         =   12938
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
      Width           =   11175
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   10320
         Picture         =   "rpt_jobno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to Exit"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8640
         Picture         =   "rpt_jobno.frx":05FF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to View"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9480
         Picture         =   "rpt_jobno.frx":0C1A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to Print"
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Random Selection"
            Height          =   255
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   930
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   4695
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
Attribute VB_Name = "rpt_jobno"
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
Call nocolor
Unload frmBusy
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "JOB NO. LIST"
Me.Top = 10
Me.Left = 10
Option1.Value = False
Option2.Value = True
WebBrowser.Navigate "About:Blank"
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not ls.EOF
List1.AddItem ls(0) & "  -  " & ls(1)
ls.MoveNext
Wend
ls.Close
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

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from jobno where job_key='" & nm(0) & "' order by jobno_code", Cn, 3, 2
While Not rs.EOF
  cnt = cnt + 1 '********************************
                                If cnt >= 55 Then
                                fs.WriteLine "</table><P></P>"
                                RPTHEADING fs
                                cnt = 0
                                End If
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
fs.WriteLine "            <td  >" & rs!job_key & "</td>"
If rs!jobno_code = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!jobno_code & "</td>"
End If
If rs!jobno_desc = "" Then
fs.WriteLine "            <td  >&nbsp;</td>"
Else
fs.WriteLine "            <td  >" & rs!jobno_desc & "</td>"
End If
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

 


Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
    
   
            
                fs.WriteLine "        <tr bgcolor=white  height=20 class=TableFont>"
                fs.WriteLine "            <td colspan=3><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td  ><b>JOB LIST</td>"
                fs.WriteLine "           <td colspan=1 >Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                          
                fs.WriteLine "        </tr>"
    
    
   
   fs.WriteLine "        <tr bgcolor=black height=20 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Project Key</td>"
   fs.WriteLine "            <td Nowrap><font color=white>JobNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Job Desc</td>"
   fs.WriteLine "            <td width=200><font color=white>Notes</td>"
   
   fs.WriteLine "        </tr>"
End Sub
