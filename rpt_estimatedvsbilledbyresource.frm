VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form rpt_estimatedvsbilledbyresource 
   BackColor       =   &H00DC7E5A&
   Caption         =   "Estimated vs Billed by Resource"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.Frame frmExportToExcel 
      BackColor       =   &H80000009&
      Caption         =   "Export as Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3720
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmd_apply 
         Height          =   615
         Left            =   3480
         Picture         =   "rpt_estimatedvsbilledbyresource.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   960
         Width           =   2415
      End
      Begin MSComctlLib.ImageList ImageList51 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   39
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":0612
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":1364
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":167E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":1AD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":1DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2104
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":241E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2738
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2A52
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2EA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":32F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":3610
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":392A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":3C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":18DB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":1F050
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":252EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":25604
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2575E
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":25A78
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":25ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":261E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":264FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":26950
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":26C6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":270BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":27516
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":27830
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":27B4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":27E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2817E
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":28498
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":288EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":28D3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":29056
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":29370
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":2968A
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":299A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_estimatedvsbilledbyresource.frx":29DF6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkSaveAsExcel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save Report as Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   1935
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   13815
         Begin VB.ListBox lst_resc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   930
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   480
            Width           =   3735
         End
         Begin VB.ListBox lst_prj 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   930
            Left            =   6000
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   450
            Width           =   3975
         End
         Begin VB.ComboBox cbo_year 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4440
            TabIndex        =   10
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "All Projects By Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3960
            TabIndex        =   9
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CommandButton command2 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   12480
            Picture         =   "rpt_estimatedvsbilledbyresource.frx":2A588
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Click to View"
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   12480
            Picture         =   "rpt_estimatedvsbilledbyresource.frx":2ABA3
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Click to Exit"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            Left            =   3960
            TabIndex        =   15
            Top             =   600
            Visible         =   0   'False
            Width           =   855
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
            Left            =   6000
            TabIndex        =   14
            Top             =   240
            Width           =   975
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
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5640
         Top             =   120
         Width           =   5415
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
      Width           =   11655
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   9240
         Picture         =   "rpt_estimatedvsbilledbyresource.frx":2B1A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to Exit"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_show 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   7560
         Picture         =   "rpt_estimatedvsbilledbyresource.frx":2B7A1
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click to View"
         Top             =   80
         Width           =   735
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00DC7E5A&
         Height          =   480
         Left            =   8400
         Picture         =   "rpt_estimatedvsbilledbyresource.frx":2BDBC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to Print"
         Top             =   80
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   8
         Height          =   1335
         Left            =   5565
         Top             =   0
         Width           =   5415
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   6495
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   11415
      ExtentX         =   20135
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
End
Attribute VB_Name = "rpt_estimatedvsbilledbyresource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rg As New ADODB.Recordset
 Dim hgg As Integer
Function ComputeTotal(decUnirate As Double, decExchangeRate As Double, decQuantity As Double, decPeriod As Double)
 Dim decComputeTotal
 If decQuantity > 0 Then
 ComputeTotal = decUnirate * (decExchangeRate / decQuantity) * decPeriod
 Else
 ComputeTotal = 0
 End If
 End Function
 Public Sub RPTHEADING_Standard(fs As Object)
fs.WriteLine "    <table border=1 class=TableFont cellspacing=0 BORDERCOLOR=gray width=95%>"
fs.WriteLine "            <td colspan=2>" & GetCompanyName & "</td>"
fs.WriteLine "           <td><b>Project</td>"
fs.WriteLine "           <td><b>Resource</b></td>"
fs.WriteLine "           <td align=center>See end of Report</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=white  height=25 class=TableFont>"
fs.WriteLine "            <td Colspan=2>ESTIMATED vs BILLED BY RESOURCE</td>"
fs.WriteLine "           <td >&nbsp;</td>"
fs.WriteLine "           <td><b>Cut-OffDate</td>"
fs.WriteLine "           <td align=center>" & main.DTPcutdate1.Value & "</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
fs.WriteLine "           <td colspan=4><font color=white><b>PrintDate</td>"
fs.WriteLine "           <td colspan=4><font color=white>" & Format(Date, "dd/MM/yyyy") & "</td>"
fs.WriteLine "        </tr>"
fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
fs.WriteLine "            <td nowrap><font color=white>Resource Code</td>"
fs.WriteLine "            <td nowrap><font color=white>Resource Name</td>"
fs.WriteLine "            <td nowrap><font color=white>Cost Code</td>"
fs.WriteLine "            <td nowrap><font color=white>Billed</td>"
fs.WriteLine "            <td nowrap><font color=white>Estimated</td>"
fs.WriteLine "        </tr>"
End Sub
Function WriteByResource(boolSaveAsExcel As Boolean)
Dim fso As New FileSystemObject
   Dim fs As Object
   Dim decDuration As Double, decTotal, decGrandTotal
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
        RPTHEADING_Standard fs
        cnt = 0
Dim strProjectList As String
Dim decSelListCount As Integer
decSelListCount = 0
For decSelListCount = 0 To lst_prj.ListCount - 1
If lst_prj.Selected(decSelListCount) = True Then
nmd = Split(lst_prj.List(decSelListCount), "  -  ", Len(lst_prj.List(decSelListCount)), vbTextCompare)
'arrProjectlist(decSelListCount + 1) = nmd(0)
strProjectList = strProjectList & "'" & nmd(0) & "'" & ", "
End If
Next decSelListCount
strProjectList = Mid$(strProjectList, 1, Len(strProjectList) - 2)
 i = 0
'With flex_grid
'        .Rows = 1
For i = 0 To lst_resc.ListCount - 1
If lst_resc.Selected(i) = True Then
nmm = Split(lst_resc.List(i), "  -  ", Len(lst_resc.List(i)), vbTextCompare)
'
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
Dim fldata1 As New ADODB.Recordset
If fldata1.State Then fldata1.Close
Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
Dim jc As New ADODB.Recordset
        If jc.State Then jc.Close
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
'fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_year='" & cbo_year.Text & "' and bd_costtype='E'   order by bd_sdate,bd_edate", Cn, 3, 2
'fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey in (" & strProjectList & ") and bd_costtype='X'   order by bd_sdate,bd_edate", Cn, 3, 2
'rsBilledData.Open "select sum(bd_extdamt) from cost  where bd_resccode='" & nmm(0) & "'  and bd_projectkey in (" & strProjectList & ") and bd_costtype='X'", Cn, 3, 2
fldata.Open "select sum(bd_extdamt) from cost where bd_resccode='" & nmm(0) & "' and bd_projectkey in (" & strProjectList & ") and bd_costtype = 'X' group by bd_resccode, bd_rescname, bd_costtype", Cn, 3, 2
If Not fldata.EOF Then
If CDbl(fldata(0)) > 0 Then
fldata1.Open "select bd_resccode, bd_rescname,bd_costtype, sum(bd_extdamt) as cost, bd_costcode from cost where bd_resccode='" & nmm(0) & "' and bd_projectkey in (" & strProjectList & ") and bd_costtype = 'X' group by bd_resccode, bd_rescname, bd_costtype, bd_costcode", Cn, 3, 2
fldata2.Open "select sum(bd_extdamt) from cost where bd_resccode='" & nmm(0) & "' and bd_projectkey in (" & strProjectList & ") and bd_costtype = 'E' group by bd_resccode, bd_rescname, bd_costtype", Cn, 3, 2
   While Not fldata1.EOF
        cnt = cnt + 1 '********************************
                If cnt >= 53 Then
                fs.WriteLine "</table><P></P>"
                RPTHEADING_Standard fs
                cnt = 0
                End If
        fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
        fs.WriteLine "            <td nowrap>" & fldata1!bd_resccode & "</td>"
        fs.WriteLine "            <td nowrap>" & fldata1!bd_rescname & "</td>"
        fs.WriteLine "            <td nowrap>" & fldata1!bd_costcode & "</td>"
        fs.WriteLine "            <td nowrap ALIGN=RIGHT>" & Format(fldata1!cost, "###,###,##0.00") & "</td>"
        If Not fldata2.EOF Then
        fs.WriteLine "            <td nowrap ALIGN=RIGHT>" & Format(fldata2(0), "###,###,##0.00") & "</td>"
        estimatedTotal = estimatedTotal + fldata2(0)
        Else
        fs.WriteLine "            <td nowrap ALIGN=RIGHT>" & Format(0, "###,###,##0.00") & "</td>"
        estimatedTotal = estimatedTotal + 0
        End If
        fs.WriteLine "        </tr>"
        billedtotal = billedtotal + fldata1!cost

        fldata1.MoveNext
    Wend
End If
'''Next j
End If
End If
Next i
   fs.WriteLine "  <tr bgcolor=black class=TableFont>"
   fs.WriteLine " <td colspan=4 ALIGN=RIGHT><font color=white><b>BILLED TOTAL - " & Format(billedtotal, "###,###,##0.00") & "</td>"
   fs.WriteLine " <td ALIGN=RIGHT><font color=white><b>ESTIMATED TOTAL - " & Format(estimatedTotal, "###,###,##0.00") & "</td> </tr>"
   'WebBrowser.Navigate App.Path & "\rep.html"
   fs.WriteLine "    </table><br>"
   PrintEndofReport fs
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
  If boolSaveAsExcel = True Then
  WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
  Else
   WebBrowser.Navigate App.Path & "\rep.html"
 End If
'End With
End Function
Private Sub PrintEndofReport(fs As Object)
   fs.WriteLine "    <table border=1 class=TableFont cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0
fs.WriteLine "           <br></br> <td ><b>Resources Selected</td>"
For f = 0 To lst_resc.ListCount - 1
If lst_resc.Selected(f) = True Then
hh = Split(lst_resc.List(f), "  -  ", Len(lst_resc.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & lst_resc.List(f) & "</td></tr>"
End If
Next f

 
 Dim r As Integer
r = 0
fs.WriteLine "            <td > <b>Projects Selected</td>"
For r = 0 To lst_prj.ListCount - 1
If lst_prj.Selected(r) = True Then
hh = Split(lst_prj.List(r), "  -  ", Len(lst_prj.List(r)), vbTextCompare)
 fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & lst_prj.List(r) & "</td></tr>"
End If
Next r
 
fs.WriteLine " </table>"

End Sub
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

Private Sub cmd_apply_Click()
Dim st As String
st = Format(Date, "dd-MMM-yyyy")
filpat = "Estimated_vs_Billed by Resource" & "-" & txt_name.Text & "-" & st & ".xls"
ms = MsgBox("Do you want to save the report with the name  " & filpat, vbYesNo)
If ms = vbYes Then
Call WriteByResource(True)
frmExportToExcel.Visible = False
Else
frmExportToExcel.Visible = False
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
'Call flex_dataallreport
Unload frmBusy

End Sub

Public Sub RPTHEADING(fs As Object)
fs.WriteLine "    <table border=1 cellspacing=0 BORDERCOLOR=gray width=95%>"
fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
                fs.WriteLine "            <td colspan=14><font color=white><b>" & GetCompanyName & "</td>"
                fs.WriteLine "        </tr>"
                
                fs.WriteLine "        <tr bgcolor=black  height=20 class=TableFont>"
                
                fs.WriteLine "            <td colspan=7><font color=white><b>EIC BY RESOURCE</td>"
                fs.WriteLine "           <td colspan=4><font color=white><b>PrintDate</td>"
                fs.WriteLine "           <td colspan=3><font color=white>" & Format(Date, "dd/MM/yyyy") & "</td>"
                fs.WriteLine "        </tr>"
                fs.WriteLine "        <tr bgcolor=black  height=15 class=TableFont>"
                fs.WriteLine "            <td Nowrap><font color=white>StartDate</td>"
                fs.WriteLine "            <td Nowrap><font color=white>EndDate</td>"
                fs.WriteLine "            <td Nowrap><font color=white>Duration (Days)</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Jobcharge</td>"
                fs.WriteLine "            <td Nowrap><font color=white>Qty</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Curr</td>"
                fs.WriteLine "            <td Nowrap><font color=white>UnitRate</td>"
                fs.WriteLine "            <td Nowrap><font color=white>UOM</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Xchg</td>"
                fs.WriteLine "            <td Nowrap><font color=white>Total Cost(RM)</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Spread</td>"
                fs.WriteLine "            <td Nowrap ><font color=white>Type</td>"
                fs.WriteLine "            <td Nowrap><font color=white>CostCode</td>"
                fs.WriteLine "            <td Nowrap><font color=white>Notes</td>"
                fs.WriteLine "        </tr>"
    
End Sub

Private Sub Form_Load()
main.lbltitle.Caption = "Report - Estimated vs Billed By RESOURCE"
Me.Top = 10
Me.Left = 10
Me.WindowState = vbMaximized
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
Private Sub cboResource_Click()
spp = Split(cboResource.Text, "  -  ", Len(cboResource.Text), vbTextCompare)
List2.Clear
lstProjects.Clear
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
            'If rc.State Then rc.Close
            'rc.Open "select DISTINCT(bd_resccode) from cost c, jobcharge j where c.bd_jobcharge=j.job_code  and  bd_costtype='E' order by c.bd_resccode", Cn, 3, 2
            'While Not rc.EOF
            Dim rcd As New ADODB.Recordset
            If rcd.State Then rcd.Close
            'rcd.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rc(0) & "' ", Cn, 3, 2
            rcd.Open "select distinct(bd_projectkey), bd_projectdesc from cost where bd_resccode = '" & spp(0) & "' and bd_projectkey in (select DISTINCT(p.proj_key) from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "')", Cn, 3, 2
                   While Not rcd.EOF
                   lstProjects.AddItem rcd(0) & "  -  " & rcd(1)
            rcd.MoveNext
            Wend
            rcd.Close

  
Check1.Value = 1
Check2.Value = 1
Check3.Value = 1


            hgg = 0
            For hgg = 0 To List2.ListCount - 1
            List2.Selected(hgg) = False
            Next hgg
            hgg = 0
            For hgg = 0 To lstProjects.ListCount - 1
            lstProjects.Selected(hgg) = False
            Next hgg
            Option1.Value = 0
            Option2.Value = 0
            Option3.Value = 0
            Option4.Value = 0
End Sub



Private Sub cmd_show_Click()
frmBusy.Show
SetParent frmBusy.hwnd, rpt_estimatedvsbilledbyresource.hwnd
frmBusy.lblBusyString = "Please Wait Report Under Process......"
If chkSaveAsExcel.Value Then
frmExportToExcel.Visible = True
Else
Call WriteByResource(False)
End If
Unload frmBusy
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub Option1_Click()
Option3.Value = 0
Option4.Value = 0
hgg = 0
            For hgg = 0 To lstProjects.ListCount - 1
            lstProjects.Selected(hgg) = False
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
            For hgg = 0 To lstProjects.ListCount - 1
            lstProjects.Selected(hgg) = False
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
For g = 0 To lstProjects.ListCount - 1
lstProjects.Selected(g) = False
Next g
End If
End Sub
Private Sub Option4_Click()
If Option4.Value = True Then
Dim f As Integer
f = 0
For f = 0 To lstProjects.ListCount - 1
lstProjects.Selected(f) = True
Next f
End If
End Sub
