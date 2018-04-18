VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_VendorInvoiceSummary 
   BackColor       =   &H00DC7E5A&
   Caption         =   "Vendor Invoice Summary"
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
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmd_apply 
         Height          =   615
         Left            =   3480
         Picture         =   "rpt_VendorInvoiceSummary.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   8
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
               Picture         =   "rpt_VendorInvoiceSummary.frx":0612
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":1364
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":167E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":1AD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":1DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2104
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":241E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2738
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2A52
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2EA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":32F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":3610
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":392A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":3C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":18DB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":1F050
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":252EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":25604
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2575E
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":25A78
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":25ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":261E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":264FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":26950
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":26C6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":270BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":27516
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":27830
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":27B4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":27E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2817E
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":28498
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":288EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":28D3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":29056
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":29370
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":2968A
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":299A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "rpt_VendorInvoiceSummary.frx":29DF6
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
         TabIndex        =   10
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
      Left            =   5760
      TabIndex        =   11
      Top             =   840
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
         TabIndex        =   3
         Top             =   0
         Width           =   13815
         Begin VB.CommandButton cmd_print 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   9120
            Picture         =   "rpt_VendorInvoiceSummary.frx":2A588
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Click to Print"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmd_show 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   8280
            Picture         =   "rpt_VendorInvoiceSummary.frx":2AAFB
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Click to View"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmd_close 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   9960
            Picture         =   "rpt_VendorInvoiceSummary.frx":2B116
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Click to Exit"
            Top             =   840
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   300
            Left            =   8280
            TabIndex        =   14
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   39154
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   930
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   480
            Width           =   5055
         End
         Begin VB.CommandButton command2 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   12480
            Picture         =   "rpt_VendorInvoiceSummary.frx":2B715
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Click to View"
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00DC7E5A&
            Height          =   480
            Left            =   12480
            Picture         =   "rpt_VendorInvoiceSummary.frx":2BD30
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Click to Exit"
            Top             =   1320
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   300
            Left            =   5760
            TabIndex        =   13
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   39154
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Left            =   5760
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Left            =   8280
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor"
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
            TabIndex        =   6
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
      TabIndex        =   19
      Top             =   2280
      Width           =   10335
      ExtentX         =   18230
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
Attribute VB_Name = "rpt_VendorInvoiceSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rg As New ADODB.Recordset
 Dim hgg As Integer
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "REPORT - VENDOR INVOICE SUMMARY"
Me.Top = 10
Me.Left = 10
Option1.Value = False
Option2.Value = True
WebBrowser.Navigate "About:Blank"
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select Distinct(vendor_code),vendor_desc from vendormaster order by vendor_desc", Cn, 3, 2
While Not ls.EOF
List1.AddItem ls(0) & "  -  " & ls(1)
ls.MoveNext
Wend
ls.Close
Me.Width = 11415
Me.Height = 9750
dtpFrom.Value = Date
dtpto.Value = Date
End Sub

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
                fs.WriteLine "            <td colspan=4><b>" & GetCompanyName & "</td>"
                fs.WriteLine "           <td  ><b>VENDOR INVOICE SUMMARY</td>"
                fs.WriteLine "           <td colspan=2 >Report Date :  " & Format(Date, "dd/MM/yyyy") & "</td>"
                          
                fs.WriteLine "        </tr>"
   fs.WriteLine "        <tr bgcolor=black height=15 class=TableFont>"
   fs.WriteLine "            <td Nowrap><font color=white>SNo</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Vendor Code</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Description</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Inv. Date</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Inv. No</td>"
   fs.WriteLine "            <td Nowrap><font color=white>Job Charge</td>"
   fs.WriteLine "            <td align='center'><font color=white>Amount</td>"
   fs.WriteLine "        </tr>"
Dim sn As Integer
sn = 1
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
nm = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
'rs.Open "select * from vendormaster where vendor_code='" & nm(0) & "' order by vendor_code", Cn, 3, 2
rs.Open "select bd_vendor,vendor_desc, bd_invdate, bd_inv, bd_jobcharge, bd_extdamt From cost Inner Join vendormaster on cost.bd_vendor =  vendormaster.vendor_code   where bd_costtype = 'X' and  bd_invdate between '" & Format(dtpFrom.Value, "yyyy-MM-dd") & "' and  '" & Format(dtpto.Value, "yyyy-MM-dd") & "' and bd_vendor = '" & nm(0) & "' order by bd_invdate", Cn, 3, 2
While Not rs.EOF
fs.WriteLine "        <tr bgcolor=white height=15 class=TableFont>"
fs.WriteLine "            <td  >" & sn & "</td>"
fs.WriteLine "            <td  >" & rs!bd_vendor & "</td>"
fs.WriteLine "            <td  >" & rs!vendor_desc & "</td>"
fs.WriteLine "            <td  >" & Format(rs!bd_invdate, "dd/MM/yyyy") & "</td>"
fs.WriteLine "            <td  >" & rs!bd_inv & "</td>"
fs.WriteLine "            <td  >" & rs!bd_JobCharge & "</td>"
fs.WriteLine "            <td  align='right'>" & Format(rs!bd_extdamt, "###,###,##0.00") & "</td>"
decTotal = decTotal + rs!bd_extdamt
fs.WriteLine "        </tr>"
 sn = sn + 1
 rs.MoveNext
Wend
End If
decGrandTotal = decTotal + decGrandTotal
If List1.Selected(l) = True And decTotal > 0 Then fs.WriteLine "           <tr  bgcolor=white height=15 class=TableFont> <td colspan=7 align='right' ><b> Total = " & Format(decTotal, "###,###,##0.00") & "</td></td>"
decTotal = 0
Next l
 fs.WriteLine "           <tr  bgcolor=white height=15 class=TableFont> <td colspan=7 align='right' ><b> Grand total = " & Format(decGrandTotal, "###,###,##0.00") & "</td></td>"
   fs.WriteLine " </table>"
   PrintEndofReport fs
   fs.WriteLine "    </body>"
   fs.WriteLine "    <html>"
  If boolSaveAsExcel = True Then
  WebBrowser.Navigate "C:\PCIS-Reports\" & filpat
  Else
   WebBrowser.Navigate App.Path & "\rep.html"
 End If
End Function
Private Sub PrintEndofReport(fs As Object)
   fs.WriteLine "    <table border=1 class=TableFont cellspacing=0 BORDERCOLOR=gray width=95%>"

Dim f As Integer
f = 0

fs.WriteLine " <tr><td ><b>Date From : " & Format(dtpFrom.Value, "DD/MM/yyyy") & " to " & Format(dtpto.Value, "dd/MM/yyyy") & " </td></tr>"
fs.WriteLine "           <br></br> <td ><b>Vendors Selected</td>"
For f = 0 To List1.ListCount - 1
If List1.Selected(f) = True Then
hh = Split(List1.List(f), "  -  ", Len(List1.List(f)), vbTextCompare)
fs.WriteLine "        <tr bgcolor=white  class=TableFont>"
fs.WriteLine "            <td > " & List1.List(f) & "</td></tr>"
End If
Next f
 Dim r As Integer
r = 0
 
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
filpat = "Vendor Invoice Summary" & "-" & txt_name.Text & "-" & st & ".xls"
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
'frmBusy.Show
'SetParent frmBusy.hwnd, rpt_VendorInvoiceSummary.hwnd
'frmBusy.lblBusyString = "Please Wait Report Under Process......"
If chkSaveAsExcel.Value Then
frmExportToExcel.Visible = True
Else
Call WriteByResource(False)
End If
'Unload frmBusy
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
