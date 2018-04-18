VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_progressdurationdetails 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ESTIMATED PROGRESS DURATION BY SPREAD"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   11190
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11415
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10935
         Begin VB.ComboBox cbo_spr 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Text            =   " "
            Top             =   150
            Width           =   5535
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   7920
            TabIndex        =   7
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Spread"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   600
            TabIndex        =   6
            Top             =   240
            Width           =   1140
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      TextStyle       =   3
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "ar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "hlp"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excel"
            Object.ToolTipText     =   "Transfer To Excel"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7920
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   2
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   58
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_progressdurationdetails.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_progressdurationdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub cbo_spr_Change()
Call flex_data
End Sub
Private Sub cbo_spr_Click()
Call flex_data
End Sub
Private Sub cbo_spr_KeyPress(KeyAscii As Integer)
On Error Resume Next
'KeyAscii = 0
End Sub

Private Sub dt1_Change()
Call halfsum
End Sub

Private Sub dt1_Click()
Call halfsum
End Sub

 

 

Private Sub cmdCloseframe_Click()
frmDuplicateSpreadTransactions.Visible = False
End Sub

Private Sub cmdDuplicate_Click()
If lst_jobCharge.ListCount = 0 Then LoadJobCharges
frmDuplicateSpreadTransactions.Visible = True
End Sub
Function LoadJobCharges()
Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(job_code), job_desc from jobcharge order by job_code", Cn, 3, 2
While Not jc.EOF
lst_jobCharge.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
End Function
Function GetResources()
Dim rsResources As New ADODB.Recordset
If rsResources.State Then rsResources.Close
rsResources.Open "select * from resourcedetails", Cn, 3, 2
While Not rsResources.EOF
'resources
Wend
End Function
Private Sub flex_grid_Click()
On Error Resume Next
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
'bacl color
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbYellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture
 Call halfsum
 vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
'bacl color
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbYellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture


'--END---------


Unload progressduration
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)
 
progressduration.cbo_spreadcode.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
progressduration.cbo_jobcharge.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
progressduration.txt_type.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
progressduration.DTP_startdate.Value = Format(flex_grid.TextMatrix(flex_grid.Row, 4), "dd-MM-yyyy H:mm:ss")
progressduration.DTP_enddate.Value = Format(flex_grid.TextMatrix(flex_grid.Row, 5), "dd-MM-yyyy H:mm:ss")
Dim ju As New ADODB.Recordset
If ju.State Then ju.Close
ju.Open "select * from progressdurationdetails where prgs_id=" & id, Cn, 3, 2
If Not ju.EOF Then
progressduration.txt_days.Text = ju!prgs_days
Else
progressduration.txt_days.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
End If
progressduration.txt_remarks.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
progressduration.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 8)
 
progressduration.Show
progressduration.Top = 3200
progressduration.Left = 0
progressduration.Height = 3210
progressduration.Width = 6420

progressduration.cbo_jobcharge.Enabled = False

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "ESTIMATED PROGRESS DURATION BY SPREAD"
Call flex_title
 
Call flex_data
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(spread_code),spread_desc from spreadmaster where spread_code <>'NA' order by spread_code", Cn, 3, 2
While Not spr.EOF
cbo_spr.AddItem spr(0) & "  -  " & spr(1)
spr.MoveNext
Wend
spr.Close

dt1.Value = Format(Date, "dd-MM-yyyy H:MM:ss")
 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()

On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Spread"
        .ColWidth(1) = 1100
        .ColAlignment(2) = 0
        .TextMatrix(0, 2) = "JobCharge"
        .ColWidth(2) = 3500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Type"
        .ColWidth(3) = 300
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Start Date"
        .ColWidth(4) = 2000
        .TextMatrix(0, 5) = "End Date"
        .ColWidth(5) = 2000
        .TextMatrix(0, 6) = "Days"
        .ColWidth(6) = 800
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 6800
        .ColAlignment(7) = 0
        .ColWidth(8) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload progressduration
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
 progressduration.cbo_jobcharge.Enabled = True
If cbo_spr.Text = " " Then
MsgBox "select Spread"
cbo_spr.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload progressduration
progressduration.Show
progressduration.Top = 3200
progressduration.Left = 0
progressduration.Height = 3210
progressduration.Width = 6420
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
If progressduration.cbo_spreadcode.Text = "" Then
MsgBox "Select Spread"
progressduration.cbo_spreadcode.SetFocus
Exit Sub
End If
If progressduration.cbo_jobcharge.Text = "" Then
MsgBox "Select JobCharge"
progressduration.cbo_jobcharge.SetFocus
Exit Sub
End If
nb = Split(progressduration.cbo_spreadcode.Text, "  -  ", Len(progressduration.cbo_spreadcode.Text), vbTextCompare)
nbb = Split(progressduration.cbo_jobcharge.Text, "  -  ", Len(progressduration.cbo_jobcharge.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from progressdurationdetails", Cn, 3, 2
sv.AddNew
sv!prgs_spread_code = nb(0)
sv!prgs_job_key = nbb(0)
sv!prgs_startdate = Format(progressduration.DTP_startdate.Value, "dd-MM-yyyy H:mm:ss")
sv!prgs_enddate = Format(progressduration.DTP_enddate.Value, "dd-MM-yyyy H:mm:ss")
sv!prgs_remarks = progressduration.txt_remarks.Text
sv!prgs_days = progressduration.txt_days.Text
sv!t_date = progressduration.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!prgs_type = progressduration.txt_type.Text
sv.Update
sv.Close

MsgBox "New Progress Duration Added Succesfully"
Call progcost
Unload progressduration
Call flex_data
Call flex_title
Exit Sub
assad:
      MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad2
nb = Split(progressduration.cbo_spreadcode.Text, "  -  ", Len(progressduration.cbo_spreadcode.Text), vbTextCompare)
nbb = Split(progressduration.cbo_jobcharge.Text, "  -  ", Len(progressduration.cbo_jobcharge.Text), vbTextCompare)

Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from progressdurationdetails where prgs_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!prgs_spread_code = nb(0)
md!prgs_job_key = nbb(0)
md!prgs_startdate = Format(progressduration.DTP_startdate.Value, "dd-MM-yyyy H:mm:ss")
md!prgs_enddate = Format(progressduration.DTP_enddate.Value, "dd-MM-yyyy H:mm:ss")
md!prgs_remarks = progressduration.txt_remarks.Text
md!prgs_days = progressduration.txt_days.Text
md!t_date = progressduration.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md!prgs_type = progressduration.txt_type.Text
md.Update
md.Close
MsgBox "Selected Progress Duration Modified"
End If
Call progcost
Unload progressduration
Call flex_data
Call flex_title
Exit Sub
assad2:
      MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
gf = Split(flex_grid.TextMatrix(flex_grid.Row, 1), "  -  ", Len(flex_grid.TextMatrix(flex_grid.Row, 1)), vbTextCompare)
bf = Split(flex_grid.TextMatrix(flex_grid.Row, 2), "  -  ", Len(flex_grid.TextMatrix(flex_grid.Row, 2)), vbTextCompare)
Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from cost where bd_spread='" & gf(0) & "' and bd_jobcharge='" & bf(0) & "' and bd_costtype='E' ", Cn, 3, 2
If Not dlk.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If


Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then

Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from progressdurationdetails where prgs_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload progressduration
Call flex_data
Call flex_title
Else
Unload progressduration
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload progressduration
ElseIf Button.Caption = "Excel" Then

Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
AppActivate "Export to Excel"
For i = 0 To flex_grid.Rows - 1
   flex_grid.Row = i
    For n = 1 To 7
        flex_grid.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flex_grid.Text
    Next
Next
End If




End Sub

Public Sub flex_data()
On Error Resume Next
Dim sdate As Date
Dim edate As Date
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double

nb = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
Dim i As Integer
i = 0

Dim v As Date
Dim fl  As New ADODB.Recordset
If fl.State Then fl.Close
fl.Open "select * from progressdurationdetails where prgs_spread_code='" & nb(0) & "'   order by prgs_startdate, prgs_id desc ", Cn, 3, 2
For i = 0 To fl.RecordCount - 1
v = fl!prgs_enddate
fl.MoveNext
fl!prgs_startdate = v
Dim dt As Double
Dim dt1 As Double
Dim dt2 As Double
dt = 0: dt1 = 0: dt2 = 0
dt = CDbl(fl!prgs_startdate)
dt1 = CDbl(fl!prgs_days)
dt2 = dt + dt1
fl!prgs_enddate = Format(dt2, "dd-MM-yyyy H:mm:ss")
fl.Update

Next i

Dim dys1 As Double
dys1 = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from progressdurationdetails where prgs_spread_code='" & nb(0) & "'   order by prgs_startdate ", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata(1) & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 1) = fldata(1) & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 1) = fldata(1)
        End If
        spr.Close
        Dim jc As New ADODB.Recordset
        If jc.State Then jc.Close
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata(2) & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata(2) & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata(2)
        End If
        .TextMatrix(.Rows - 1, 3) = fldata!prgs_type
        .TextMatrix(.Rows - 1, 4) = Format(fldata(3), "dd-MM-yyyy H:mm:ss")
        .TextMatrix(.Rows - 1, 5) = Format(fldata(4), "dd-MM-yyyy H:mm:ss")
        .TextMatrix(.Rows - 1, 6) = Round(fldata("prgs_days"), 5)
        dys1 = dys1 + fldata!prgs_days
        .TextMatrix(.Rows - 1, 7) = fldata("prgs_remarks")
        .TextMatrix(.Rows - 1, 8) = fldata("t_date")
        fldata.MoveNext
    Wend
End With

Label3.Caption = " Actual Days" & " " & Format(dys1, "###,###,##0.00")
 

End Sub
Public Sub progcost()
  If progressduration.cbo_jobcharge.Text = "" Then Exit Sub
  If progressduration.cbo_spreadcode.Text = "" Then Exit Sub
bh = Split(progressduration.cbo_jobcharge.Text, "  -  ", Len(progressduration.cbo_jobcharge.Text), vbTextCompare)
Pi = Split(progressduration.cbo_spreadcode.Text, "  -  ", Len(progressduration.cbo_spreadcode.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim ntotal As Double
ntotal = 0
Dim iddd As Double
iddd = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost where bd_jobcharge='" & bh(0) & "' and bd_spread='" & Pi(0) & "' and bd_costtype='E' and bd_spread <>'NA' ", Cn, 3, 2


    While Not fldata.EOF

     iddd = fldata!bd_id
mm = Split(fldata!bd_spread, "  -  ", Len(fldata!bd_spread), vbTextCompare)
mmm = Split(fldata!bd_jobcharge, "  -  ", Len(fldata!bd_jobcharge), vbTextCompare)
mmmm = Split(fldata!bd_resccode, "  -  ", Len(fldata!bd_resccode), vbTextCompare)

Dim dt1 As Date
Dim dt2 As Date
Dim pp As New ADODB.Recordset
If pp.State Then pp.Close
pp.Open "select * from progressdurationdetails where prgs_spread_code='" & fldata!bd_spread & "' and prgs_type='" & fldata!bd_type & "' and prgs_job_key='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
If Not pp.EOF Then
dt1 = pp!prgs_startdate
dt2 = pp!prgs_enddate
End If

Dim fldata2 As New ADODB.Recordset
If fldata2.State Then fldata2.Close
fldata2.Open "select * from cost where    bd_jobcharge='" & fldata!bd_jobcharge & "' and bd_costtype='E'  and bd_spread='" & fldata!bd_spread & "' and bd_id=" & iddd, Cn, 3, 2 'and bd_spread <> 'NA'

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


Dim cid As Integer
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select * from cost where    bd_jobcharge='" & bh(0) & "' and bd_costtype='E'  and bd_spread='" & Pi(0) & "' and bd_costtype='E' and bd_spread ='NA' ", Cn, 3, 2
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
                    
                 
 End If
cd.Update

cd.MoveNext
Wend
End Sub

Public Sub halfsum()
Dim ko As Double
ko = flex_grid.TextMatrix(flex_grid.Row, 0)

nb = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
Dim dyse As Double
dyse = 0
 


 
    For l = 1 To flex_grid.Row
      
 
        dyse = dyse + flex_grid.TextMatrix(l, 6)
  
    Next l
 

Label3.Caption = " Actual Days" & " " & dyse
End Sub
