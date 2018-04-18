VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_budgetedduration 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Budgeted Duration"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11415
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   11175
         Begin VB.ComboBox cbo_spr 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Text            =   " "
            Top             =   150
            Width           =   5535
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8040
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   240
            Width           =   1140
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11415
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   3
      Cols            =   7
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
      Width           =   11505
      _ExtentX        =   20294
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
         Left            =   8200
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
            Picture         =   "frm_budgetedduration.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_budgetedduration.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FF8080&
      Caption         =   "   Budgeted Cost"
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
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frm_budgetedduration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
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

Private Sub flex_grid_Click()

On Error Resume Next
'back color

Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True



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
'Set flex_nob.CellPicture = ImageList1.ListImages(11).Picture

'---------------END------------------



Call halfsum
 

vprev = flex_grid.Row

End Sub

Private Sub flex_grid_DblClick()

On Error Resume Next
'back color

Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True



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
'Set flex_nob.CellPicture = ImageList1.ListImages(11).Picture

'---------------END------------------




Unload budgetedduration
Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
 
budgetedduration.cbo_spreadcode.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
budgetedduration.cbo_jobcharge.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
budgetedduration.txt_bdgtdays.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
budgetedduration.txt_per_wrkcmpltd.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
budgetedduration.txt_remarks.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
budgetedduration.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 6)
 
budgetedduration.Show
budgetedduration.Top = 3200
budgetedduration.Left = 0
budgetedduration.Height = 3105
budgetedduration.Width = 5730

budgetedduration.cbo_jobcharge.Enabled = False

vprev = flex_grid.Row


End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BUDGETED DURATION BY SPREAD"
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

Me.Width = 11415
Me.Height = 9750


End Sub
Public Sub flex_title()

On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Spread "
        .ColWidth(1) = 1100
        .ColAlignment(1) = 0
         
        .TextMatrix(0, 2) = "JobCharge"
        .ColWidth(2) = 3500
        .ColAlignment(2) = 0
        
        .TextMatrix(0, 3) = "Bdgt Days"
        .ColWidth(3) = 850
        
         
        .TextMatrix(0, 4) = "%WC"
        .ColWidth(4) = 0
        
        .TextMatrix(0, 5) = "Notes"
        .ColWidth(5) = 6000
        .ColAlignment(5) = 0
         .ColWidth(6) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload budgetedduration
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Caption = "New" Then
budgetedduration.cbo_jobcharge.Enabled = True
If cbo_spr.Text = " " Then
MsgBox "select Spread"
cbo_spr.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False


Unload budgetedduration
budgetedduration.Show
budgetedduration.Top = 3200
budgetedduration.Left = 0
budgetedduration.Height = 3105
budgetedduration.Width = 5730 ' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
If budgetedduration.txt_per_wrkcmpltd.Text = "" Then
budgetedduration.txt_per_wrkcmpltd.Text = 0
End If
If budgetedduration.cbo_spreadcode.Text = "" Then
MsgBox "select Spread"
budgetedduration.cbo_spreadcode.SetFocus
Exit Sub
End If
If budgetedduration.cbo_jobcharge.Text = "" Then
MsgBox "select JobCharge"
budgetedduration.cbo_jobcharge.SetFocus
Exit Sub
End If
If budgetedduration.txt_bdgtdays.Text = "" Then
budgetedduration.txt_bdgtdays.Text = 0
End If
nm = Split(budgetedduration.cbo_spreadcode.Text, "  -  ", Len(budgetedduration.cbo_spreadcode.Text), vbTextCompare)
mm = Split(budgetedduration.cbo_jobcharge.Text, "  -  ", Len(budgetedduration.cbo_jobcharge.Text), vbTextCompare)



Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from budgeteddurationdetails", Cn, 3, 2
sv.AddNew
sv!bdgt_spread_code = nm(0)
sv!bdgt_job_key = mm(0)
sv!bdgt_days = budgetedduration.txt_bdgtdays.Text
sv!bdgt_per_workcomplete = budgetedduration.txt_per_wrkcmpltd.Text
sv!bdgt_remarks = budgetedduration.txt_remarks.Text
sv!t_date = budgetedduration.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv.Update
sv.Close
MsgBox "New Budgeted Duration Added Succesfully"
Call budcost
Unload budgetedduration

Call flex_data
Call flex_title

Exit Sub
assad:
 
   MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
 On Error GoTo assad1

If budgetedduration.txt_per_wrkcmpltd.Text = "" Then
budgetedduration.txt_per_wrkcmpltd.Text = 0
End If
Toolbar1.Buttons(3).Enabled = False
nm = Split(budgetedduration.cbo_spreadcode.Text, "  -  ", Len(budgetedduration.cbo_spreadcode.Text), vbTextCompare)
mm = Split(budgetedduration.cbo_jobcharge.Text, "  -  ", Len(budgetedduration.cbo_jobcharge.Text), vbTextCompare)
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from budgeteddurationdetails where bdgt_id=" & id1, Cn, 3, 2
If Not md.EOF Then
'md!bdgt_spread_code = nm(0)
'md!bdgt_job_key = mm(0)
md!bdgt_days = budgetedduration.txt_bdgtdays.Text
md!bdgt_per_workcomplete = budgetedduration.txt_per_wrkcmpltd.Text
md!bdgt_remarks = budgetedduration.txt_remarks.Text
md!t_date = budgetedduration.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md.Update
md.Close
MsgBox "Selected Budgeted Duration Modified"
End If
Call budcost
Unload budgetedduration

Call flex_data

Call flex_title

Exit Sub
assad1:
 
   MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
gf = Split(flex_grid.TextMatrix(flex_grid.Row, 1), "  -  ", Len(flex_grid.TextMatrix(flex_grid.Row, 1)), vbTextCompare)
bf = Split(flex_grid.TextMatrix(flex_grid.Row, 2), "  -  ", Len(flex_grid.TextMatrix(flex_grid.Row, 2)), vbTextCompare)
Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from cost where bd_spread='" & gf(0) & "' and bd_jobcharge='" & bf(0) & "' and bd_costtype='B'", Cn, 3, 2
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
                                Cn.Execute "delete from budgeteddurationdetails where bdgt_id=" & id2
                                MsgBox "Selected Record Has Been Deleted"
                                Unload budgetedduration
                                Call flex_data
                                Call flex_title
                                Else
                                Unload budgetedduration
                                End If


ElseIf Button.Caption = "Close" Then
Unload Me
Unload budgetedduration
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
AppActivate "FlexGrid To Excel"
For i = 0 To flex_grid.Rows - 1
  flex_grid.Row = i
    For n = 0 To 7
        flex_grid.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flex_grid.Text
    Next
Next
End If






End Sub

Public Sub flex_data()
On Error Resume Next
Dim dys1 As Double
dys1 = 0
nnn = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from budgeteddurationdetails where bdgt_spread_code='" & nnn(0) & "' ", Cn, 3, 2

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
        .TextMatrix(.Rows - 1, 3) = fldata(3)
        dys1 = dys1 + fldata(3)
        .TextMatrix(.Rows - 1, 4) = Format(fldata(4), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = fldata(5)
        .TextMatrix(.Rows - 1, 6) = fldata("t_date")
        fldata.MoveNext
    Wend
End With
Label3.Caption = " Budgeted Days" & " " & Format(dys1, "###,###,##0.00")
End Sub


Public Sub budcost()
nm = Split(budgetedduration.cbo_spreadcode.Text, "  -  ", Len(budgetedduration.cbo_spreadcode.Text), vbTextCompare)
mm = Split(budgetedduration.cbo_jobcharge.Text, "  -  ", Len(budgetedduration.cbo_jobcharge.Text), vbTextCompare)
Dim idddd As Double
idddd = 0
Dim ass As New ADODB.Recordset
If ass.State Then ass.Close
ass.Open "select * from cost  where bd_spread='" & nm(0) & "' and bd_jobcharge= '" & mm(0) & "' and  bd_costtype='B'", Cn, 3, 2
While Not ass.EOF
        If ass!bd_spread <> "NA" Then
                        idddd = ass!bd_ID
                        Dim dys As Double
                        Dim perw As Double
                        dys = 0: perw = 0
                        nh = Split(ass!bd_jobcharge, "  -  ", Len(ass!bd_jobcharge), vbTextCompare)
                        ng = Split(ass!bd_spread, "  -  ", Len(ass!bd_spread), vbTextCompare)
                        Dim bd As New ADODB.Recordset
                        If bd.State Then bd.Close
                        bd.Open "select * from budgeteddurationdetails where bdgt_job_key='" & ass!bd_jobcharge & "' and bdgt_spread_code='" & ass!bd_spread & "'", Cn, 3, 2
                        If Not bd.EOF Then
                        dys = bd!bdgt_days
                        'perw = bd!bdgt_per_workcomplete
                        End If
        
                        Dim fl As New ADODB.Recordset
                        If fl.State Then fl.Close
                        fl.Open "select * from cost where   bd_jobcharge='" & ass!bd_jobcharge & "' and bd_spread='" & ass!bd_spread & "'  and  bd_costtype='B' and bd_id=" & idddd, Cn, 3, 2
                        If Not fl.EOF Then
                                If fl!bd_spread <> "NA" Then
                                fl!bd_days = dys
                                fl!bd_tqty = (fl!bd_qty) * (dys)
                                End If
        
                        fl!bd_extdamt = (fl!bd_xchg) * (fl!bd_unitrate) * (fl!bd_tqty) * ((100 + fl!bd_downtime) / 100) * ((100 + fl!bd_escl) / 100)
                        'fl!bd_wrkcomp = perw
                        fl!bd_bcwpamt = (fl!bd_wrkcomp / 100) * (fl!bd_extdamt)
                        fl.Update
                        End If
        End If
ass.MoveNext
Wend


 

End Sub


Public Sub halfsum()
Dim ko As Double
ko = flex_grid.TextMatrix(flex_grid.Row, 0)

nb = Split(cbo_spr.Text, "  -  ", Len(cbo_spr.Text), vbTextCompare)
Dim dyse As Double
dyse = 0
 


 
    For l = 1 To flex_grid.Row
      
 
        dyse = dyse + flex_grid.TextMatrix(l, 3)
  
    Next l
 Label3.Caption = " Budgeted Days" & " " & dyse
End Sub
