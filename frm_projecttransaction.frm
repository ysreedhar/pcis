VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_projecttransaction 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REVENUE @ PROJECTKEY LEVEL"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   15901
      _Version        =   393216
      Rows            =   3
      Cols            =   11
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
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Picture         =   "frm_projecttransaction.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_projecttransaction.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_projecttransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub flex_grid_Click()
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True

On Error Resume Next
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


 

vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True

On Error Resume Next
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



'------END---------


Unload projecttransaction
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)
 
projecttransaction.cbo_projkey.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
projecttransaction.txt_projdesc.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
projecttransaction.txt_lye_revn.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
projecttransaction.txt_lye_cost.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
projecttransaction.txt_lme_revn.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
projecttransaction.txt_lme_cost.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
projecttransaction.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 9)
 projecttransaction.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
projecttransaction.txt_lye_revn1.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
projecttransaction.txt_lme_revn1.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
projecttransaction.Show
projecttransaction.Top = 3200
projecttransaction.Left = 0
projecttransaction.Height = 4935
projecttransaction.Width = 5595


 
vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "REVENUE @ PROJECTKEY LEVEL"
Call flex_title
Call flex_data
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()

On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Project Key"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Project TranX Desc"
        .ColWidth(2) = 2500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "PTD_LYE_Revn(B)"
        .ColWidth(3) = 1600
         
        .TextMatrix(0, 4) = "PTD-LYE-Cost"
        .ColWidth(4) = 0
         
        .TextMatrix(0, 5) = "YTD_LME_Revn(B)"
        .ColWidth(5) = 1600
        .TextMatrix(0, 6) = "YTD-LME-cost"
        .ColWidth(6) = 0
        .ColWidth(9) = 0
        .TextMatrix(0, 10) = "MAIN/CO"
        .ColWidth(10) = 1000
        .ColAlignment(10) = 0
        .TextMatrix(0, 7) = "PTD_LYE_Revn(UB)"
        .ColWidth(7) = 1600
        .TextMatrix(0, 8) = "YTD_LME_Revn(UB))"
        .ColWidth(8) = 1600
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload projecttransaction
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload projecttransaction
projecttransaction.Show
projecttransaction.Top = 3200
projecttransaction.Left = 90
projecttransaction.Height = 4935
projecttransaction.Width = 5595
' to save new record
ElseIf Button.Caption = "Save" Then
If projecttransaction.cbo_projkey.Text = "" Then
MsgBox "Select Project"
projecttransaction.cbo_projkey.SetFocus
Exit Sub
End If
If projecttransaction.txt_projdesc.Text = "" Then
projecttransaction.txt_projdesc.Text = "-"
End If
If projecttransaction.txt_lye_revn.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lye_revn.SetFocus
Exit Sub
End If
If projecttransaction.txt_lye_revn1.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lye_revn1.SetFocus
Exit Sub
End If
If projecttransaction.txt_lme_revn.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lme_revn.SetFocus
Exit Sub
End If
If projecttransaction.txt_lme_revn1.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lme_revn1.SetFocus
Exit Sub
End If
 


nn = Split(projecttransaction.cbo_projkey.Text, "  -  ", Len(projecttransaction.cbo_projkey.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from projecttransaction", Cn, 3, 2
sv.AddNew
sv!pk_projkey = nn(0)
sv!pk_projdesc = nn(1)
sv!ptd_lye_revn = projecttransaction.txt_lye_revn.Text
sv!ptd_lye_cost = projecttransaction.txt_lye_cost.Text
sv!ytd_lme_revn = projecttransaction.txt_lme_revn.Text
sv!ytd_lme_cost = projecttransaction.txt_lme_cost.Text
sv!ptd_lye_revn1 = projecttransaction.txt_lye_revn1.Text
sv!ytd_lme_revn1 = projecttransaction.txt_lme_revn1.Text
sv!t_date = projecttransaction.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!notes = projecttransaction.txt_notes.Text
sv.Update
sv.Close
MsgBox "New Project Transaction Added Succesfully"
Unload projecttransaction
Call flex_data
Call flex_title
'to modify existing record
ElseIf Button.Caption = "Modify" Then
If projecttransaction.cbo_projkey.Text = "" Then
MsgBox "Select Project"
projecttransaction.cbo_projkey.SetFocus
Exit Sub
End If
If projecttransaction.txt_projdesc.Text = "" Then
projecttransaction.txt_projdesc.Text = "-"
End If
If projecttransaction.txt_lye_revn.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lye_revn.SetFocus
Exit Sub
End If
If projecttransaction.txt_lye_revn1.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lye_revn1.SetFocus
Exit Sub
End If
If projecttransaction.txt_lme_revn.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lme_revn.SetFocus
Exit Sub
End If
If projecttransaction.txt_lme_revn1.Text = "" Then
MsgBox "Enter Amount"
projecttransaction.txt_lme_revn1.SetFocus
Exit Sub
End If
nn = Split(projecttransaction.cbo_projkey.Text, "  -  ", Len(projecttransaction.cbo_projkey.Text), vbTextCompare)
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from projecttransaction where pk_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!pk_projkey = nn(0)
md!pk_projdesc = nn(1)
md!ptd_lye_revn = projecttransaction.txt_lye_revn.Text
md!ptd_lye_cost = projecttransaction.txt_lye_cost.Text
md!ytd_lme_revn = projecttransaction.txt_lme_revn.Text
md!ytd_lme_cost = projecttransaction.txt_lme_cost.Text
md!ptd_lye_revn1 = projecttransaction.txt_lye_revn1.Text
md!ytd_lme_revn1 = projecttransaction.txt_lme_revn1.Text
md!t_date = projecttransaction.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md!notes = projecttransaction.txt_notes.Text
md.Update
md.Close
MsgBox "Selected Project transaction Modified"
End If

Unload projecttransaction
Call flex_data
Call flex_title

'to delete
ElseIf Button.Caption = "Delete" Then
Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from projecttransaction where pk_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload projecttransaction
Call flex_data
Call flex_title
Else
Unload projecttransaction
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload projecttransaction
End If




End Sub

Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from projecttransaction p , userproject u where p.pk_projkey=u.project and u.username='" & main.Label2.Caption & "' ", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata(1)
        .TextMatrix(.Rows - 1, 2) = fldata(2)
        .TextMatrix(.Rows - 1, 3) = Format(fldata(3), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 4) = Format(fldata(4), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = Format(fldata(5), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 6) = Format(fldata(6), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 9) = fldata("t_date")
        .TextMatrix(.Rows - 1, 10) = fldata("notes")
        .TextMatrix(.Rows - 1, 7) = Format(fldata("ptd_lye_revn1"), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 8) = Format(fldata("ytd_lme_revn1"), "###,###,##0.00")
        fldata.MoveNext
    Wend
End With
End Sub


