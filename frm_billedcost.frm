VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_billedcost 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Billed Cost"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11175
      Begin VB.TextBox txt_gtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7680
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cbo_pproj 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
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
         Height          =   210
         Left            =   7680
         TabIndex        =   6
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Project"
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
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
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
         TabIndex        =   1
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
            Picture         =   "frm_billedcost.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_billedcost.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8175
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   3
      Cols            =   15
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
End
Attribute VB_Name = "frm_billedcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ds As New ADODB.Recordset

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_pproj_Click()

Call flex_data
End Sub
 
Private Sub flex_grid_Click()
'On Error Resume Next
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture

 
 
 





vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
'On Error Resume Next
'back color
Unload billedcost
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture



'------END---------

billedcost.Show
billedcost.Top = 3200
billedcost.Left = 0
billedcost.Height = 3300
billedcost.Width = 10620
 

Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
 billedcost.cbo_tranx.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
 billedcost.cbo_resc.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
 billedcost.txt_inv.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
 billedcost.DTP_inv.Value = flex_grid.TextMatrix(flex_grid.Row, 4)
 billedcost.cbo_vendor.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
  billedcost.cbo_jobcharge.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
  billedcost.cbo_costcode.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
 billedcost.txt_totdays.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
 billedcost.cbo_uom.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
 billedcost.cbo_curr.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
 billedcost.txt_unitrate = flex_grid.TextMatrix(flex_grid.Row, 11)
 billedcost.txt_Xrate.Text = flex_grid.TextMatrix(flex_grid.Row, 12)
 billedcost.txt_Extdamt.Text = flex_grid.TextMatrix(flex_grid.Row, 13)
 billedcost.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 14)
vprev = flex_grid.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "BILLED COST"
Call flex_data
Call flex_title
Me.Top = 5
Me.Left = 5
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pr.EOF
cbo_pproj.AddItem pr(0) & "  -  " & pr(1)
pr.MoveNext
Wend
pr.Close
Exit Sub
 
 Me.Width = 11415
 Me.Height = 9750
 


End Sub
Public Sub flex_title()
On Error Resume Next

    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "TrnX Type"
        .ColWidth(1) = 500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Resource"
        .ColWidth(2) = 2500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Invoice No"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Inv Date"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Vendor"
        .ColWidth(5) = 3300
        .ColAlignment(5) = 0
         .TextMatrix(0, 6) = "JobCharge"
        .ColWidth(6) = 3300
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "CostCode"
        .ColWidth(7) = 2500
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "Total Qty"
        .ColWidth(8) = 1000
        .TextMatrix(0, 9) = "UOM"
        .ColWidth(9) = 1000
         .TextMatrix(0, 10) = "Curcy"
        .ColWidth(10) = 1000
        .TextMatrix(0, 11) = "Unit Rate"
        .ColWidth(11) = 1000
        .TextMatrix(0, 12) = "XRate"
        .ColWidth(12) = 1000
        .TextMatrix(0, 13) = "Extd Amt(RM)"
        .ColWidth(13) = 1100
        .TextMatrix(0, 14) = " Notes"
        .ColWidth(14) = 4000
        .ColAlignment(14) = 0
         
        
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload billedcost
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
If cbo_pproj.Text = "" Then
MsgBox "select Project"
cbo_pproj.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload billedcost
billedcost.Show
billedcost.Top = 3200
billedcost.Left = 0
billedcost.Height = 3300
billedcost.Width = 10620
' to save new record
ElseIf Button.Caption = "Save" Then

If billedcost.cbo_tranx.Text = "" Then
MsgBox "Select TranX"
billedcost.cbo_tranx.SetFocus
Exit Sub
End If

If billedcost.txt_inv = "" Then
MsgBox "Enter Invoice No."
billedcost.txt_inv.SetFocus
Exit Sub
End If

If billedcost.cbo_vendor.Text = "" Then
MsgBox "Select Vendor"
billedcost.cbo_vendor.SetFocus
Exit Sub
End If

If billedcost.cbo_resc.Text = "" Then
MsgBox "Select Resource"
billedcost.cbo_resc.SetFocus
Exit Sub
End If
 
If billedcost.cbo_jobcharge.Text = "" Then
MsgBox "Select JobCharge"
billedcost.cbo_jobcharge.SetFocus
Exit Sub
End If
 
If billedcost.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
billedcost.cbo_costcode.SetFocus
Exit Sub
End If
 
If billedcost.txt_totdays.Text = "" Then
MsgBox "Enter Days"
billedcost.txt_totdays.SetFocus
Exit Sub
End If

If billedcost.cbo_uom.Text = "" Then
MsgBox "Select UOM"
billedcost.cbo_uom.SetFocus
Exit Sub
End If
 
If billedcost.cbo_curr.Text = "" Then
MsgBox "Select Currency"
billedcost.cbo_curr.SetFocus
Exit Sub
End If
 
If billedcost.txt_unitrate.Text = "" Then
MsgBox "Enter UnitRate"
billedcost.txt_unitrate.SetFocus
Exit Sub
End If
 
If billedcost.txt_Xrate.Text = "" Then
MsgBox "Enter Xrate"
billedcost.txt_Xrate.SetFocus
Exit Sub
End If
 



nn = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
nj = Split(billedcost.cbo_resc.Text, "  -  ", Len(billedcost.cbo_resc.Text), vbTextCompare)
njj = Split(billedcost.cbo_vendor.Text, "  -  ", Len(billedcost.cbo_vendor.Text), vbTextCompare)
njjj = Split(billedcost.cbo_jobcharge.Text, "  -  ", Len(billedcost.cbo_jobcharge.Text), vbTextCompare)
njjjj = Split(billedcost.cbo_costcode.Text, "  -  ", Len(billedcost.cbo_costcode.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from  cost", Cn, 3, 2
sv.AddNew
 
sv!bd_resccode = nj(0)
sv!bd_rescname = nj(1)
 
sv!bd_vendor = njj(0)
sv!bd_projectkey = nn(0)
sv!bd_projectdesc = nn(1)
sv!bd_costtype = "X"
 
sv!bd_tranx = billedcost.cbo_tranx.Text
sv!bd_JobCharge = njjj(0)
sv!bd_costcode = njjjj(0)
sv!bd_tqty = billedcost.txt_totdays.Text
sv!bd_uom = billedcost.cbo_uom.Text
sv!bd_curr = billedcost.cbo_curr.Text
sv!bd_unitrate = billedcost.txt_unitrate
sv!bd_xchg = billedcost.txt_Xrate.Text
sv!bd_extdamt = billedcost.txt_Extdamt.Text
sv!bd_notes = billedcost.txt_notes.Text
sv!bd_inv = billedcost.txt_inv.Text
sv!bd_invdate = billedcost.DTP_inv.Value
sv!t_date = billedcost.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv.Update
sv.Close
MsgBox "New Record Added Succesfully"
Unload billedcost
Call flex_data
Call flex_title
'to modify existing record
ElseIf Button.Caption = "Modify" Then
If billedcost.cbo_tranx.Text = "" Then
MsgBox "Select TranX"
billedcost.cbo_tranx.SetFocus
Exit Sub
End If
If billedcost.txt_inv = "" Then
MsgBox "Enter Invoice No."
billedcost.txt_inv.SetFocus
Exit Sub
End If
If billedcost.cbo_vendor.Text = "" Then
MsgBox "Select Vendor"
billedcost.cbo_vendor.SetFocus
Exit Sub
End If
If billedcost.cbo_resc.Text = "" Then
MsgBox "Select Resource"
billedcost.cbo_resc.SetFocus
Exit Sub
End If
If billedcost.cbo_jobcharge.Text = "" Then
MsgBox "Select JobCharge"
billedcost.cbo_jobcharge.SetFocus
Exit Sub
End If
If billedcost.cbo_costcode.Text = "" Then
MsgBox "Select CostCode"
billedcost.cbo_costcode.SetFocus
Exit Sub
End If
If billedcost.txt_totdays.Text = "" Then
MsgBox "Enter Days"
billedcost.txt_totdays.SetFocus
Exit Sub
End If
If billedcost.cbo_uom.Text = "" Then
MsgBox "Select UOM"
billedcost.cbo_uom.SetFocus
Exit Sub
End If
If billedcost.cbo_curr.Text = "" Then
MsgBox "Select Currency"
billedcost.cbo_curr.SetFocus
Exit Sub
End If
If billedcost.txt_unitrate.Text = "" Then
MsgBox "Enter UnitRate"
billedcost.txt_unitrate.SetFocus
Exit Sub
End If
If billedcost.txt_Xrate.Text = "" Then
MsgBox "Enter Xrate"
billedcost.txt_Xrate.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = False
nn = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
nn = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
nj = Split(billedcost.cbo_resc.Text, "  -  ", Len(billedcost.cbo_resc.Text), vbTextCompare)
njj = Split(billedcost.cbo_vendor.Text, "  -  ", Len(billedcost.cbo_vendor.Text), vbTextCompare)
njjj = Split(billedcost.cbo_jobcharge.Text, "  -  ", Len(billedcost.cbo_jobcharge.Text), vbTextCompare)
njjjj = Split(billedcost.cbo_costcode.Text, "  -  ", Len(billedcost.cbo_costcode.Text), vbTextCompare)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from  cost where bd_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!bd_resccode = nj(0)
md!bd_rescname = nj(1)
md!bd_vendor = njj(0)
md!bd_projectkey = nn(0)
md!bd_projectdesc = nn(1)
md!bd_costtype = "X"
md!bd_tranx = billedcost.cbo_tranx.Text
md!bd_JobCharge = njjj(0)
md!bd_costcode = njjjj(0)
md!bd_tqty = billedcost.txt_totdays.Text
md!bd_uom = billedcost.cbo_uom.Text
md!bd_curr = billedcost.cbo_curr.Text
md!bd_unitrate = billedcost.txt_unitrate
md!bd_xchg = billedcost.txt_Xrate.Text
md!bd_extdamt = billedcost.txt_Extdamt.Text
md!bd_notes = billedcost.txt_notes.Text
md!bd_inv = billedcost.txt_inv.Text
md!bd_invdate = billedcost.DTP_inv.Value
md!t_date = billedcost.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md.Update
md.Close
MsgBox "Selected Record Modified"
End If

Unload billedcost
Call flex_data
Call flex_title

'to delete
ElseIf Button.Caption = "Delete" Then

dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from  cost where bd_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload billedcost
Call flex_data
Call flex_title
Else
Unload billedcost
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload billedcost
End If
End Sub
Public Sub flex_data()
'On Error Resume Next
If cbo_pproj.Text = "" Then Exit Sub
rscc = Split(cbo_pproj.Text, "  -  ", Len(cbo_pproj.Text), vbTextCompare)
Dim gtotal As Double
gtotal = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from  cost where bd_projectkey='" & rscc(0) & "'   and bd_costtype='X' ", Cn, 3, 2
With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        .TextMatrix(.Rows - 1, 1) = fldata!bd_tranx
        Dim rc As New ADODB.Recordset
        If rc.State Then rc.Close
        rc.Open "select Distinct(resc_desc) from resourcemaster where resc_code='" & fldata!bd_resccode & "' ", Cn, 3, 2
        If Not rc.EOF Then
          .TextMatrix(.Rows - 1, 2) = fldata!bd_resccode & "  -  " & rc(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata!bd_resccode
        End If
        .TextMatrix(.Rows - 1, 3) = fldata!bd_inv
        .TextMatrix(.Rows - 1, 4) = fldata!bd_invdate
        Dim jl As New ADODB.Recordset
        If jl.State Then jl.Close
        jl.Open "select DISTINCT(vendor_desc) from vendormaster where vendor_code='" & fldata!bd_vendor & "' ", Cn, 3, 2
        If Not jl.EOF Then
        .TextMatrix(.Rows - 1, 5) = fldata!bd_vendor & "  -  " & jl(0)
        Else
        .TextMatrix(.Rows - 1, 5) = fldata!bd_vendor
        End If
        jl.Close
        Dim jcg As New ADODB.Recordset
        If jcg.State Then jcg.Close
        jcg.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_JobCharge & "' ", Cn, 3, 2
        If Not jcg.EOF Then
        .TextMatrix(.Rows - 1, 6) = fldata!bd_JobCharge & "  -  " & jcg(0)
        Else
        .TextMatrix(.Rows - 1, 6) = fldata!bd_JobCharge
        End If
        jcg.Close
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 7) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 7) = fldata!bd_costcode
        End If
        cs.Close
        .TextMatrix(.Rows - 1, 8) = Format(fldata!bd_tqty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 9) = fldata!bd_uom
        .TextMatrix(.Rows - 1, 10) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 11) = Format(fldata!bd_unitrate, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 12) = Format(fldata!bd_xchg, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 13) = Format(fldata!bd_extdamt, "###,###,##0.00")
        gtotal = gtotal + fldata!bd_extdamt
        .TextMatrix(.Rows - 1, 14) = fldata!bd_notes
         
        
        fldata.MoveNext
    Wend
End With
Txt_gtotal.Text = Format(gtotal, "###,###,##0.00")
vprev = 0
End Sub

