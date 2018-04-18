VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_l0 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OTHER INC/EXP & OVERHEAD-EST/RECOVERY"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10575
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_year 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15266
      _Version        =   393216
      Rows            =   3
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
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
      Width           =   11535
      _ExtentX        =   20346
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
      Left            =   2400
      Top             =   120
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
            Picture         =   "frm_l0.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_l0.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frm_l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public flg As Integer

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_year_Click()
Call flex_data
Call flex_title
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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture
 


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
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'--END---------



Dim ID As Double
ID = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
ID = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim mdd As New ADODB.Recordset
If mdd.State Then mdd.Close
mdd.Open "select * from oitranx where oi_id=" & ID, Cn, 3, 2
If Not mdd.EOF Then
 If mdd!flg = 1 Then
 flg = 1
 Unload oitran
oitran.Show
oitran.Top = 3200
oitran.Left = 0
oitran.Height = 6030
oitran.Width = 7815

cbo_year.Text = mdd!oi_Year
        Dim otr1 As New ADODB.Recordset
        If otr1.State Then otr1.Close
        otr1.Open "select * from othertransaction where ot_desc='" & mdd!tranx & "' ", Cn, 3, 2
        If Not otr1.EOF Then
       oitran.txt_tranx.Text = otr1!ot_tranx & "  -  " & mdd!tranx
        Else
       oitran.txt_tranx.Text = mdd!tranx
        End If
'oitran.txt_tranx.Text = mdd!tranx
oitran.txt_bdgt.Text = Format(mdd!bdgt, "###,###,###,##0.00")
oitran.txt_bcwpbl.Text = Format(mdd!bcwpbl, "###,###,###,##0.00")
oitran.txt_bcwpdays.Text = Format(mdd!bcwpdays, "###,###,###,##0.00")
oitran.txt_etcbl.Text = Format(mdd!etcbl, "###,###,###,##0.00")
oitran.txt_etcdays.Text = Format(mdd!etcdays, "###,###,###,##0.00")
 
oitran.txt_acwpacc.Text = Format(mdd!acwpacc, "###,###,###,##0.00")
oitran.txt_acwpbl.Text = Format(mdd!acwpbl, "###,###,###,##0.00")
oitran.txt_acwpadj.Text = Format(mdd!acwpadj, "###,###,###,##0.00")
oitran.txt_eac.Text = Format(mdd!eac, "###,###,###,##0")
oitran.txt_bcwp.Text = Format(mdd!bcwp, "###,###,###,##0")
oitran.txt_acwp.Text = Format(mdd!acwp, "###,###,###,##0")
oitran.txt_etc.Text = Format(mdd!etc, "###,###,###,##0")
oitran.txt_ytd.Text = Format(mdd!ytd, "###,###,###,##0")
oitran.txt_ctd.Text = Format(mdd!ctd, "###,###,###,##0")
oitran.txt_chg.Text = Format(mdd!chng, "###,###,###,##0")
oitran.txt_adjustment.Text = Format((mdd!ectcadj), "###,###,###,##0.00")
oitran.dtp_asat.Value = mdd!asatdate
main.DTPcutdate1.Value = mdd!ctdate

oitran.txt_adjbl.Text = mdd!adjbl
oitran.txt_rateb4.Text = mdd!rateb4
oitran.txt_rateaft.Text = mdd!rateaft


dys = 0
 
dys = main.DTPcutdate1.Value - oitran.dtp_asat.Value
  
 ppr = 0
Dim prw As New ADODB.Recordset
If prw.State Then prw.Close
prw.Open "select * from parameters", Cn, 3, 2
If Not pr.EOF Then
ppr = prw!p_ydays
ppsd = prw!p_sdate
pped = prw!p_edate
'ppcd = pr!p_cdate
End If
ppcd = main.DTPcutdate1.Value
ElseIf mdd!flg = 2 Then
flg = 2
Unload oitran1
oitran1.Show
oitran1.Top = 3200
oitran1.Left = 0
oitran1.Height = 6030
oitran1.Width = 7815
cbo_year.Text = mdd!oi_Year
        Dim otr2 As New ADODB.Recordset
        If otr2.State Then otr2.Close
        otr2.Open "select * from othertransaction where ot_desc='" & mdd!tranx & "' ", Cn, 3, 2
        If Not otr2.EOF Then
       oitran1.txt_tranx.Text = otr2!ot_tranx & "  -  " & mdd!tranx
        Else
       oitran1.txt_tranx.Text = mdd!tranx
        End If
 oitran1.txt_tranx.Enabled = False
'oitran1.txt_tranx.Text = mdd!tranx
oitran1.txt_bdgt.Text = Format(mdd!bdgt, "###,###,###,##0.00")
oitran1.txt_eac.Text = Format(mdd!eac, "###,###,###,##0")
oitran1.txt_bcwp.Text = Format(mdd!bcwp, "###,###,###,##0")
oitran1.txt_acwp.Text = Format(mdd!acwp, "###,###,###,##0")
oitran1.txt_etc.Text = Format(mdd!etc, "###,###,###,##0")
oitran1.txt_ytd.Text = Format(mdd!ytd, "###,###,###,##0")
oitran1.txt_ctd.Text = Format(mdd!ctd, "###,###,###,##0")
oitran1.txt_chg.Text = Format(mdd!chng, "###,###,###,##0")
main.DTPcutdate1.Value = mdd!ctdate
End If
End If
''oitran.Label6.Caption = ""
''oitran.Label6.Caption = "(B/L Bdgt)/" & ppr
''oitran.Label3.Caption = ""
''oitran.Label3.Caption = "((B/L)/" & ppr & ") * " & dys
''oitran.Label15.Caption = ""
''oitran.Label15.Caption = "(B/L Bdgt)/" & ppr
vprev = flex_grid.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "OTHER INC/EXP & OVERHEAD-EST/RECOVERY"
Call flex_title
Call flex_data
fab = 1
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5

Dim j As Integer
j = 0
For j = 2000 To 2050
cbo_year.AddItem j
Next j
 Me.Width = 11415
 Me.Height = 9750
  
End Sub
Public Sub flex_title()
On Error Resume Next
   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Transaction"
        .ColWidth(1) = 2500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "BDGT"
        .ColWidth(2) = 1100
        .TextMatrix(0, 3) = "EAC"
        .ColWidth(3) = 1100
        .TextMatrix(0, 4) = "BCWP"
        .ColWidth(4) = 1100
        .TextMatrix(0, 5) = "ACWP"
        .ColWidth(5) = 1100
        .TextMatrix(0, 6) = "ETC"
        .ColWidth(6) = 1100
        .TextMatrix(0, 7) = "YTD-LME"
        .ColWidth(7) = 1100
        .TextMatrix(0, 8) = "CTD"
        .ColWidth(8) = 1100
        .TextMatrix(0, 9) = "Chg-CurrMonth"
        .ColWidth(9) = 1100
        .TextMatrix(0, 10) = "Notes"
        .ColWidth(10) = 4000
        .ColAlignment(10) = 0
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload oitran
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
 If cbo_year.Text = "" Then
 MsgBox "Select Year"
 cbo_year.SetFocus
 Exit Sub
 End If
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
xx = MsgBox("For Actual Click YES .For Recovery Click NO", vbYesNo)
 If xx = vbYes Then
 flg = 1
Unload oitran
oitran.Show
oitran.Top = 3200
oitran.Left = 0
oitran.Height = 6030
oitran.Width = 7815
ElseIf xx = vbNo Then
flg = 2
Unload oitran1
oitran1.Show
oitran1.Top = 3200
oitran1.Left = 0
oitran1.Height = 6030
oitran1.Width = 7815

End If
' to save new record
ElseIf Button.Caption = "Save" Then
 'validate
 If cbo_year.Text = "" Then
 MsgBox "Select Year"
 cbo_year.SetFocus
 Exit Sub
 End If
 
 
 If flg = 1 Then ''flag
 
If oitran.txt_tranx.Text = "" Then
MsgBox "Select TranX"
oitran.txt_tranx.SetFocus
Exit Sub
End If
If oitran.txt_bdgt.Text = "" Then
oitran.txt_bdgt.Text = 0

End If
If oitran.txt_eac.Text = "" Then
oitran.txt_eac.Text = 0

End If
If oitran.txt_bcwp.Text = "" Then
oitran.txt_bcwp.Text = 0
End If

If oitran.txt_acwp.Text = "" Then
oitran.txt_acwp.Text = 0
End If
If oitran.txt_etc.Text = "" Then
oitran.txt_etc.Text = 0
End If
If oitran.txt_ytd.Text = "" Then
oitran.txt_ytd.Text = 0
End If
If oitran.txt_ctd.Text = "" Then
oitran.txt_ctd.Text = 0
End If
If oitran.txt_chg.Text = "" Then
oitran.txt_chg.Text = 0
End If
If oitran.txt_adjustment.Text = "" Then
oitran.txt_adjustment.Text = 0
End If
 
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from oitranx", Cn, 3, 2
sv.AddNew
jl = Split(oitran.txt_tranx.Text, "  -  ", Len(oitran.txt_tranx.Text), vbTextCompare)
sv!oi_Year = cbo_year.Text
sv!tranx = jl(1)
sv!bdgt = oitran.txt_bdgt.Text
sv!bcwpbl = oitran.txt_bcwpbl.Text
sv!bcwpdays = oitran.txt_bcwpdays.Text
sv!etcbl = oitran.txt_etcbl.Text
sv!etcdays = oitran.txt_etcdays.Text
 
sv!acwpacc = oitran.txt_acwpacc.Text
sv!acwpbl = oitran.txt_acwpbl.Text
sv!acwpadj = oitran.txt_acwpadj.Text
sv!eac = oitran.txt_eac.Text
sv!bcwp = oitran.txt_bcwp.Text
sv!acwp = oitran.txt_acwp.Text
sv!etc = oitran.txt_etc.Text
sv!ytd = oitran.txt_ytd.Text
sv!ctd = oitran.txt_ctd.Text
sv!chng = oitran.txt_chg.Text

sv!asatdate = oitran.dtp_asat.Value
sv!ctdate = main.DTPcutdate1.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!notes = oitran.txt_notes.Text
sv!adjbl = oitran.txt_adjbl.Text
sv!rateb4 = oitran.txt_rateb4.Text
sv!rateaft = oitran.txt_rateaft.Text
sv!ectcadj = oitran.txt_adjustment.Text
sv!flg = 1
sv.Update
sv.Close

 Call flex_data
Call flex_title

ElseIf flg = 2 Then

 On Error GoTo assad3
 
If oitran1.txt_bdgt.Text = "" Then
oitran1.txt_bdgt.Text = 0

End If
If oitran1.txt_eac.Text = "" Then
oitran1.txt_eac.Text = 0

End If
If oitran1.txt_bcwp.Text = "" Then
oitran1.txt_bcwp.Text = 0
End If

If oitran1.txt_acwp.Text = "" Then
oitran1.txt_acwp.Text = 0
End If
If oitran1.txt_etc.Text = "" Then
oitran1.txt_etc.Text = 0
End If

 
Dim sva As New ADODB.Recordset
If sva.State Then sv.Close
sva.Open "select * from oitranx", Cn, 3, 2
sva.AddNew
jl = Split(oitran1.txt_tranx.Text, "  -  ", Len(oitran1.txt_tranx.Text), vbTextCompare)
sva!oi_Year = cbo_year.Text
sva!tranx = jl(1)
sva!bdgt = oitran1.txt_bdgt.Text
sva!bcwpbl = 0
sva!bcwpdays = 0
sva!etcbl = 0
sva!etcdays = 0
 
sva!acwpacc = 0
sva!acwpbl = 0
sva!acwpadj = 0
sva!eac = oitran1.txt_eac.Text
sva!bcwp = oitran1.txt_bcwp.Text
sva!acwp = oitran1.txt_acwp.Text
sva!etc = oitran1.txt_etc.Text
sva!ytd = oitran1.txt_ytd.Text
sva!ctd = oitran1.txt_ctd.Text
sva!chng = oitran1.txt_chg.Text
sva!ctdate = main.DTPcutdate1.Value
sva!u_date = Now
sva!t_user = main.Label2.Caption
sva!notes = oitran1.txt_notes.Text
sva!adjbl = 0
sva!rateb4 = 0
sva!rateaft = 0
sva!ectcadj = 0
sva!flg = 2
sva.Update
sva.Close

 
End If 'flag
MsgBox "New Record Added Succesfully"
Unload oitran
Call flex_data
Call flex_title
Exit Sub
assad3:
'
'       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1
If flg = 1 Then


If oitran.txt_tranx.Text = "" Then
MsgBox "Select TranX"
oitran.txt_tranx.SetFocus
Exit Sub
End If
If oitran.txt_bdgt.Text = "" Then
oitran.txt_bdgt.Text = 0

End If
If oitran.txt_eac.Text = "" Then
oitran.txt_eac.Text = 0

End If
If oitran.txt_bcwp.Text = "" Then
oitran.txt_bcwp.Text = 0
End If

If oitran.txt_acwp.Text = "" Then
oitran.txt_acwp.Text = 0
End If
If oitran.txt_etc.Text = "" Then
oitran.txt_etc.Text = 0
End If
If oitran.txt_ytd.Text = "" Then
oitran.txt_ytd.Text = 0
End If
If oitran.txt_ctd.Text = "" Then
oitran.txt_ctd.Text = 0
End If
If oitran.txt_chg.Text = "" Then
oitran.txt_chg.Text = 0
End If
If oitran.txt_adjustment.Text = "" Then
oitran.txt_adjustment.Text = 0
End If
 
Toolbar1.Buttons(3).Enabled = False
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
jl2 = Split(oitran.txt_tranx.Text, "  -  ", Len(oitran.txt_tranx.Text), vbTextCompare)

Dim md1 As New ADODB.Recordset
If md1.State Then md1.Close
md1.Open "select * from oitranx where oi_id=" & id2, Cn, 3, 2
If Not md1.EOF Then
md1!oi_Year = cbo_year.Text
md1!tranx = jl2(1)
md1!bdgt = oitran.txt_bdgt.Text
md1!bcwpbl = oitran.txt_bcwpbl.Text
md1!bcwpdays = oitran.txt_bcwpdays.Text
md1!etcbl = oitran.txt_etcbl.Text
md1!etcdays = oitran.txt_etcdays.Text
 
md1!acwpacc = oitran.txt_acwpacc.Text
md1!acwpbl = oitran.txt_acwpbl.Text
md1!acwpadj = oitran.txt_acwpadj.Text
md1!eac = oitran.txt_eac.Text
md1!bcwp = oitran.txt_bcwp.Text
md1!acwp = oitran.txt_acwp.Text
md1!etc = oitran.txt_etc.Text
md1!ytd = oitran.txt_ytd.Text
md1!ctd = oitran.txt_ctd.Text
md1!chng = oitran.txt_chg.Text

md1!asatdate = oitran.dtp_asat.Value
md1!ctdate = main.DTPcutdate1.Value
md1!u_date = Now
md1!t_user = main.Label2.Caption
md1!notes = oitran.txt_notes.Text
md1!adjbl = oitran.txt_adjbl.Text
md1!rateb4 = oitran.txt_rateb4.Text
md1!rateaft = oitran.txt_rateaft.Text
md1!ectcadj = oitran.txt_adjustment.Text
md1!flg = 1
md1.Update
md1.Close
End If
Call flex_data
Call flex_title
assad1:
'
'       MsgBox "Duplicate Entries Not Allowed"
ElseIf flg = 2 Then
'validate
On Error GoTo assad4
 If cbo_year.Text = "" Then
 MsgBox "Select Year"
 cbo_year.SetFocus
 Exit Sub
 End If
 
If oitran1.txt_bdgt.Text = "" Then
oitran1.txt_bdgt.Text = 0

End If
If oitran1.txt_eac.Text = "" Then
oitran1.txt_eac.Text = 0

End If
If oitran1.txt_bcwp.Text = "" Then
oitran1.txt_bcwp.Text = 0
End If

If oitran1.txt_acwp.Text = "" Then
oitran1.txt_acwp.Text = 0
End If
If oitran1.txt_etc.Text = "" Then
oitran1.txt_etc.Text = 0
End If

Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
jl2 = Split(oitran1.txt_tranx.Text, "  -  ", Len(oitran1.txt_tranx.Text), vbTextCompare)

Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from oitranx where oi_id=" & id1, Cn, 3, 2
If Not md.EOF Then
 

md!oi_Year = cbo_year.Text
md!tranx = jl2(1)
md!bdgt = oitran1.txt_bdgt.Text
md!bcwpbl = 0
md!bcwpdays = 0
md!etcbl = 0
md!etcdays = 0
 
md!acwpacc = 0
md!acwpbl = 0
md!acwpadj = 0
md!eac = oitran1.txt_eac.Text
md!bcwp = oitran1.txt_bcwp.Text
md!acwp = oitran1.txt_acwp.Text
md!etc = oitran1.txt_etc.Text
md!ytd = oitran1.txt_ytd.Text
md!ctd = oitran1.txt_ctd.Text
md!chng = oitran1.txt_chg.Text
md!asatdate = 0
md!ctdate = main.DTPcutdate1.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md!notes = oitran1.txt_notes.Text
md!adjbl = 0
md!rateb4 = 0
md!rateaft = 0
md!ectcadj = 0
md!flg = 2
md.Update
md.Close
MsgBox "Selected Transaction Modified"
End If
 
 
Unload oitran1
Call flex_data
Call flex_title
Exit Sub
assad4:
'
'       MsgBox "Duplicate Entries Not Allowed"
       End If
'to delete
ElseIf Button.Caption = "Delete" Then
Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id3 As Double
id3 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id3 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from oitranx where oi_id=" & id3
MsgBox "Selected Transaction Has Been Deleted"
Unload oitran

Call flex_data
Call flex_title
Else
Unload oitran
Unload oitran1
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload oitran
Unload oitran1
End If


 
 
End Sub

Public Sub flex_data()
On Error Resume Next
If cbo_year.Text = "" Then Exit Sub
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select ot.oi_id,oth.ot_tranx,ot.tranx,ot.bdgt,ot.eac,ot.bcwp,ot.acwp,ot.etc,ot.ytd,ot.ctd,ot.chng,ot.notes,ot.asatdate from oitranx ot , othertransaction oth where ot.tranx=oth.ot_desc and  ot.oi_year='" & cbo_year.Text & "' order by oth.ot_tranx ", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
         
'        Dim otr As New ADODB.Recordset
'        If otr.State Then otr.Close
'        otr.Open "select * from othertransaction where ot_desc='" & fldata!tranx & "' ", Cn, 3, 2
'        If Not otr.EOF Then
'        .TextMatrix(.Rows - 1, 1) = otr!ot_tranx & "  -  " & fldata!tranx
'        Else
        .TextMatrix(.Rows - 1, 1) = fldata(1) & "  -  " & fldata(2)
'        End If
        .TextMatrix(.Rows - 1, 2) = Format(fldata(3), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 3) = Format(fldata(4), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 4) = Format(fldata(5), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 5) = Format(fldata(6), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 6) = Format(fldata(7), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 7) = Format(fldata(8), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 8) = Format(fldata(9), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 9) = Format(fldata(10), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 10) = Format(fldata(11), "###,###,###,##0")
        .TextMatrix(.Rows - 1, 11) = Format(fldata(12), "###,###,###,##0")
        fldata.MoveNext
    Wend
End With
End Sub



