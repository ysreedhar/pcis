VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_costcode 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cost Code"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   15901
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   14450266
      ForeColorFixed  =   16777215
      ForeColorSel    =   -2147483639
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
      Width           =   11070
      _ExtentX        =   19526
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
            Picture         =   "frm_costcode.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_costcode.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_costcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub flex_grid_Click()
 

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

Dim j As Integer
For j = 0 To costcode.List1.ListCount - 1
costcode.List1.Selected(j) = False
Next

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




On Error Resume Next
Unload costcode
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from costcode where cc_id=" & id, Cn, 3, 2
If Not sh.EOF Then
costcode.txt_costcode.Text = sh!cc_code
costcode.txt_Desc.Text = sh!cc_desc
costcode.DTP_tdate.Value = sh!t_date
costcode.txt_notes.Text = sh!notes
End If
Dim ls As New ADODB.Recordset
If ls.State Then ls.Close
ls.Open "select * from resourcecostcode where rcc_id=" & id, Cn, 3, 2
While Not ls.EOF
Dim rc As String
rc = ls!rcc_resource

Dim i1 As Integer
For i1 = 0 To costcode.List1.ListCount - 1

jr = Split(costcode.List1.List(i1), "  -  ", Len(costcode.List1.List(i1)), vbTextCompare)
 
If jr(0) = rc Then costcode.List1.Selected(i1) = True
Next


ls.MoveNext
Wend

costcode.Show
costcode.Top = 3200
costcode.Left = 0
costcode.Height = 4455
costcode.Width = 8670
sh.Close

vprev = flex_grid.Row
costcode.txt_costcode.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "COST CODE"
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
        
        .TextMatrix(0, 1) = "Cost Code"
        .ColWidth(1) = 1000
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "CostCode Description"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Notes"
        .ColWidth(3) = 6600
        .ColAlignment(3) = 0
        .ColWidth(0) = 0
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload costcode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

If Button.Caption = "New" Then
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload costcode
costcode.Show
costcode.Top = 3200
costcode.Left = 0
costcode.Height = 4455
costcode.Width = 8670
costcode.txt_costcode.Enabled = True
' to save new record
ElseIf Button.Caption = "Save" Then
 On Error GoTo assad
If costcode.txt_costcode.Text = "" Then
MsgBox "Enter Cost Code"
costcode.txt_costcode.SetFocus
Exit Sub
End If
If costcode.List1.SelCount = 0 Then
MsgBox "select Resource Code"
costcode.List1.SetFocus
Exit Sub
End If
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from costcode", Cn, 3, 2
sv.AddNew
sv!cc_code = costcode.txt_costcode.Text
sv!cc_desc = costcode.txt_Desc.Text
sv!t_date = costcode.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!notes = costcode.txt_notes.Text
sv.Update
sv.Close
MsgBox "New CostCode Added Succesfully"
Dim rsc As New ADODB.Recordset
If rsc.State Then rsc.Close
rsc.Open "select MAX(cc_id) from costcode", Cn, 3, 2
Dim mx As Integer
mx = 0
If Not rsc.EOF Then
mx = rsc(0)
End If
Dim rt As New ADODB.Recordset
If rt.State Then rt.Close
rt.Open "select * from resourcecostcode", Cn, 3, 2

Dim k As Integer

For k = 0 To costcode.List1.ListCount - 1
If costcode.List1.Selected(k) = True Then
rt.AddNew
rt!rcc_id = mx
jj = Split(costcode.List1.List(k), "  -  ", Len(costcode.List1.List(k)), vbTextCompare)
rt!rcc_resource = jj(0)
rt!t_date = costcode.DTP_tdate.Value
rt!u_date = Now
rt!t_user = main.Label2.Caption
rt.Update
End If
Next





Unload costcode
Call flex_data
Call flex_title

Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then

On Error GoTo assad1
If costcode.txt_costcode.Text = "" Then
MsgBox "Enter Cost Code"
costcode.txt_costcode.SetFocus
Exit Sub
End If
If costcode.List1.SelCount = 0 Then
MsgBox "Select Resource Code"
costcode.List1.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
 
'Cn.Execute "delete from costcode where cc_id=" & id1
Cn.Execute "delete from resourcecostcode where rcc_id=" & id1


Dim sv1 As New ADODB.Recordset
If sv1.State Then sv1.Close
sv1.Open "select * from costcode where cc_id=" & id1, Cn, 3, 2
If Not sv1.EOF Then
 
sv1!cc_desc = costcode.txt_Desc.Text
sv1!t_date = costcode.DTP_tdate.Value
sv1!u_date = Now
sv1!t_user = main.Label2.Caption
sv1!notes = costcode.txt_notes.Text
sv1.Update
sv1.Close
End If

'Dim rsc1 As New ADODB.Recordset
'If rsc1.State Then rsc1.Close
'rsc1.Open "select MAX(cc_id) from costcode", Cn, 3, 2
'Dim mx1 As Integer
'mx1 = 0
'If Not rsc1.EOF Then
'mx1 = rsc1(0)
'End If
Dim rt1 As New ADODB.Recordset
If rt1.State Then rt1.Close
rt1.Open "select * from resourcecostcode", Cn, 3, 2

Dim k1 As Integer

For k1 = 0 To costcode.List1.ListCount - 1
If costcode.List1.Selected(k1) = True Then
rt1.AddNew
rt1!rcc_id = id1
jjj = Split(costcode.List1.List(k1), "  -  ", Len(costcode.List1.List(k1)), vbTextCompare)
rt1!rcc_resource = jjj(0)
rt1!t_date = costcode.DTP_tdate.Value
rt1!u_date = Now
rt1!t_user = main.Label2.Caption
rt1.Update
End If
Next
MsgBox " CostCode Modified Succesfully"

Unload costcode
Call flex_data
Call flex_title
Exit Sub
assad1:
       
       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then

Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from resourcecostcode where rcc_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "'", Cn, 3, 2
If Not dlk.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If
Dim dlk1 As New ADODB.Recordset
If dlk1.State Then dlk1.Close
dlk1.Open "select * from cost where bd_costcode='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk1.EOF Then
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
Cn.Execute "delete from costcode  where cc_id=" & id2
Cn.Execute "delete from resourcecostcode  where rcc_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload costcode
Call flex_data
Call flex_title
Else
Unload costcode
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload costcode
End If




End Sub

Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from costcode order by cc_code", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = (fldata(1))
        .TextMatrix(.Rows - 1, 2) = fldata(2)
        .TextMatrix(.Rows - 1, 3) = fldata("notes")
        .TextMatrix(.Rows - 1, 4) = fldata("t_date")
        fldata.MoveNext
    Wend
End With

End Sub


