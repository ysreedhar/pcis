VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_jobcharge 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jobcharge"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10905
   ScaleWidth      =   11265
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
         Width           =   10815
         Begin VB.ComboBox cbo_proj 
            Height          =   315
            Left            =   3120
            TabIndex        =   5
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Key - Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   600
            TabIndex        =   6
            Top             =   240
            Width           =   2340
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10935
      _ExtentX        =   19288
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
      Width           =   11265
      _ExtentX        =   19870
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
            Picture         =   "frm_jobcharge.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_jobcharge.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_jobcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public yin As Double

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_proj_Change()
Call flex_data
End Sub

Private Sub cbo_proj_Click()
Call flex_data
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


Unload jobcharge
jobcharge.Show
jobcharge.Top = 3200
jobcharge.Left = 0
jobcharge.Height = 5010
jobcharge.Width = 6285
jobcharge.cbo_projkey.Enabled = False
jobcharge.cbo_jobno.Enabled = False
jobcharge.cbo_subjobno.Enabled = False
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)
 
jobcharge.cbo_projkey = flex_grid.TextMatrix(flex_grid.Row, 2)
jobcharge.cbo_projstatus = flex_grid.TextMatrix(flex_grid.Row, 3)
jobcharge.txt_jobcharge = flex_grid.TextMatrix(flex_grid.Row, 1)
jobcharge.txt_jobdesc = flex_grid.TextMatrix(flex_grid.Row, 4)
jobcharge.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 5)
jobcharge.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
nm = Split(flex_grid.TextMatrix(flex_grid.Row, 1), "-", Len(flex_grid.TextMatrix(flex_grid.Row, 1)), vbTextCompare)
Dim jn As New ADODB.Recordset
If jn.State Then jn.Close
jn.Open "select DISTINCT(jobno_desc) from jobno where jobno_code='" & nm(0) & "'", Cn, 3, 2
If Not jn.EOF Then
jobcharge.cbo_jobno.Text = nm(0) & "  -  " & jn(0)
Else
jobcharge.cbo_jobno.Text = nm(0)
End If
Dim sj As New ADODB.Recordset
If sj.State Then sj.Close
sj.Open "select DISTINCT(subjobno_desc) from subjobno where subjobno_code='" & nm(1) & "' ", Cn, 3, 2
If Not sj.EOF Then
jobcharge.cbo_subjobno.Text = nm(1) & "  -  " & sj(0)
 Else
 jobcharge.cbo_subjobno.Text = nm(1)
 End If

 

jobcharge.txt_jobcharge.Enabled = False
vprev = flex_grid.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "JOBCHARGE NO."
Call flex_title
Call flex_data
Me.Top = 5
Me.Left = 5
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_desc from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2 & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_proj.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close
 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()

On Error Resume Next
    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Job Charge"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Project Key"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Status"
        .ColWidth(3) = 800
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "JobCharge Description"
        .ColWidth(4) = 3300
        .ColAlignment(4) = 0
        .ColWidth(5) = 0
        .TextMatrix(0, 6) = "Notes"
        .ColWidth(6) = 4000
        .ColAlignment(6) = 0
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload jobcharge
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
If cbo_proj.Text = "" Then
MsgBox "select Project"
cbo_proj.SetFocus
Exit Sub
End If
jobcharge.txt_jobcharge.Enabled = True
jobcharge.cbo_projkey.Enabled = True
jobcharge.cbo_jobno.Enabled = True
jobcharge.cbo_subjobno.Enabled = True
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload jobcharge
jobcharge.Show
jobcharge.Top = 3200
jobcharge.Left = 90
jobcharge.Height = 5010
jobcharge.Width = 6285
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
If jobcharge.cbo_projkey.Text = "" Then
MsgBox "Select Project Key"
jobcharge.cbo_projkey.SetFocus
Exit Sub
End If
If jobcharge.cbo_jobno.Text = "" Then
MsgBox "Select Job No"
jobcharge.cbo_jobno.SetFocus
Exit Sub
End If
If jobcharge.cbo_subjobno.Text = "" Then
MsgBox "Select SUB-JobNO"
jobcharge.cbo_subjobno.SetFocus
Exit Sub
End If
If jobcharge.txt_jobcharge.Text = "" Then
MsgBox "Enter Job Charge"
jobcharge.txt_jobcharge.SetFocus
Exit Sub
End If
hj = Split(jobcharge.cbo_projkey, "  -  ", Len(jobcharge.cbo_projkey), vbTextCompare)
hjj = Split(jobcharge.cbo_jobno.Text, "  -  ", Len(jobcharge.cbo_jobno.Text), vbTextCompare)
hjjj = Split(jobcharge.cbo_subjobno.Text, "  -  ", Len(jobcharge.cbo_subjobno.Text), vbTextCompare)
  
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from jobcharge", Cn, 3, 2
sv.AddNew
sv!job_proj_key = hj(0)
sv!job_proj_status = jobcharge.cbo_projstatus
sv!jobno = hjj(0)
sv!subjobno = hjjj(0)
sv!job_code = jobcharge.txt_jobcharge.Text
sv!job_desc = jobcharge.txt_jobdesc
sv!t_date = jobcharge.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv!notes = jobcharge.txt_notes.Text
sv.Update
sv.Close
MsgBox "New Jobcharge Added Succesfully"
Unload jobcharge
Call flex_data
Call flex_title
Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1
If jobcharge.cbo_projkey.Text = "" Then
MsgBox "Select Project Key"
jobcharge.cbo_projkey.SetFocus
Exit Sub
End If
If jobcharge.cbo_jobno.Text = "" Then
MsgBox "Select Job No"
jobcharge.cbo_jobno.SetFocus
Exit Sub
End If
If jobcharge.cbo_subjobno.Text = "" Then
MsgBox "Select SUB-JobNO"
jobcharge.cbo_subjobno.SetFocus
Exit Sub
End If
If jobcharge.txt_jobcharge.Text = "" Then
MsgBox "Enter Job Charge"
jobcharge.txt_jobcharge.SetFocus
Exit Sub
End If
Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
hj = Split(jobcharge.cbo_projkey, "  -  ", Len(jobcharge.cbo_projkey), vbTextCompare)
hjj = Split(jobcharge.cbo_jobno.Text, "  -  ", Len(jobcharge.cbo_jobno.Text), vbTextCompare)
hjjj = Split(jobcharge.cbo_subjobno.Text, "  -  ", Len(jobcharge.cbo_subjobno.Text), vbTextCompare)
Dim iit As Integer
iit = 0
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from jobcharge where job_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!job_proj_key = hj(0)
md!job_proj_status = jobcharge.cbo_projstatus
'md!jobno = hjj(0)
'md!subjobno = hjjj(0)
'md!job_code = jobcharge.txt_jobcharge.Text
md!job_desc = jobcharge.txt_jobdesc
md!t_date = jobcharge.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md!notes = jobcharge.txt_notes.Text
iit = md!job_id
md.Update
md.Close
If jobcharge.cbo_projstatus.Text = "InActive" Then
Call delscope
Call delscopebdgt
Call delscopeprgs
ElseIf jobcharge.cbo_projstatus.Text = "Active" Then
Call addscope
Call addscopebdgt
'Call addscopeprgs
 
End If
MsgBox "Selected Jobcharge Modified"
End If
 
Unload jobcharge
Call flex_data
Call flex_title

Exit Sub
assad1:
       
       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then

Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from cost where bd_jobcharge='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If

Dim dlk1 As New ADODB.Recordset
If dlk1.State Then dlk1.Close
dlk1.Open "select * from progressdurationdetails where prgs_job_key='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk1.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If

Dim dlk2 As New ADODB.Recordset
If dlk2.State Then dlk2.Close
dlk2.Open "select * from budgeteddurationdetails where bdgt_job_key='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk2.EOF Then
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
Cn.Execute "delete from jobcharge where job_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload jobcharge
Call flex_data
Call flex_title
Else
Unload jobcharge
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload jobcharge
End If




End Sub

Public Sub flex_data()
On Error Resume Next
hj = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
If hj(0) = "" Then Exit Sub
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from jobcharge where job_proj_key='" & hj(0) & "' order by job_code", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata("job_code")
        Dim fd As New ADODB.Recordset
        If fd.State Then fd.Close
        fd.Open "select DISTINCT(proj_desc) from projectmaster where proj_key='" & fldata("job_proj_key") & "'", Cn, 3, 2
        If Not fd.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata("job_proj_key") & "  -  " & fd(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata("job_proj_key")
        End If
        .TextMatrix(.Rows - 1, 3) = fldata("job_proj_status")
        .TextMatrix(.Rows - 1, 4) = fldata("job_desc")
        .TextMatrix(.Rows - 1, 5) = fldata("t_date")
        .TextMatrix(.Rows - 1, 6) = fldata("notes")
        fldata.MoveNext
    Wend
End With
End Sub

Public Sub delscope()
dels = Split(jobcharge.cbo_projkey, "  -  ", Len(jobcharge.cbo_projkey), vbTextCompare)
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)
Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from cost where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "' order by bd_id ", Cn, 3, 2
 yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub
 
Cn.Execute "delete from costdelscope where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from cost where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "' order by bd_id ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from costdelscope", Cn, 3, 2
          dcs.AddNew
   
    dcs!bd_year = cs!bd_year
    dcs!bd_resccode = cs!bd_resccode
    dcs!bd_rescname = cs!bd_rescname
    dcs!bd_brate = cs!bd_brate
    dcs!bd_crate = cs!bd_crate
    dcs!bd_vendor = cs!bd_vendor
    dcs!bd_projectkey = cs!bd_projectkey
    dcs!bd_projectdesc = cs!bd_projectdesc
    dcs!bd_costtype = cs!bd_costtype
    dcs!bd_respcode = cs!bd_respcode
    dcs!bd_respname = cs!bd_respname
    dcs!bd_cuttdate = cs!bd_cuttdate
    dcs!bd_spread = cs!bd_spread
    dcs!bd_tranx = cs!bd_tranx
    dcs!bd_jobcharge = cs!bd_jobcharge
    dcs!bd_costcode = cs!bd_costcode
    dcs!bd_qty = cs!bd_qty
    dcs!bd_days = cs!bd_days
    dcs!bd_tqty = cs!bd_tqty
    dcs!bd_uom = cs!bd_uom
    dcs!bd_curr = cs!bd_curr
    dcs!bd_unitrate = cs!bd_unitrate
    dcs!bd_xchg = cs!bd_xchg
    dcs!bd_downtime = cs!bd_downtime
    dcs!bd_escl = cs!bd_escl
    dcs!bd_extdamt = cs!bd_extdamt
    dcs!bd_wrkcomp = cs!bd_wrkcomp
    dcs!bd_bcwpamt = cs!bd_bcwpamt
    dcs!bd_e_days = cs!bd_e_days
    dcs!bd_e_tqty = cs!bd_e_tqty
    dcs!bd_e_extdamt = cs!bd_e_extdamt
    dcs!bd_chk = cs!bd_chk
    dcs!bd_sdate = cs!bd_sdate
    dcs!bd_edate = cs!bd_edate
    dcs!bd_notes = cs!bd_notes
    dcs!t_user = cs!t_user
    dcs!t_date = cs!t_date
    dcs!u_date = cs!u_date
    dcs!bd_inv = cs!bd_inv
    dcs!bd_invdate = cs!bd_invdate
    dcs!bd_type = cs!bd_type
    dcs!bd_obs = cs!bd_obs
    dcs!estid = cs!estid
    dcs!bd_chk1 = cs!bd_chk1
    dcs.Update
    cs.MoveNext
    Wend

Cn.Execute "delete from cost where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "'"
End Sub

Public Sub addscope()
dels = Split(jobcharge.cbo_projkey, "  -  ", Len(jobcharge.cbo_projkey), vbTextCompare)
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)
Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from costdelscope where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "' order by bd_id ", Cn, 3, 2
  yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub

Cn.Execute "delete from cost where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from costdelscope where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "' order by bd_id ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from cost", Cn, 3, 2
          dcs.AddNew
   
    dcs!bd_year = cs!bd_year
    dcs!bd_resccode = cs!bd_resccode
    dcs!bd_rescname = cs!bd_rescname
    dcs!bd_brate = cs!bd_brate
    dcs!bd_crate = cs!bd_crate
    dcs!bd_vendor = cs!bd_vendor
    dcs!bd_projectkey = cs!bd_projectkey
    dcs!bd_projectdesc = cs!bd_projectdesc
    dcs!bd_costtype = cs!bd_costtype
    dcs!bd_respcode = cs!bd_respcode
    dcs!bd_respname = cs!bd_respname
    dcs!bd_cuttdate = cs!bd_cuttdate
    dcs!bd_spread = cs!bd_spread
    dcs!bd_tranx = cs!bd_tranx
    dcs!bd_jobcharge = cs!bd_jobcharge
    dcs!bd_costcode = cs!bd_costcode
    dcs!bd_qty = cs!bd_qty
    dcs!bd_days = cs!bd_days
    dcs!bd_tqty = cs!bd_tqty
    dcs!bd_uom = cs!bd_uom
    dcs!bd_curr = cs!bd_curr
    dcs!bd_unitrate = cs!bd_unitrate
    dcs!bd_xchg = cs!bd_xchg
    dcs!bd_downtime = cs!bd_downtime
    dcs!bd_escl = cs!bd_escl
    dcs!bd_extdamt = cs!bd_extdamt
    dcs!bd_wrkcomp = cs!bd_wrkcomp
    dcs!bd_bcwpamt = cs!bd_bcwpamt
    dcs!bd_e_days = cs!bd_e_days
    dcs!bd_e_tqty = cs!bd_e_tqty
    dcs!bd_e_extdamt = cs!bd_e_extdamt
    dcs!bd_chk = cs!bd_chk
    dcs!bd_sdate = cs!bd_sdate
    dcs!bd_edate = cs!bd_edate
    dcs!bd_notes = cs!bd_notes
    dcs!t_user = cs!t_user
    dcs!t_date = cs!t_date
    dcs!u_date = cs!u_date
    dcs!bd_inv = cs!bd_inv
    dcs!bd_invdate = cs!bd_invdate
    dcs!bd_type = cs!bd_type
    dcs!bd_obs = cs!bd_obs
    dcs!estid = cs!estid
    dcs!bd_chk1 = cs!bd_chk1
    dcs.Update
    cs.MoveNext
Wend


Cn.Execute "delete from costdelscope where bd_projectkey='" & dels(0) & "' and bd_jobcharge='" & dels1(0) & "'"
End Sub

Public Sub delscopebdgt()
 
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)
Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from budgeteddurationdetails where   bdgt_job_key='" & dels1(0) & "'  ", Cn, 3, 2
 yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub

Cn.Execute "delete from budgeteddurationdetailsdelscope where  bdgt_job_key='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from budgeteddurationdetails where   bdgt_job_key='" & dels1(0) & "'  ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from budgeteddurationdetailsdelscope", Cn, 3, 2
          dcs.AddNew
   
dcs!bdgt_spread_code = cs!bdgt_spread_code
dcs!bdgt_job_key = cs!bdgt_job_key
dcs!bdgt_days = cs!bdgt_days
dcs!bdgt_per_workcomplete = cs!bdgt_per_workcomplete
dcs!bdgt_remarks = cs!bdgt_remarks
dcs!t_date = cs!t_date
dcs!u_date = cs!u_date
dcs!t_user = cs!t_user
    dcs.Update
    cs.MoveNext
    Wend

Cn.Execute "delete from budgeteddurationdetails where  bdgt_job_key='" & dels1(0) & "'"
End Sub

Public Sub addscopebdgt()
 
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)

Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from budgeteddurationdetailsdelscope where bdgt_job_key='" & dels1(0) & "' ", Cn, 3, 2
 yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub

Cn.Execute "delete from budgeteddurationdetails where bdgt_job_key='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from budgeteddurationdetailsdelscope where bdgt_job_key='" & dels1(0) & "' ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from budgeteddurationdetails", Cn, 3, 2
          dcs.AddNew
   
dcs!bdgt_spread_code = cs!bdgt_spread_code
dcs!bdgt_job_key = cs!bdgt_job_key
dcs!bdgt_days = cs!bdgt_days
dcs!bdgt_per_workcomplete = cs!bdgt_per_workcomplete
dcs!bdgt_remarks = cs!bdgt_remarks
dcs!t_date = cs!t_date
dcs!u_date = cs!u_date
dcs!t_user = cs!t_user
    dcs.Update
    cs.MoveNext
Wend


Cn.Execute "delete from budgeteddurationdetailsdelscope where  bdgt_job_key='" & dels1(0) & "'"
End Sub
Public Sub delscopeprgs()
 
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)
Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from progressdurationdetails where   prgs_job_key='" & dels1(0) & "'  ", Cn, 3, 2
 yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub

Cn.Execute "delete from progressdurationdetailsdelscope where  prgs_job_key='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from progressdurationdetails where   prgs_job_key='" & dels1(0) & "'  ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from progressdurationdetailsdelscope", Cn, 3, 2
          dcs.AddNew
   
dcs!prgs_spread_code = cs!prgs_spread_code
dcs!prgs_job_key = cs!prgs_job_key
dcs!prgs_startdate = cs!prgs_startdate
dcs!prgs_enddate = cs!prgs_enddate
dcs!prgs_remarks = cs!prgs_remarks
dcs!prgs_days = cs!prgs_days
dcs!t_date = cs!t_date
dcs!u_date = cs!u_date
dcs!t_user = cs!t_user
dcs!prgs_type = cs!prgs_type
    dcs.Update
    cs.MoveNext
    Wend

Cn.Execute "delete from progressdurationdetails where  prgs_job_key='" & dels1(0) & "'"
End Sub

Public Sub addscopeprgs()
 
dels1 = Split(jobcharge.txt_jobcharge.Text, "  -  ", Len(jobcharge.txt_jobcharge.Text), vbTextCompare)
Dim csq As New ADODB.Recordset
If csq.State Then csq.Close
csq.Open "select * from progressdurationdetailsdelscope where prgs_job_key='" & dels1(0) & "' ", Cn, 3, 2
 yin = 0
 yin = csq.RecordCount
 If yin = 0 Then Exit Sub


Cn.Execute "delete from progressdurationdetails where prgs_job_key='" & dels1(0) & "'"
Dim cs As New ADODB.Recordset
If cs.State Then cs.Close
cs.Open "select * from progressdurationdetailsdelscope where prgs_job_key='" & dels1(0) & "' ", Cn, 3, 2
While Not cs.EOF
          Dim dcs As New ADODB.Recordset
          If dcs.State Then dcs.Close
          dcs.Open "select * from progressdurationdetails", Cn, 3, 2
          dcs.AddNew
dcs!prgs_spread_code = cs!prgs_spread_code
dcs!prgs_job_key = cs!prgs_job_key
dcs!prgs_startdate = cs!prgs_startdate
dcs!prgs_enddate = cs!prgs_enddate
dcs!prgs_remarks = cs!prgs_remarks
dcs!prgs_days = cs!prgs_days
dcs!t_date = cs!t_date
dcs!u_date = cs!u_date
dcs!t_user = cs!t_user
dcs!prgs_type = cs!prgs_type
    dcs.Update
    cs.MoveNext
Wend


Cn.Execute "delete from progressdurationdetailsdelscope where  prgs_job_key='" & dels1(0) & "'"
End Sub

