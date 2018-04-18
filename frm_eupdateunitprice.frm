VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_eupdateunitprice 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Update UnitRate/ EIC"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.ComboBox cbo_job 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   4215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   4440
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8895
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9960
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update method"
         Height          =   1095
         Left            =   8640
         TabIndex        =   2
         Top             =   0
         Width           =   3015
         Begin VB.ComboBox cbo_curr 
            Height          =   315
            Left            =   1800
            TabIndex        =   19
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specific % Line"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Individual Line"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Line Items"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Currency"
            Height          =   255
            Left            =   1800
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   " "
         Height          =   255
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobcharge"
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
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   9960
         TabIndex        =   12
         Top             =   1155
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Existing Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   8880
         TabIndex        =   11
         Top             =   1155
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   8640
         X2              =   8640
         Y1              =   120
         Y2              =   1680
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
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
         Left            =   6885
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   17
         Top             =   0
         Width           =   4215
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
            Picture         =   "frm_eupdateunitprice.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_eupdateunitprice.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   7335
      Left            =   0
      TabIndex        =   18
      Top             =   2040
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12938
      _Version        =   393216
      Rows            =   3
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   16711680
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
Attribute VB_Name = "frm_eupdateunitprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_job_Click()
List1.Clear
nm1 = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
nmd = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
Dim rs2 As New ADODB.Recordset
If rs2.State Then rs2.Close
rs2.Open "select DISTINCT(bd_resccode) from cost where bd_projectkey='" & nm1(0) & "' and bd_jobcharge='" & nmd(0) & "'  and bd_costtype='E'  order by bd_resccode", Cn, 3, 2
While Not rs2.EOF
Dim ki As New ADODB.Recordset
If ki.State Then ki.Close
ki.Open "select DISTINCT(resc_desc) from resourcemaster where resc_code='" & rs2(0) & "' ", Cn, 3, 2
If Not ki.EOF Then
List1.AddItem rs2(0) & "  -  " & ki(0)
Else
List1.AddItem rs2(0)
End If
rs2.MoveNext
Wend
rs2.Close
'
'Form_Load
End Sub
Private Sub cbo_job_KeyPress(KeyAscii As Integer)
On Error Resume Next
 
End Sub
Private Sub cbo_proj_Click()
cbo_job.Clear
List1.Clear
nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(c.bd_jobcharge),j.job_desc from cost c , jobcharge j where c.bd_jobcharge=j.job_code and  c.bd_projectkey='" & nm(0) & "'   order by bd_jobcharge", Cn, 3, 2
While Not rs1.EOF
cbo_job.AddItem rs1(0) & "  -  " & rs1(1)
rs1.MoveNext
Wend
rs1.Close
 

Form_Load
End Sub

Private Sub cbo_proj_KeyPress(KeyAscii As Integer)
On Error Resume Next
 
End Sub

Private Sub Command1_Click()
On Error Resume Next
If cbo_curr.Text = "" Then
MsgBox "Select Currency"
Exit Sub
End If
If Option1.Value = True Then
Dim i As Integer
i = 0
For i = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0
   If CDbl(flex_grid.TextMatrix(i, 8)) = Text2.Text Then
   flex_grid.TextMatrix(i, 8) = Format(Text1.Text, "###,##0.00")
   flex_grid.TextMatrix(i, 5) = cbo_curr.Text
    Dim cr3 As New ADODB.Recordset
If cr3.State Then cr3.Close
cr3.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr3.EOF Then
    flex_grid.TextMatrix(i, 6) = cr3!cur_xchgrate
    End If
cr3.Close
   End If
            
Next
ElseIf Option3.Value = True Then
Dim j As Integer
j = 0
For j = 1 To flex_grid.Rows - 1
 
   
   flex_grid.TextMatrix(j, 8) = Format(Text1.Text, "###,##0.00")
    flex_grid.TextMatrix(j, 5) = cbo_curr.Text
      Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr2.EOF Then
    flex_grid.TextMatrix(j, 6) = cr2!cur_xchgrate
    End If
cr2.Close
Next
 
 
 
End If
End Sub

Private Sub flex_grid_Click()

On Error Resume Next
If Text1.Text = "" Then
MsgBox "Enter Unit Rate"
Exit Sub
End If
If cbo_curr.Text = "" Then
MsgBox "Select Currency"
Exit Sub
End If
'back color

Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = False



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
flex_grid.TextMatrix(current, 8) = Text1.Text
 flex_grid.TextMatrix(current, 5) = cbo_curr.Text
 Dim cr1 As New ADODB.Recordset
If cr1.State Then cr1.Close
cr1.Open "select * from currencymaster where cur_currency='" & cbo_curr.Text & "' ", Cn, 3, 2
    If Not cr1.EOF Then
    flex_grid.TextMatrix(current, 6) = cr1!cur_xchgrate
    End If
cr1.Close
Next
flex_grid.Col = 1
'Set flex_nob.CellPicture = ImageList1.ListImages(11).Picture

'---------------END------------------

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "UPDATE UNITRATE - EIC TRANSACTIONS"
Text1.Enabled = False
Text2.Enabled = False
Me.Top = 5
Me.Left = 5


Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(7).Enabled = False

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_proj.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close



flex_grid.Clear
cbo_curr.Clear
 Dim cr2 As New ADODB.Recordset
If cr2.State Then cr2.Close
cr2.Open "select * from currencymaster order by cur_currency", Cn, 3, 2
While Not cr2.EOF
cbo_curr.AddItem cr2!cur_currency
cr2.MoveNext
Wend
cr2.Close
Call flex_title
'            If flex_grid.TextMatrix(0, 5) = "% WC" Then
'            flex_grid.CellBackColor = vbGreen
'            End If
Call flex_data
  Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()

On Error Resume Next

    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Type"
        .ColWidth(1) = 500
        
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Spread"
        .ColWidth(2) = 1500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "JobCharge"
        .ColWidth(3) = 3500
        
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "CostCode"
        .ColWidth(4) = 2400
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Currency"
        .ColWidth(5) = 700
        .CellBackColor = vbGreen
        
        .TextMatrix(0, 6) = "Xchg"
        .ColWidth(6) = 600
        
'        .TextMatrix(0, 7) = "Days"
'        .ColWidth(7) = 600
         
        .TextMatrix(0, 7) = "Qty"
        .ColWidth(7) = 500
        
        .TextMatrix(0, 8) = "Unit Price"
        .ColWidth(8) = 1200
        .TextMatrix(0, 9) = "ACWP"
        '.ColWidth(9) = 1200
        .ColWidth(9) = 0
        .ColWidth(10) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub

Private Sub List1_Click()
 Call flex_title
Call flex_data
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Text1.Enabled = True
Text2.Enabled = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Text2.Enabled = False
Text1.Enabled = True
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text1.Enabled = True
Text2.Enabled = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "Modify" Then

Toolbar1.Buttons(3).Enabled = False

For i = 1 To flex_grid.Rows - 1
Dim id1 As Double
id1 = 0
            If flex_grid.TextMatrix(i, 0) = "" Then Exit Sub
            id1 = flex_grid.TextMatrix(i, 0)
                                Dim md As New ADODB.Recordset
                                If md.State Then md.Close
                                md.Open "select * from cost where bd_id=" & id1, Cn, 3, 2
                                If Not md.EOF Then
                                    md!bd_curr = flex_grid.TextMatrix(i, 5)
                                    md!bd_unitrate = flex_grid.TextMatrix(i, 8)
                                    md!bd_xchg = flex_grid.TextMatrix(i, 6)
                                md.Update
                                md.Close
                                End If
             
            
Next
MsgBox "UnitRate Updated Successfully"
Call flex_data
Call flex_title

'to delete


ElseIf Button.Caption = "Close" Then
Unload Me
End If

End Sub

Public Sub flex_data()
'On Error Resume Next
'Call flex_title
nmt = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
nmd = Split(cbo_job.Text, "  -  ", Len(cbo_job.Text), vbTextCompare)
With flex_grid
        .Rows = 1
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
nmm = Split(List1.List(i), "  -  ", Len(List1.List(i)), vbTextCompare)

Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost  where bd_resccode='" & nmm(0) & "'  and bd_jobcharge='" & nmd(0) & "' and bd_projectkey='" & nmt(0) & "' and bd_costtype='E'   order by bd_spread ,bd_jobcharge,bd_costcode", Cn, 3, 2


    
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata!bd_id
        .TextMatrix(.Rows - 1, 1) = fldata!bd_type
         
        Dim spr As New ADODB.Recordset
        If spr.State Then spr.Close
        spr.Open "select DISTINCT(spread_desc) from spreadmaster where spread_code='" & fldata!bd_spread & "' ", Cn, 3, 2
        If Not spr.EOF Then
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread & "  -  " & spr(0)
        Else
        .TextMatrix(.Rows - 1, 2) = fldata!bd_spread
        End If
        spr.Close
        Dim jc As New ADODB.Recordset
        If jc.State Then jc.Close
        jc.Open "select DISTINCT(job_desc) from jobcharge where job_code='" & fldata!bd_jobcharge & "' ", Cn, 3, 2
        If Not jc.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge & "  -  " & jc(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata!bd_jobcharge
        End If
        Dim cs As New ADODB.Recordset
        If cs.State Then cs.Close
        cs.Open "select DISTINCT(cc_desc) from costcode where cc_code='" & fldata!bd_costcode & "' ", Cn, 3, 2
        If Not cs.EOF Then
        .TextMatrix(.Rows - 1, 4) = fldata!bd_costcode & "  -  " & cs(0)
        Else
        .TextMatrix(.Rows - 1, 4) = fldata!bd_costcode
        End If
        cs.Close
        
        .TextMatrix(.Rows - 1, 5) = fldata!bd_curr
        .TextMatrix(.Rows - 1, 6) = Format(fldata!bd_xchg, "###,###,##0.00")
        '.TextMatrix(.Rows - 1, 7) = Format(fldata!bd_days, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 7) = Format(fldata!bd_qty, "###,###,##0.00")
        .TextMatrix(.Rows - 1, 8) = Format(fldata!bd_unitrate, "###,###,##0.00")
        
        '.TextMatrix(.Rows - 1, 9) = Format(fldata!bd_extdamt, "###,###,##0.00")
      
        fldata.MoveNext
    Wend



End If
Next i
End With
 
End Sub





