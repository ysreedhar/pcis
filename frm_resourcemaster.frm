VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_resourcemaster 
   BackColor       =   &H00FFFFFF&
   Caption         =   "RESOURCE CODE"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14505
      TabIndex        =   3
      Top             =   7080
      Width           =   14535
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   14445
         _ExtentX        =   25479
         _ExtentY        =   582
         ButtonWidth     =   1561
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList5"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Key             =   "s"
               Object.ToolTipText     =   "Invoice  details"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modify"
               Key             =   "m"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "d"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.PictureBox Picture3 
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
            TabIndex        =   5
            Top             =   0
            Width           =   2295
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Resource Details"
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
               Height          =   240
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   1815
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
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
            Key             =   "n"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "s"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "m"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "d"
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
            Picture         =   "frm_resourcemaster.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_resourcemaster.frx":13162
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid1 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   3
      Cols            =   8
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
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   6735
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   3
      Cols            =   8
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
End
Attribute VB_Name = "frm_resourcemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub cmd_close_Click()

End Sub

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub



Private Sub flex_grid_Click()
flex_grid1.Visible = True
Toolbar2.Visible = True
Picture2.Visible = True
 
 Toolbar2.Buttons(1).Enabled = False
 
Call flex_datanew
Call flex_titlenew
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

vprev = flex_grid.Row
End Sub



Private Sub flex_grid_DblClick()
'On Error Resume Next
resourcemaster.SSTab1.Tab = 0
flex_grid1.Visible = True
 
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




'-----------END----------

Unload resourcemaster
resourcemaster.Show
resourcemaster.Top = 3000
resourcemaster.Left = 0
resourcemaster.Height = 4005
resourcemaster.Width = 10590
resourcemaster.txt_rescourcecode.Enabled = False
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

 
 
resourcemaster.txt_rescourcecode = flex_grid.TextMatrix(flex_grid.Row, 1)
resourcemaster.txt_resourcedesc = flex_grid.TextMatrix(flex_grid.Row, 2)
resourcemaster.txt_standardrate = flex_grid.TextMatrix(flex_grid.Row, 3)
resourcemaster.cbo_vendor = flex_grid.TextMatrix(flex_grid.Row, 4)
resourcemaster.cbo_uom = flex_grid.TextMatrix(flex_grid.Row, 6)
resourcemaster.cbo_resp.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
resourcemaster.DTP_tdate.Value = flex_grid.TextMatrix(flex_grid.Row, 7)
 


Call flex_datanew
Call flex_titlenew
vprev = flex_grid.Row


Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True

Toolbar2.Buttons(1).Enabled = False
Toolbar2.Buttons(3).Enabled = False
Toolbar2.Buttons(5).Enabled = False
'resourcemaster.Frame1.Enabled = False
'resourcemaster.SSTab1.Tab = 1
End Sub

Private Sub flex_grid1_Click()

On Error Resume Next
 
'flex_grid_DblClick
'back color

Static vprev As Integer

current = flex_grid1.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid1.Row = vprev
    flex_grid1.Col = 1
    Set flex_grid1.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid1.Cols - 1
    flex_grid1.Col = i
    flex_grid1.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid1.Row = current
For i = 1 To flex_grid1.Cols - 1
flex_grid1.Col = i
flex_grid1.CellBackColor = vbYellow
Next
flex_grid1.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture




'-------------END----------

Unload resourcemaster
Dim id As Double
id = 0
If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id = flex_grid1.TextMatrix(flex_grid1.Row, 0)
Dim id5 As String
id5 = 0
If flex_grid1.TextMatrix(flex_grid1.Row, 5) = "" Then Exit Sub
id5 = flex_grid1.TextMatrix(flex_grid1.Row, 5)

Dim sh As New ADODB.Recordset
If sh.State Then sh.Close
sh.Open "select * from resourcedetails rd, resourcemaster rm where rd.resc_id=rm.resc_id and  rd.dresc_ratetype='" & id5 & "' and rd.dresc_id=" & id, Cn, 3, 2
If Not sh.EOF Then
resourcemaster.cbo_resccode.Text = sh!dresc_code
resourcemaster.DTP_resc.Text = sh!dresc_year
resourcemaster.cbo_curcy.Text = sh!dresc_curcy
resourcemaster.txt_rate.Text = sh!dresc_rate
resourcemaster.txt_ratetype.Text = sh!dresc_ratetype
resourcemaster.txt_notes = sh!dresc_notes
resourcemaster.DTP_tdate.Value = sh!t_date

resourcemaster.cbo_projkey = sh!dresc_proj
resourcemaster.txt_rescourcecode = sh!resc_code
resourcemaster.txt_resourcedesc = sh!resc_desc
resourcemaster.txt_standardrate = sh!resc_type
resourcemaster.cbo_vendor = sh!resc_vendorcode
resourcemaster.cbo_uom = sh!resc_uom
  
resourcemaster.cbo_resp.Text = sh!resc_respcode
 
resourcemaster.DTP_tdate.Value = sh!t_date
End If
resourcemaster.Show
resourcemaster.Top = 3000
resourcemaster.Left = 0
resourcemaster.Height = 4005
resourcemaster.Width = 10590
sh.Close
resourcemaster.SSTab1.Tab = 1
vprev = flex_grid1.Row

Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False

Toolbar2.Buttons(1).Enabled = False
Toolbar2.Buttons(3).Enabled = True
Toolbar2.Buttons(5).Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "RESOURCE CODE"
a = 0
Call flex_title
Call flex_data
Me.Top = 5
Me.Left = 5
 

Toolbar2.Buttons(1).Enabled = False
Toolbar2.Buttons(3).Enabled = False
Toolbar2.Buttons(5).Enabled = False

Toolbar1.Buttons(1).Enabled = True
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False

flex_grid1.Visible = False
Toolbar2.Visible = False
Picture2.Visible = False

 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()
On Error Resume Next

    With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Resource Code"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Resource Description"
        .ColWidth(2) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Resource Type"
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Vendor Code"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Resp Code"
        .ColWidth(5) = 1500
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "U.O.M"
        .ColWidth(6) = 700
        .ColAlignment(6) = 0
        .ColWidth(7) = 0
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload resourcemaster
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
If Button.Caption = "New" Then
'resourcemaster.Frame1.Enabled = True

Unload resourcemaster
resourcemaster.Show
resourcemaster.Top = 3000
resourcemaster.Left = 0
resourcemaster.Height = 4005
resourcemaster.Width = 10590
resourcemaster.txt_rescourcecode.Enabled = True
Toolbar1.Buttons(1).Enabled = True
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
If resourcemaster.txt_rescourcecode.Text = "" Then
MsgBox "Enter Resource Code"
resourcemaster.txt_rescourcecode.SetFocus
Exit Sub
End If
If resourcemaster.txt_resourcedesc.Text = "" Then
MsgBox "Enter Resource Desc"
resourcemaster.txt_resourcedesc.SetFocus
Exit Sub
End If
If resourcemaster.txt_standardrate.Text = "" Then
MsgBox "Select Resource Type"
resourcemaster.txt_standardrate.SetFocus
Exit Sub
End If
If resourcemaster.cbo_vendor.Text = "" Then
MsgBox "Select Vendor"
resourcemaster.cbo_vendor.SetFocus
Exit Sub
End If

If resourcemaster.cbo_resp.Text = "" Then
MsgBox "Select Resc Responsible"
resourcemaster.cbo_resp.SetFocus
Exit Sub
End If

If resourcemaster.cbo_uom.Text = "" Then
MsgBox "Select UOM"
resourcemaster.cbo_uom.SetFocus
Exit Sub
End If
rv1 = Split(resourcemaster.cbo_projkey.Text, "  -  ", Len(resourcemaster.cbo_projkey.Text), vbTextCompare)
rv = Split(resourcemaster.cbo_vendor.Text, "  -  ", Len(resourcemaster.cbo_vendor.Text), vbTextCompare)
rvv = Split(resourcemaster.cbo_uom.Text, "  -  ", Len(resourcemaster.cbo_uom.Text), vbTextCompare)
rvvv = Split(resourcemaster.cbo_resp.Text, "  -  ", Len(resourcemaster.cbo_resp.Text), vbTextCompare)
rvv1 = Split(resourcemaster.txt_standardrate.Text, "  -  ", Len(resourcemaster.txt_standardrate.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from resourcemaster", Cn, 3, 2
sv.AddNew
 
sv!resc_code = resourcemaster.txt_rescourcecode.Text
sv!resc_desc = resourcemaster.txt_resourcedesc.Text
sv!resc_type = rvv1(0)
sv!resc_vendorcode = rv(0)
sv!resc_uom = rvv(0)
sv!resc_respcode = rvvv(0)
 
sv!t_date = resourcemaster.DTP_tdate.Value
sv!u_date = Now
sv!t_user = main.Label2.Caption
sv.Update
sv.Close
MsgBox "New Resource Added Succesfully"

Unload resourcemaster
Call flex_data
Call flex_title
Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1

If resourcemaster.txt_rescourcecode.Text = "" Then
MsgBox "Enter Resource Code"
resourcemaster.txt_rescourcecode.SetFocus
Exit Sub
End If
If resourcemaster.txt_resourcedesc.Text = "" Then
MsgBox "Enter Resource Desc"
resourcemaster.txt_resourcedesc.SetFocus
Exit Sub
End If
If resourcemaster.txt_standardrate.Text = "" Then
MsgBox "Select Resource Type"
resourcemaster.txt_standardrate.SetFocus
Exit Sub
End If
If resourcemaster.cbo_vendor.Text = "" Then
MsgBox "Select Vendor"
resourcemaster.cbo_vendor.SetFocus
Exit Sub
End If

If resourcemaster.cbo_resp.Text = "" Then
MsgBox "Select Resc Responsible"
resourcemaster.cbo_resp.SetFocus
Exit Sub
End If
If resourcemaster.cbo_uom.Text = "" Then
MsgBox "Select UOM"
resourcemaster.cbo_uom.SetFocus
Exit Sub
End If
Toolbar1.Buttons(1).Enabled = False

rv1 = Split(resourcemaster.cbo_projkey.Text, "  -  ", Len(resourcemaster.cbo_projkey.Text), vbTextCompare)
rv = Split(resourcemaster.cbo_vendor.Text, "  -  ", Len(resourcemaster.cbo_vendor.Text), vbTextCompare)
rvv = Split(resourcemaster.cbo_uom.Text, "  -  ", Len(resourcemaster.cbo_uom.Text), vbTextCompare)
rvvv = Split(resourcemaster.cbo_resp.Text, "  -  ", Len(resourcemaster.cbo_resp.Text), vbTextCompare)

Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from resourcemaster where resc_id=" & id1, Cn, 3, 2
If Not md.EOF Then
 
 
md!resc_desc = resourcemaster.txt_resourcedesc
md!resc_type = resourcemaster.txt_standardrate
md!resc_vendorcode = rv(0)
md!resc_uom = rvv(0)
md!resc_respcode = rvvv(0)
md!t_date = resourcemaster.DTP_tdate.Value
md!u_date = Now
md!t_user = main.Label2.Caption
md.Update
md.Close
MsgBox "Selected Resource Modified"
End If

Unload resourcemaster
Call flex_data
Call flex_title
Exit Sub
assad1:
       
       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from cost where bd_resccode='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If


Toolbar1.Buttons(1).Enabled = False



dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from resourcemaster where resc_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload resourcemaster
Call flex_data
Call flex_title
Else
Unload resourcemaster
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload resourcemaster
End If




End Sub

Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from resourcemaster order by resc_code", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!resc_code
        .TextMatrix(.Rows - 1, 2) = fldata!resc_desc
        Dim rt As New ADODB.Recordset
        If rt.State Then rt.Close
        rt.Open "select DISTINCT(r_desc) from resourcetype where r_type='" & fldata!resc_type & "' ", Cn, 3, 2
        If Not rt.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata!resc_type & "  -  " & rt(0)
        Else
         .TextMatrix(.Rows - 1, 3) = fldata!resc_type
        End If
        rt.Close
        Dim vc As New ADODB.Recordset
              
        If vc.State Then vc.Close
        vc.Open "select DISTINCT(vendor_desc) from vendormaster where vendor_code='" & fldata!resc_vendorcode & "' ", Cn, 3, 2
        If Not vc.EOF Then
        .TextMatrix(.Rows - 1, 4) = fldata!resc_vendorcode & "  -  " & vc(0)
        Else
        .TextMatrix(.Rows - 1, 4) = fldata!resc_vendorcode
        End If
        vc.Close
        Dim rr As New ADODB.Recordset
        If rr.State Then rr.Close
        rr.Open "select DISTINCT(resp_desc) from responsiblemaster where resp_code='" & fldata!resc_respcode & "' ", Cn, 3, 2
        If Not rr.EOF Then
        .TextMatrix(.Rows - 1, 5) = fldata!resc_respcode & "  -  " & rr(0)
        Else
        .TextMatrix(.Rows - 1, 5) = fldata!resc_respcode
        End If
        rr.Close
        Dim um As New ADODB.Recordset
        If um.State Then um.Close
        um.Open "select Distinct(uom_desc) from uom where uom_uom='" & fldata!resc_uom & "' ", Cn, 3, 2
        If Not um.EOF Then
        .TextMatrix(.Rows - 1, 6) = fldata!resc_uom & "  -  " & um(0)
        Else
        .TextMatrix(.Rows - 1, 6) = fldata!resc_uom
        End If
        .TextMatrix(.Rows - 1, 7) = fldata!t_date
        fldata.MoveNext
    Wend
End With
flex_grid.Refresh
End Sub


Public Sub flex_titlenew()
On Error Resume Next
    With flex_grid1
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Resc Code"
        .ColWidth(1) = 900
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Year"
        .ColWidth(2) = 800
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Currency"
        .ColWidth(3) = 1100
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Rate"
        .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Rate Type"
        .ColWidth(5) = 1000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Project"
        .ColWidth(6) = 3300
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 4000
        .ColAlignment(7) = 0
        
    End With
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

If Button.Caption = "Save" Then
If resourcemaster.cbo_projkey.Text = "" Then
MsgBox "Select Project Key"
resourcemaster.cbo_projkey.SetFocus
Exit Sub
End If
If resourcemaster.DTP_resc.Text = "" Then
MsgBox "Select Year"
resourcemaster.DTP_resc.SetFocus
Exit Sub
End If
If resourcemaster.cbo_curcy.Text = "" Then
MsgBox "Select Currency"
resourcemaster.cbo_curcy.SetFocus
Exit Sub
End If
If resourcemaster.txt_rate.Text = "" Then
MsgBox "Enter Rate"
resourcemaster.txt_rate.SetFocus
Exit Sub
End If
If resourcemaster.txt_ratetype.Text = "" Then
MsgBox "Select Rate Type"
resourcemaster.txt_ratetype.SetFocus
Exit Sub
End If
rvt = Split(resourcemaster.cbo_projkey.Text, "  -  ", Len(resourcemaster.cbo_projkey.Text), vbTextCompare)
rvt1 = Split(resourcemaster.cbo_curcy.Text, "  -  ", Len(resourcemaster.cbo_curcy.Text), vbTextCompare)

Dim id4 As Double
id4 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id4 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from resourcedetails", Cn, 3, 2
sv.AddNew
 sv!dresc_proj = rvt(0)
 sv!dresc_code = resourcemaster.cbo_resccode.Text
 sv!dresc_year = resourcemaster.DTP_resc.Text
 sv!dresc_curcy = rvt1(0)
 sv!dresc_rate = resourcemaster.txt_rate.Text
 sv!dresc_ratetype = resourcemaster.txt_ratetype.Text
 sv!dresc_notes = resourcemaster.txt_notes.Text
 sv!resc_id = id4
 sv!t_date = resourcemaster.DTP_tdate.Value
 sv!u_date = Now
 sv!t_user = main.Label2.Caption
sv.Update
sv.Close
MsgBox "New Resource Details Added Succesfully"
'Call bdgtcost
Unload resourcemaster
Call flex_datanew
Call flex_titlenew
'to modify existing record
ElseIf Button.Caption = "Modify" Then
If resourcemaster.cbo_projkey.Text = "" Then
MsgBox "Select Project Key"
resourcemaster.cbo_projkey.SetFocus
Exit Sub
End If
If resourcemaster.DTP_resc.Text = "" Then
MsgBox "Select Year"
resourcemaster.DTP_resc.SetFocus
Exit Sub
End If
If resourcemaster.cbo_curcy.Text = "" Then
MsgBox "Select Currency"
resourcemaster.cbo_curcy.SetFocus
Exit Sub
End If
If resourcemaster.txt_rate.Text = "" Then
MsgBox "Enter Rate"
resourcemaster.txt_rate.SetFocus
Exit Sub
End If
If resourcemaster.txt_ratetype.Text = "" Then
MsgBox "Select Rate Type"
resourcemaster.txt_ratetype.SetFocus
Exit Sub
End If
Toolbar1.Buttons(1).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id1 = flex_grid1.TextMatrix(flex_grid1.Row, 0)
rvt = Split(resourcemaster.cbo_projkey.Text, "  -  ", Len(resourcemaster.cbo_projkey.Text), vbTextCompare)
rvt1 = Split(resourcemaster.cbo_curcy.Text, "  -  ", Len(resourcemaster.cbo_curcy.Text), vbTextCompare)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from resourcedetails where dresc_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!dresc_proj = rvt(0)
 md!dresc_code = resourcemaster.cbo_resccode.Text
 md!dresc_year = resourcemaster.DTP_resc.Text
 md!dresc_curcy = rvt1(0)
 md!dresc_rate = resourcemaster.txt_rate.Text
 md!dresc_ratetype = resourcemaster.txt_ratetype.Text
 md!dresc_notes = resourcemaster.txt_notes.Text
  
 md!t_date = resourcemaster.DTP_tdate.Value
 md!u_date = Now
 md!t_user = main.Label2.Caption
md.Update
md.Close
MsgBox "Selected Resource Details Modified"
End If
'Call bdgtcost
Unload resourcemaster
Call flex_datanew
Call flex_titlenew

'to delete
ElseIf Button.Caption = "Delete" Then
Dim dlk As New ADODB.Recordset
If dlk.State Then dlk.Close
dlk.Open "select * from cost where bd_resccode='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
If Not dlk.EOF Then
MsgBox "Cannot Delete This Record"
Exit Sub
End If
Toolbar1.Buttons(1).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid1.TextMatrix(flex_grid1.Row, 0) = "" Then Exit Sub
id2 = flex_grid1.TextMatrix(flex_grid1.Row, 0)
Cn.Execute "delete from resourcedetails where dresc_id=" & id2
MsgBox "Selected Record Has Been Deleted"
Unload resourcemaster
Call flex_datanew
Call flex_titlenew
Else
Unload resourcemaster
End If
ElseIf Button.Caption = "Close" Then
 
flex_grid1.Visible = False
Toolbar2.Visible = False
Picture2.Visible = False
End If

End Sub

Public Sub flex_datanew()
On Error Resume Next
Dim id3 As String
id3 = 0
If flex_grid.TextMatrix(flex_grid.Row, 1) = "" Then Exit Sub
id3 = flex_grid.TextMatrix(flex_grid.Row, 1)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from resourcedetails r,userproject u where r.dresc_proj=u.project and u.username='" & main.Label2.Caption & "' and  r.dresc_code='" & id3 & "' order by r.dresc_year,r.dresc_proj", Cn, 3, 2
With flex_grid1
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata(1)
        .TextMatrix(.Rows - 1, 2) = fldata(2)
        Dim cur As New ADODB.Recordset
        If cur.State Then cur.Close
        cur.Open "select DISTINCT(c_desc) from currency where c_name='" & fldata(3) & "' ", Cn, 3, 2
        If Not cur.EOF Then
        .TextMatrix(.Rows - 1, 3) = fldata(3) & "  -  " & cur(0)
        Else
        .TextMatrix(.Rows - 1, 3) = fldata(3)
        End If
        cur.Close
        .TextMatrix(.Rows - 1, 4) = Format(fldata(4), "###,###,##0.00")
        .TextMatrix(.Rows - 1, 5) = fldata(5)
        Dim pj As New ADODB.Recordset
        If pj.State Then pj.Close
        pj.Open "select DISTINCT(proj_desc) from projectmaster where proj_key='" & fldata("dresc_proj") & "' ", Cn, 3, 2
        If Not pj.EOF Then
        .TextMatrix(.Rows - 1, 6) = fldata("dresc_proj") & "  -  " & pj(0)
        Else
        .TextMatrix(.Rows - 1, 6) = fldata("dresc_proj")
        End If
        pj.Close
        .TextMatrix(.Rows - 1, 7) = fldata("dresc_notes")
        fldata.MoveNext
    Wend
End With
End Sub

Public Sub bdgtcost()
rvt = Split(resourcemaster.cbo_projkey.Text, "  -  ", Len(resourcemaster.cbo_projkey.Text), vbTextCompare)
rvt1 = Split(resourcemaster.cbo_curcy.Text, "  -  ", Len(resourcemaster.cbo_curcy.Text), vbTextCompare)
Dim bd As New ADODB.Recordset
If bd.State Then bd.Close
bd.Open "select * from cost where bd_resccode='" & resourcemaster.cbo_resccode.Text & "' and bd_projectkey='" & rvt(0) & "' and bd_costtype='B'    and bd_curr='" & rvt1(0) & "' ", Cn, 3, 2
While Not bd.EOF
bd!bd_unitrate = resourcemaster.txt_rate.Text
bd.Update
bd.MoveNext
Wend
End Sub
