VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form resourcemaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resource Item"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Resource Master"
      TabPicture(0)   =   "resourcemaster.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resource Details"
      TabPicture(1)   =   "resourcemaster.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   -75000
         TabIndex        =   32
         Top             =   360
         Width           =   10455
         Begin VB.ComboBox cbo_projkey 
            Height          =   315
            Left            =   1080
            TabIndex        =   43
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox cbo_vendor1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6720
            TabIndex        =   38
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txt_standardrate1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6720
            TabIndex        =   37
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox txt_rescourcecode1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   36
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txt_resourcedesc1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   525
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox cbo_uom1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   6720
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Key"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5280
            TabIndex        =   33
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5280
            TabIndex        =   42
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5280
            TabIndex        =   41
            Top             =   480
            Width           =   960
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Desc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   -75000
         TabIndex        =   17
         Top             =   1920
         Width           =   10455
         Begin VB.ComboBox cbo_resccode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txt_notes 
            Height          =   375
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1080
            Width           =   9495
         End
         Begin VB.TextBox txt_rate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cbo_curcy 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4200
            TabIndex        =   20
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox txt_ratetype 
            Height          =   315
            Left            =   7080
            TabIndex        =   19
            Top             =   480
            Width           =   1095
         End
         Begin VB.ComboBox DTP_resc 
            Height          =   315
            Left            =   2760
            TabIndex        =   18
            Top             =   480
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTP_tdate1 
            Height          =   315
            Left            =   8280
            TabIndex        =   24
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64225281
            CurrentDate     =   38733
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Notes"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Type"
            Height          =   195
            Left            =   7080
            TabIndex        =   29
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            Height          =   195
            Left            =   5640
            TabIndex        =   28
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4200
            TabIndex        =   27
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Year"
            Height          =   255
            Left            =   2760
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   8280
            TabIndex        =   25
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   0
         TabIndex        =   8
         Top             =   300
         Width           =   10455
         Begin VB.ComboBox cbo_resp 
            Height          =   315
            Left            =   5520
            TabIndex        =   4
            Top             =   1080
            Width           =   4335
         End
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   5520
            TabIndex        =   3
            Top             =   360
            Width           =   4335
         End
         Begin VB.ComboBox txt_standardrate 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   1800
            Width           =   4335
         End
         Begin VB.ComboBox cbo_uom 
            Height          =   315
            Left            =   5520
            TabIndex        =   5
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txt_rescourcecode 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   435
            Width           =   3615
         End
         Begin VB.TextBox txt_resourcedesc 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   1080
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   7200
            TabIndex        =   6
            Top             =   1755
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64225281
            CurrentDate     =   38733
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Responsible Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   16
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   15
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   14
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   11
            Top             =   900
            Width           =   1560
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7200
            TabIndex        =   10
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   9
            Top             =   1560
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "resourcemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

 

Private Sub cbo_curcy_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_projkey_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub cbo_resp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cbo_uom_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cbo_vendor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub DTP_resc_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
cbo_projkey.Clear
cbo_uom.Clear
cbo_vendor.Clear
cbo_resp.Clear
txt_standardrate.Clear

DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not pk.EOF
cbo_projkey.AddItem pk(0) & "  -  " & pk(1)
pk.MoveNext
Wend
pk.Close
Dim pk1 As New ADODB.Recordset
If pk1.State Then pk1.Close
pk1.Open "select DISTINCT(uom_uom),uom_desc from UOM order by uom_uom", Cn, 3, 2
While Not pk1.EOF
cbo_uom.AddItem pk1(0) & "  -  " & pk1(1)
pk1.MoveNext
Wend
pk1.Close
Dim vn1 As New ADODB.Recordset
If vn1.State Then vn1.Close
vn1.Open "select DISTINCT(vendor_code),vendor_desc from vendormaster order by vendor_code", Cn, 3, 2
While Not vn1.EOF
cbo_vendor.AddItem vn1(0) & "  -  " & vn1(1)
vn1.MoveNext
Wend
vn1.Close
Dim rp As New ADODB.Recordset
If rp.State Then rp.Close
rp.Open "select DISTINCT(resp_code),resp_desc from responsiblemaster order by resp_code", Cn, 3, 2
While Not rp.EOF
cbo_resp.AddItem rp(0) & "  -  " & rp(1)
rp.MoveNext
Wend
rp.Close
Dim rty As New ADODB.Recordset
If rty.State Then rty.Close
rty.Open "select DISTINCT(r_type),r_desc from resourcetype order by r_type", Cn, 3, 2
While Not rty.EOF
txt_standardrate.AddItem rty(0) & "  -  " & rty(1)
rty.MoveNext
Wend
rty.Close

frm_resourcemaster.Toolbar1.Buttons(5).Enabled = True
frm_resourcemaster.Toolbar1.Buttons(7).Enabled = True
frm_resourcemaster.Toolbar2.Buttons(1).Enabled = False
frm_resourcemaster.Toolbar2.Buttons(3).Enabled = False
frm_resourcemaster.Toolbar2.Buttons(5).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_resourcemaster.Toolbar1.Buttons(1).Enabled = True
frm_resourcemaster.Toolbar1.Buttons(5).Enabled = False
frm_resourcemaster.Toolbar1.Buttons(7).Enabled = False

frm_resourcemaster.Toolbar2.Buttons(1).Enabled = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "Resource Details" Then
DTP_tdate1.Value = Format(Date, "dd/MM/yyyy")
Dim k1 As Integer
k1 = 0
For k1 = 2000 To 2050
DTP_resc.AddItem k1
Next k1
'loading resource master deatils

 
txt_rescourcecode1.Text = resourcemaster.txt_rescourcecode
txt_standardrate1.Text = resourcemaster.txt_standardrate
cbo_vendor1.Text = resourcemaster.cbo_vendor.Text
cbo_uom1.Text = resourcemaster.cbo_uom.Text
txt_resourcedesc1.Text = resourcemaster.txt_resourcedesc
txt_ratetype.AddItem "BR"
txt_ratetype.AddItem "CR"
cbo_resccode.Text = txt_rescourcecode.Text

Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select DISTINCT(c_name),c_desc from currency order by c_name", Cn, 3, 2
While Not cr.EOF
cbo_curcy.AddItem cr(0) & "  -  " & cr(1)
cr.MoveNext
Wend
cr.Close
 frm_resourcemaster.Toolbar1.Buttons(1).Enabled = False
 frm_resourcemaster.Toolbar1.Buttons(3).Enabled = False
 frm_resourcemaster.Toolbar1.Buttons(5).Enabled = False
 frm_resourcemaster.Toolbar1.Buttons(7).Enabled = False
 frm_resourcemaster.Toolbar2.Buttons(1).Enabled = True
Else
frm_resourcemaster.Toolbar1.Buttons(5).Enabled = True
frm_resourcemaster.Toolbar1.Buttons(7).Enabled = True
frm_resourcemaster.Toolbar2.Buttons(1).Enabled = False
frm_resourcemaster.Toolbar2.Buttons(3).Enabled = False
frm_resourcemaster.Toolbar2.Buttons(5).Enabled = False
End If
End Sub

 
Private Sub txt_ratetype_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txt_standardrate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
