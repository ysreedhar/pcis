VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form resourcedetails 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   10095
      Begin VB.TextBox cbo_uom1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6840
         TabIndex        =   28
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txt_resourcedesc1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txt_rescourcecode1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txt_standardrate1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6840
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox cbo_projkey1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox cbo_vendor1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6840
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Desc"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Code"
         Height          =   195
         Left            =   5400
         TabIndex        =   14
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UOM"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Type"
         Height          =   195
         Left            =   5400
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Resource Details"
      TabPicture(0)   =   "resourcedetails.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9975
         Begin VB.ComboBox DTP_resc 
            Height          =   315
            Left            =   3000
            TabIndex        =   0
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox txt_ratetype 
            Height          =   315
            Left            =   8400
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cbo_curcy 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4680
            TabIndex        =   1
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txt_rate 
            Height          =   285
            Left            =   6720
            TabIndex        =   2
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txt_notes 
            Height          =   615
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   4
            Top             =   1440
            Width           =   7815
         End
         Begin VB.ComboBox cbo_resccode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            TabIndex        =   19
            Top             =   600
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   8400
            TabIndex        =   26
            Top             =   1515
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57016321
            CurrentDate     =   38733
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   8400
            TabIndex        =   27
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Label7 
            Caption         =   "Year"
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   24
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            Height          =   195
            Left            =   6720
            TabIndex        =   23
            Top             =   360
            Width           =   345
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Type"
            Height          =   195
            Left            =   8400
            TabIndex        =   22
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label11 
            Caption         =   "Notes"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "resourcedetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim k As Integer
k = 0
For k = 2000 To 2050
DTP_resc.AddItem k
Next k
'loading resource master deatils

resourcedetails.cbo_projkey.Text = resourcemaster.cbo_projkey
resourcedetails.txt_rescourcecode.Text = resourcemaster.txt_rescourcecode
resourcedetails.txt_standardrate.Text = resourcemaster.txt_standardrate
resourcedetails.cbo_vendor.Text = resourcemaster.cbo_vendor.Text
resourcedetails.cbo_uom.Text = resourcemaster.cbo_uom.Text
resourcedetails.txt_resourcedesc.Text = resourcemaster.txt_resourcedesc

txt_ratetype.AddItem "BR"
txt_ratetype.AddItem "CR"

cbo_resccode.Text = txt_rescourcecode.Text
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from resourcemaster", Cn, 3, 2
While Not rs.EOF
cbo_resccode.AddItem rs("resc_code")
rs.MoveNext
Wend
Dim cr As New ADODB.Recordset
If cr.State Then cr.Close
cr.Open "select * from currencymaster", Cn, 3, 2
While Not cr.EOF
cbo_curcy.AddItem cr!cur_currency
cr.MoveNext
Wend


End Sub

