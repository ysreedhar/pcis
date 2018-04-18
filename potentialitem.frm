VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form potentialitem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Potential Item"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
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
      TabCaption(0)   =   "PI"
      TabPicture(0)   =   "potentialitem.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "potentialitem.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   -75000
         TabIndex        =   12
         Top             =   300
         Width           =   5415
         Begin VB.TextBox txt_notes 
            Height          =   2175
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   13
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   5415
         Begin VB.ComboBox cbo_year 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txt_cost 
            Height          =   300
            Left            =   2280
            TabIndex        =   8
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txt_revn 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txt_code 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   1200
            Width           =   1935
         End
         Begin VB.ComboBox cbo_item 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   4335
         End
         Begin MSComCtl2.DTPicker DTP_tdate 
            Height          =   315
            Left            =   3480
            TabIndex        =   10
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   64946177
            CurrentDate     =   38733
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Year"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   3480
            TabIndex        =   11
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cost"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Revn"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PI Item Code"
            Height          =   255
            Left            =   1320
            TabIndex        =   5
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item - Description"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "potentialitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(i_item),i_desc from ohpi_itemmaster where i_type='PI' order by i_item", Cn, 3, 2
While Not rs.EOF
cbo_item.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close
Dim i As Integer
i = 0
For i = 2000 To 2050
cbo_year.AddItem i
Next
End Sub

