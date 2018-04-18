VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmL0Notes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "L0 Notes & Signatures"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6588
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
      TabCaption(0)   =   "Signatures"
      TabPicture(0)   =   "frmL0Notes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "frmL0Notes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   -75000
         TabIndex        =   4
         Top             =   300
         Width           =   5415
         Begin VB.TextBox txt_notes 
            Height          =   2535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   5
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   5415
         Begin VB.TextBox txtApprovedBy 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1635
            Width           =   4335
         End
         Begin VB.TextBox txtReviewedBy 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   915
            Width           =   4335
         End
         Begin VB.TextBox txtPreparedBy 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   315
            Width           =   4335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reviewed By:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prepared By:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmL0Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
Dim rsNotesUpdate As New ADODB.Recordset
If rsNotesUpdate.State Then rsNotesUpdate.Close
rsNotesUpdate.Open "select * from tblL0Notes", Cn, 3, 2
If Not rsNotesUpdate.EOF Then
rsNotesUpdate!notes = txt_notes.Text
rsNotesUpdate!PreparedBy = txtPreparedBy.Text
rsNotesUpdate!ReviewedBy = txtReviewedBy.Text
rsNotesUpdate!ApprovedBy = txtApprovedBy.Text
rsNotesUpdate.Update
rsNotesUpdate.Close
MsgBox "L0 Report Notes Modified"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from tblL0Notes", Cn, 3, 2
While Not rs.EOF
txtPreparedBy.Text = rs(2)
txtReviewedBy.Text = rs(3)
txtApprovedBy.Text = rs(4)
txt_notes.Text = rs(1)
rs.MoveNext
Wend
End Sub

