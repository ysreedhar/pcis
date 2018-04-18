VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form costcode 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cost Code"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7858
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
      TabCaption(0)   =   "Cost Code"
      TabPicture(0)   =   "costcode.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "costcode.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4150
         Left            =   -75000
         TabIndex        =   10
         Top             =   300
         Width           =   8775
         Begin VB.TextBox txt_notes 
            Height          =   3255
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   11
            Top             =   360
            Width           =   8055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   8775
         Begin VB.TextBox txt_costcode 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   3255
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00FFFFFF&
            Height          =   2985
            Left            =   3480
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txt_desc 
            Height          =   2445
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   4
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resource Code - Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "CostCode Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2895
         End
      End
   End
   Begin MSComCtl2.DTPicker DTP_tdate 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67174401
      CurrentDate     =   38733
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Date"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   8280
      Width           =   1230
   End
End
Attribute VB_Name = "costcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
DTP_tdate.Value = Format(Date, "dd/MM/yyyy")
Dim lst As New ADODB.Recordset
If lst.State Then lst.Close
lst.Open "select DISTINCT(resc_code),resc_desc from resourcemaster order by resc_code", Cn, 3, 2
While Not lst.EOF
List1.AddItem lst(0) & "  -  " & lst(1)
lst.MoveNext
Wend
End Sub
 
Private Sub txt_costcode_Change()
Dim t As Integer
t = 0
For t = 0 To List1.ListCount - 1
List1.Selected(t) = False
Next
 Dim g As String
 
Dim k As Integer
k = 0
For k = 0 To List1.ListCount - 1
g = Mid(List1.List(k), 2, 3)
If Mid(txt_costcode.Text, 1, 3) = g Then
List1.Selected(k) = True
End If
Next k
 

End Sub

Private Sub txt_costcode_KeyPress(KeyAscii As Integer)
Dim t As Integer
t = 0
For t = 0 To List1.ListCount - 1
List1.Selected(t) = False
Next
 Dim g As String
 
Dim k As Integer
k = 0
For k = 0 To List1.ListCount - 1
g = Mid(List1.List(k), 2, 3)
If Mid(txt_costcode.Text, 1, 3) = g Then
List1.Selected(k) = True
End If
Next k
End Sub
