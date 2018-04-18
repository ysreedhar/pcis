VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "NetWork Utility"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   9180
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
            Object.ToolTipText     =   "Status Information"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:03 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "24/05/2005"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Send Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   2655
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   7440
      Width           =   8895
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12938
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   4935
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get User List"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "dmmac"
            Object.Tag             =   "dmmac"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0894
            Key             =   "cmac"
            Object.Tag             =   "cmac"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CE6
            Key             =   "dm"
            Object.Tag             =   "dm"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If tvw.SelectedItem.Selected Then
    Call LoadListBox
Else
    MsgBox "Please select a server first."
End If
End Sub

Private Sub Command2_Click()
Dim strMessage As String, strPCName As String
Dim i As Long
If tvw.SelectedItem.Selected = False Then
    MsgBox "Receipient Must Be Selected!", 0, "Select Receipient Name"
    Exit Sub
End If

If txtMessage.Text = "" Then
    MsgBox "You must enter a message!", 0, "Enter Message"
    Exit Sub
End If

strPCName = Trim(tvw.SelectedItem.Text)
strMessage = "net send " & strPCName & " " & txtMessage.Text
StatusBar1.Panels(1).Text = "Sending Message to: " & strPCName
Screen.MousePointer = vbArrowHourglass
'Send Message
'* There is also function called NetSend you can use this by sending an API your choice
'blnset = NetSend(txtMessage.Text, Node)
i = Shell(strMessage)

StatusBar1.Panels(1).Text = "Message Send to: " & strPCName
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbArrowHourglass
StatusBar1.Panels(1).Text = "Please Wait Querying Domains"
SERVERTYPE = SV_TYPE_ALL 'SV_TYPE_SQLSERVER ' '* set the types
Call FillDomainTree(SV_TYPE_DOMAIN_ENUM, Me.tvw) '* fill the tree view
Screen.MousePointer = vbDefault
Me.Top = 5
Me.Left = 5
End Sub

Public Sub LoadListBox()
Dim i As Integer
Dim NumUsers As Long
Dim strServerName As String

strServerName = "\\" & Trim(tvw.SelectedItem.Text)


NumUsers = GetUsers(strServerName) 'For local users use "" as Server Parameter
    'Fill the List
    List1.Clear
    For i = 0 To NumUsers - 1
        List1.AddItem UserInfo(i).Name & " - " & UserInfo(i).Comment
    Next i
    If NumUsers = 0 Then
        MsgBox "Please check domain Name"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()
MsgBox List1.Text
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
Command2.Caption = "Send Message to: " & tvw.SelectedItem.Text
End Sub
