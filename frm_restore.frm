VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_restore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "<<   Restore   >>"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton command1 
         Caption         =   "........................"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtfilename 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   720
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   480
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frm_restore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub Command1_Click()
 On Error Resume Next
 
cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtfilename = cdOpen.FileName
    End If

End Sub

Private Sub Command2_Click()
Call connect
On Error Resume Next
Err.Clear
dt = Split(txtfilename.Text, ".", Len(txtfilename), vbTextCompare)
''''''Sql = "RESTORE FILELISTONLY   FROM DISK =" & txtfilename & ""
'''''Sql = " RESTORE DATABASE PCMS2 "
'''''Sql = Sql & " FROM DISK = " & txtfilename & ""
'''''Sql = Sql & " WITH REPLACE,"
'''''Sql = Sql & "Move 'PCMS2_Data' To 'C:\Program Files\Microsoft SQL Server\MSSQL\data\PCMS2_Data.MDF',"
'''''Sql = Sql & "Move 'PCMS2_Log' To 'C:\Program Files\Microsoft SQL Server\MSSQL\data\PCMS_Log.LDF'"
''''''Sql = Sql & " With Move 'pcms2_data' TO pcms2_data " & "db.mdf',"
''''''Sql = Sql & " Move 'pcms2_log' TO pcms2_log " & "db.ldf'"
'''''''''''''''''Sql = "RESTORE FILELISTONLY"
'''''''''''''''''Sql = Sql & "  FROM DISK = '" & dt(0) & ".bak'"
'''''''''''''''''Sql = Sql & "  RESTORE DATABASE PCMSTEMP1"
'''''''''''''''''Sql = Sql & "  FROM DISK = '" & dt(0) & ".bak'"
'''''''''''''''''Sql = Sql & "  WITH REPLACE,"
'''''''''''''''''Sql = Sql & "  Move 'PCMSTEMP1_Data' TO 'C:\Program Files\Microsoft SQL Server\MSSQL\data\PCMSTEMP1_Data.MDF',"
'''''''''''''''''Sql = Sql & "  Move 'PCMSTEMP1_Log' TO 'C:\Program Files\Microsoft SQL Server\MSSQL\data\PCMSTEMP1_Log.LDF'"
 
'''''''''''''''''Cn.Execute Sql







 
If Err.Number = 0 Then MsgBox "Restore Succeded", vbInformation
End Sub

 

