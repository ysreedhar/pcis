VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_export 
   BackColor       =   &H00FFFFFF&
   Caption         =   "IMPORT BUDGET"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form3"
   ScaleHeight     =   10125
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IMPORT EIC"
      Height          =   975
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   9495
      Begin VB.CommandButton cmd_check 
         BackColor       =   &H00FFC0C0&
         Caption         =   "check"
         Height          =   375
         Left            =   7245
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Height          =   375
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Import"
         Height          =   375
         Left            =   6195
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtfilename 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select File"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Export"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38378
      End
      Begin MSComCtl2.DTPicker dtp_to 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38378
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "To"
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   195
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_individualmember 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   70
      FixedCols       =   0
      RowHeightMin    =   350
      ForeColor       =   12582912
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   8454143
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   9015
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   15901
      _Version        =   393216
      Cols            =   70
      FixedCols       =   0
      RowHeightMin    =   350
      ForeColor       =   12582912
      BackColorFixed  =   12582912
      ForeColorFixed  =   16777215
      BackColorSel    =   8388608
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "After you perform ""CHECK"" ,  Inconsistent entries will be displayed with RED background."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9960
      TabIndex        =   15
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Maximum of 500 Rows allowed per each transaction."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9960
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "frm_export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public cx As Integer
Public asd As Integer

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub cmd_check_Click()
Call data_check
If asd = 0 Then
Command3.Enabled = True
End If

End Sub

Private Sub cmd_load_Click()

Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
AppActivate "FlexGrid To Excel"
For i = 0 To flex_individualmember.Rows - 1
    flex_individualmember.Row = i
    For n = 0 To 69
        flex_individualmember.Col = n
        objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flex_individualmember.Text
    Next
Next

End Sub

Private Sub Command1_Click()
'On Error Resume Next
 
cdOpen.ShowOpen
    
    If Not vbCancel Then
       txtfilename = cdOpen.FileName
    End If
Command2.Enabled = True
End Sub

Private Sub command2_Click()
On Error Resume Next
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
'Set objWorkbook = objExcel.Workbooks.Open(App.Path & "\test.xls")
Set objWorkbook = objExcel.Workbooks.Open(txtfilename.Text)
Set objWorksheet = objWorkbook.ActiveSheet

With MSFlexGrid1
.Cols = 30
.Rows = 500
For i = 0 To .Rows - 1
    .Row = i
    For n = 0 To .Cols - 1
        .Col = n
        .Text = objWorksheet.Cells(i + 1, n + 1).Value
    Next
Next
End With

AppActivate Me.Caption
MsgBox "Imported Successfully"
cmd_check.Enabled = True
End Sub

Private Sub Command3_Click()
Call data_flex
MsgBox "Saved Successfully"
End Sub

Private Sub Command4_Click()
frm_import.Show
Unload Me
End Sub

Private Sub dtp_from_Change()
flex_individualmember.Clear
 
End Sub

Private Sub dtp_to_Change()
flex_individualmember.Clear
 
End Sub

Private Sub Form_Load()
dtp_from.Value = Format(Date, "dd/MM/yyyy")
dtp_to.Value = Format(Date, "dd/MM/yyyy")
 
Command3.Enabled = False
cmd_check.Enabled = False
Command2.Enabled = False
End Sub


Public Sub data_flex()
On Error Resume Next
Dim p As Double
p = 0
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from cost ", Cn, 3, 2
For p = 1 To MSFlexGrid1.Rows
   If MSFlexGrid1.TextMatrix(p, 1) = "" Then Exit Sub
        
        fldata.AddNew
        fldata!bd_year = MSFlexGrid1.TextMatrix(p, 0)
        fldata!bd_projectkey = MSFlexGrid1.TextMatrix(p, 1)
        Dim res As New ADODB.Recordset
        If res.State Then res.Close
        res.Open "select proj_desc from projectmaster where proj_key='" & MSFlexGrid1.TextMatrix(p, 1) & "' ", Cn, 3, 2
        If Not res.EOF Then
        fldata!bd_projectdesc = res(0)
        End If
        res.Close
        
        If MSFlexGrid1.TextMatrix(p, 3) = "" Then
        fldata!bd_resccode = "R" & Mid(MSFlexGrid1.TextMatrix(p, 6), 3) & "A"
        Else
        
        fldata!bd_resccode = MSFlexGrid1.TextMatrix(p, 3)
        
        End If
        
        res.Open "select * from resourcemaster  where resc_code='" & MSFlexGrid1.TextMatrix(p, 3) & "'", Cn, 3, 2
        If Not res.EOF Then
        fldata!bd_rescname = res!resc_desc
        fldata!bd_vendor = res!resc_vendorcode
        fldata!bd_costtype = "B"
        fldata!bd_respcode = res!resc_respcode
        fldata!bd_respname = "To be Advised"
        fldata!bd_brate = 0
        fldata!bd_crate = 0
        End If
        
        
        fldata!bd_spread = MSFlexGrid1.TextMatrix(p, 4)
        If MSFlexGrid1.TextMatrix(p, 4) = "NA" Then
        fldata!bd_tranx = "ME"
        Else
        fldata!bd_tranx = "SD"
        End If
        fldata!bd_jobcharge = MSFlexGrid1.TextMatrix(p, 5)
        fldata!bd_costcode = MSFlexGrid1.TextMatrix(p, 6)
        fldata!bd_qty = MSFlexGrid1.TextMatrix(p, 7)
        fldata!bd_days = MSFlexGrid1.TextMatrix(p, 8)
        fldata!bd_tqty = MSFlexGrid1.TextMatrix(p, 9)
        fldata!bd_uom = MSFlexGrid1.TextMatrix(p, 10)
        fldata!bd_curr = MSFlexGrid1.TextMatrix(p, 11)
        fldata!bd_unitrate = MSFlexGrid1.TextMatrix(p, 12)
        fldata!bd_xchg = MSFlexGrid1.TextMatrix(p, 13)
        fldata!bd_downtime = MSFlexGrid1.TextMatrix(p, 14)
        fldata!bd_escl = MSFlexGrid1.TextMatrix(p, 15)
        fldata!bd_extdamt = MSFlexGrid1.TextMatrix(p, 16)
        fldata!bd_wrkcomp = MSFlexGrid1.TextMatrix(p, 17)
        fldata!bd_bcwpamt = MSFlexGrid1.TextMatrix(p, 18)
        fldata!bd_notes = MSFlexGrid1.TextMatrix(p, 19)
        fldata!t_date = Format(Date, "dd/MM/yyyy")
        fldata!u_date = Now
        fldata!t_user = main.Label2.Caption
        fldata!bd_obs = "XX"
        fldata.Update
     'End If
Next p
End Sub

Public Sub data_check()

asd = 0
Dim ch As Double
ch = 0
Dim chf As New ADODB.Recordset

For ch = 1 To MSFlexGrid1.Rows - 1
    If MSFlexGrid1.TextMatrix(ch, 1) <> "" Then
    If chf.State Then chf.Close
    chf.Open "select * from resourcedetails where dresc_proj='" & MSFlexGrid1.TextMatrix(ch, 1) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 1
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    chf.Open "select * from resourcedetails where dresc_code='" & MSFlexGrid1.TextMatrix(ch, 3) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 3
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    If MSFlexGrid1.TextMatrix(ch, 4) <> "NA" Then
    chf.Open "select * from budgeteddurationdetails where bdgt_spread_code='" & MSFlexGrid1.TextMatrix(ch, 4) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    End If
    If MSFlexGrid1.TextMatrix(ch, 4) <> "NA" Then
    chf.Open "select * from budgeteddurationdetails where bdgt_job_key='" & MSFlexGrid1.TextMatrix(ch, 5) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 5
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    End If
    chf.Open "select * from costcode where cc_code='" & MSFlexGrid1.TextMatrix(ch, 6) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 6
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 7)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 9)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 9
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    chf.Open "select * from resourcemaster where resc_uom='" & MSFlexGrid1.TextMatrix(ch, 10) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 10
    MSFlexGrid1.CellBackColor = vbRed
    End If
    
    chf.Close
    chf.Open "select * from currencymaster where cur_currency='" & MSFlexGrid1.TextMatrix(ch, 11) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellBackColor = vbRed
    End If
    chf.Close
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 12)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 12
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 11)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellBackColor = vbRed
    Else
    chf.Open "select * from currencymaster where cur_xchgrate='" & MSFlexGrid1.TextMatrix(ch, 11) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellBackColor = vbRed
    End If
    End If
    chf.Close
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 14)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 14
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 15)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 15
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 16)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 16
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 17)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 17
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 18)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 18
    MSFlexGrid1.CellBackColor = vbRed
    End If
    End If
Next ch
MsgBox "Check Completed!", vbInformation, App.Title
End Sub
