VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_import 
   BackColor       =   &H00FFFFFF&
   Caption         =   "IMPORT EIC / BUDGET"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   LinkTopic       =   "Form2"
   ScaleHeight     =   9315
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optImportBudget 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Import Budget"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton optImportEIC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Import EIC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtEditFlexGrid 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   13440
      TabIndex        =   15
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Export"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   375
         Left            =   720
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38378
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "To"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select File"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtfilename 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Import"
         Height          =   375
         Left            =   6195
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Height          =   375
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_check 
         BackColor       =   &H00FFC0C0&
         Caption         =   "check"
         Height          =   375
         Left            =   7245
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   9015
      Left            =   120
      TabIndex        =   13
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
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flex_individualmember 
      Height          =   2655
      Left            =   120
      TabIndex        =   12
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
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   8
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "After you perform ""CHECK"" ,  Inconsistent entries will be displayed with RED background."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   9960
      TabIndex        =   14
      Top             =   720
      Width           =   5055
   End
End
Attribute VB_Name = "frm_import"
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
Function IsValidData() As Boolean
'ColumnArray() = strArrRemoveDuplicate(FlxColArray(1))
'IsFlxColArray (1)
IsValidProjectName (ColumnArray)
End Function
Public Sub ApplyColorToGridColumn(ByRef gGrid As MSFlexGrid, ByRef Col_Num As Long, ByRef ColorToApply As Long)
     Dim X  As Long
     For X = 1 To gGrid.Rows - 1
         gGrid.Col = Col_Num
         gGrid.Row = X
         gGrid.CellBackColor = ColorToApply
     Next
End Sub
Function RepaintFlexGrid()
' Reset the backcolor
For ch = 1 To MSFlexGrid1.Rows - 1
For flxcls = 0 To MSFlexGrid1.Cols - 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = flxcls
    MSFlexGrid1.CellBackColor = vbWhite
Next flxcls
Next ch
End Function
Function IsValidProjectName(strProjectArray As String) As Boolean
Set cmd = New ADODB.Command
Dim rs As ADODB.Recordset
If Cn.State Then Cn.Close
Cn.Open
cmd.ActiveConnection = Cn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "get_Projects_iter"
cmd.parameters("@Projects").Value = strProjectArray
Set rs = cmd.Execute
 If Not rs.EOF Then
 SearchGridColumn rs(2), 1
 End If
Set cmd.ActiveConnection = Nothing
End Function
Private Sub cmd_check_Click()
RepaintFlexGrid
If optImportBudget.Value Then
Call data_checkforBudget
ElseIf optImportEIC.Value Then
Call data_checkforEIC
End If
'IsValidData
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
strComputer = "."
On Error Resume Next
Dim i As Long
Dim n As Long
On Error Resume Next
If optImportEIC.Value = False And optImportBudget.Value = False Then
MsgBox "Select Import Type", vbInformation, App.Title
Exit Sub
End If
response = MsgBox("You are about to run an Import which will close all active excel Workbooks. Do you wish to continue?", 36, "PCIS - ATTENTION!")
If response = 6 Then
'Set objExcel = GetObject(, "Excel.Application")
 Set objExcel = GetObject(, "Microsoft Excel.application")
   If Err = 429 Then
       Err = 0
       Set objExcel = GetObject(, "Excel.application")
   End If
If Err.Number Then
   Err.Clear
     Set objExcel = CreateObject("Microsoft Excel.application")
   If Err = 429 Then
       Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   End If
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
'Set objWorkbook = objExcel.Workbooks.Open(App.Path & "\test.xls")
Set objWorkbook = objExcel.Workbooks.Open(txtfilename.Text)
Set objWorksheet = objWorkbook.ActiveSheet
' find the number of rows required
 Set foundcell = objExcel.ActiveSheet.Cells.Find("End of Document")
 rowcolumn = foundcell.Address
 If rowcolumn <> Empty Then
 intRowNumber = CDbl(Mid(rowcolumn, 4, Len(rowcolumn) - 3))
 End If
With MSFlexGrid1
.Cols = 27
.Rows = intRowNumber - 1
For i = 0 To .Rows - 1
    .Row = i
    For n = 0 To .Cols - 1
        .Col = n
        If objWorksheet.Cells(i + 1, n + 1).Value <> "End of Document" Then .Text = objWorksheet.Cells(i + 1, n + 1).Value Else GoTo DisplayMessage
    Next
Next
GoTo DisplayMessage
End With
DisplayMessage:
On Error Resume Next
            Do   ''''''start a loop to wait for all instances of excel to close before continuing'''''
                ''''''find excel in task manager process list''''''
                Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                Set colprocesslist = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'Excel.exe'")
                X = False
                '''''if excel is open change x to true'''''''
                For Each objprocess In colprocesslist
                X = True
                'MsgBox objprocess.Name
                objprocess.Terminate
                Next
                ''''''if x is false then no excel apps are open so we can carry on with the update''''''''''
                If X = False Then Exit Do
            Loop
            
'objWorkbook.Close
'objExcel.Application.Quit
Set objExcel = Nothing
AppActivate Me.Caption
MsgBox "Imported Successfully"
'MSFlexGrid1.Rows = i
cmd_check.Enabled = True
End If
End Sub

Private Sub Command3_Click()
RepaintFlexGrid
'Call data_flex
If optImportBudget.Value Then
InsertFromFlexBudget
ElseIf optImportEIC.Value Then
InsertFromFlexEIC
End If
MsgBox "Saved Successfully"
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
Function FlxColArray(ByVal ColNum As Integer) As String
Dim TextFromColumn As String
Dim ColumnArray() As String
 With MSFlexGrid1
 .Row = .FixedRows
 .Col = ColNum
 .RowSel = .Rows - 1
 .ColSel = ColNum
 TextFromColumn = .Clip
End With
'ColumnArray = Split(TextFromColumn, vbCr)
strColumnArray = Replace(TextFromColumn, vbCr, ",")
'StringArray = strColumnArray
ColumnArray = Split(strColumnArray, ",")
ColumnArray = strArrRemoveDuplicate(ColumnArray)
End Function
Private Sub SearchGridColumn(target_name As String, ColumnNumber As Integer)
Dim r As Integer
    ' Search for the name, skipping the column heading row.
    For r = 1 To MSFlexGrid1.Rows - 1
        If LCase$(MSFlexGrid1.TextMatrix(r, 0)) = target_name Then
            ' We found the target. Select this row.
            MSFlexGrid1.Row = r
            MSFlexGrid1.Col = ColumnNumber
            MSFlexGrid1.CellBackColor = vbRed
            Exit Sub
        End If
    Next r
End Sub
Public Sub InsertFromFlexEIC()
'On Error Resume Next
Dim p As Integer
For p = 1 To MSFlexGrid1.Rows - 1
If MSFlexGrid1.TextMatrix(p, 1) = "" Then Exit Sub
InsertNewResourceEIC MSFlexGrid1.TextMatrix(p, 0), MSFlexGrid1.TextMatrix(p, 1), MSFlexGrid1.TextMatrix(p, 3), MSFlexGrid1.TextMatrix(p, 4), MSFlexGrid1.TextMatrix(p, 5), MSFlexGrid1.TextMatrix(p, 6), MSFlexGrid1.TextMatrix(p, 7), MSFlexGrid1.TextMatrix(p, 8), MSFlexGrid1.TextMatrix(p, 9), MSFlexGrid1.TextMatrix(p, 10), MSFlexGrid1.TextMatrix(p, 11), MSFlexGrid1.TextMatrix(p, 12), MSFlexGrid1.TextMatrix(p, 13), MSFlexGrid1.TextMatrix(p, 14), MSFlexGrid1.TextMatrix(p, 15), MSFlexGrid1.TextMatrix(p, 9), MSFlexGrid1.TextMatrix(p, 23), MSFlexGrid1.TextMatrix(p, 24), MSFlexGrid1.TextMatrix(p, 25), MSFlexGrid1.TextMatrix(p, 26)
 For i = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = p
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbYellow
Next i
Next p
End Sub
Public Sub InsertFromFlexBudget()
'On Error Resume Next
Dim p As Integer
For p = 1 To MSFlexGrid1.Rows - 1
If MSFlexGrid1.TextMatrix(p, 1) = "" Then Exit Sub
InsertNewResourceBudget MSFlexGrid1.TextMatrix(p, 0), MSFlexGrid1.TextMatrix(p, 1), MSFlexGrid1.TextMatrix(p, 3), MSFlexGrid1.TextMatrix(p, 4), MSFlexGrid1.TextMatrix(p, 5), MSFlexGrid1.TextMatrix(p, 6), MSFlexGrid1.TextMatrix(p, 9), MSFlexGrid1.TextMatrix(p, 10), MSFlexGrid1.TextMatrix(p, 11), MSFlexGrid1.TextMatrix(p, 12), MSFlexGrid1.TextMatrix(p, 13), MSFlexGrid1.TextMatrix(p, 14), MSFlexGrid1.TextMatrix(p, 15), MSFlexGrid1.TextMatrix(p, 17), MSFlexGrid1.TextMatrix(p, 20), MSFlexGrid1.TextMatrix(p, 18), MSFlexGrid1.TextMatrix(p, 21), MSFlexGrid1.TextMatrix(p, 26)
 For i = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = p
        MSFlexGrid1.Col = i
        MSFlexGrid1.CellBackColor = vbYellow
Next i
Next p
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
        If res.State Then res.Close
        res.Open "select * from resourcemaster  where resc_code='" & MSFlexGrid1.TextMatrix(p, 3) & "'", Cn, 3, 2
        If Not res.EOF Then
        fldata!bd_rescname = res!resc_desc
        fldata!bd_vendor = res!resc_vendorcode
        fldata!bd_costtype = "E"
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
        fldata!bd_sdate = MSFlexGrid1.TextMatrix(p, 7)
        fldata!bd_edate = MSFlexGrid1.TextMatrix(p, 8)
        fldata!bd_type = MSFlexGrid1.TextMatrix(p, 9)
        fldata!bd_uom = MSFlexGrid1.TextMatrix(p, 10)
        fldata!bd_curr = MSFlexGrid1.TextMatrix(p, 11)
        fldata!bd_unitrate = MSFlexGrid1.TextMatrix(p, 12)
        fldata!bd_xchg = MSFlexGrid1.TextMatrix(p, 13)
        fldata!bd_qty = MSFlexGrid1.TextMatrix(p, 14)
        fldata!bd_days = MSFlexGrid1.TextMatrix(p, 15)
        fldata!bd_tqty = MSFlexGrid1.TextMatrix(p, 16)
        fldata!bd_extdamt = MSFlexGrid1.TextMatrix(p, 17)
        fldata!bd_e_days = MSFlexGrid1.TextMatrix(p, 18)
        fldata!bd_e_tqty = MSFlexGrid1.TextMatrix(p, 19)
        fldata!bd_e_extdamt = MSFlexGrid1.TextMatrix(p, 20)
        If MSFlexGrid1.TextMatrix(p, 4) = "NA" Then
                    If MSFlexGrid1.TextMatrix(p, 15) = "" And MSFlexGrid1.TextMatrix(p, 18) = "" Then
                    fldata!bd_chk1 = 1
                    fldata!bd_chk = 0
                    Else
                    fldata!bd_chk = 1
                    fldata!bd_chk1 = 0
                    End If
         Else
                    fldata!bd_chk = 1
                    fldata!bd_chk1 = 0
        End If
        fldata!bd_notes = MSFlexGrid1.TextMatrix(p, 21)
        fldata!t_date = Format(Date, "dd/MM/yyyy")
        fldata!u_date = Now
        fldata!t_user = main.Label2.Caption
        fldata!bd_obs = "XX"
        fldata.Update
     'End If
Next p
End Sub
Public Function strArrRemoveDuplicate(ByRef StringArray() As String) As String()
    Dim LowBound As Long, UpBound As Long
    Dim TempArray() As String, cur As Long
    Dim a As Long, b As Long
    'check for empty array
    If (Not StringArray()) = True Then Exit Function
    'we need these often
    LowBound = LBound(StringArray)
    UpBound = UBound(StringArray)
    
    'reserve check buffer
    ReDim TempArray(LowBound To UpBound)
    
    'set first item
    cur = LowBound
    TempArray(cur) = StringArray(LowBound)
    'loop through all items
    For a = LowBound + 1 To UpBound
        'make a comparison against all items
        For b = LowBound To cur
            'if is a duplicate, exit array
            If LenB(TempArray(b)) = LenB(StringArray(a)) Then
                If InStrB(1, StringArray(a), TempArray(b), vbBinaryCompare) = 1 Then Exit For
            End If
        Next b
        'check if the loop was exited: add new item to check buffer if not
        If b > cur Then cur = b: TempArray(cur) = StringArray(a)
    Next a
    'fix size
    ReDim Preserve TempArray(LowBound To cur)
    'copy
    StringArray = TempArray
End Function
Public Function InsertNewResourceEIC(StrYear As String, strProjectKey As String, StrResccode As String, StrSpread As String, StrJobcharge As String, StrCostcode As String, strStartDate As String, strEndDate As String, StrType As String, StrUom As String, StrCurr As String, StrUnitrate As String, StrXchg As String, StrQty As String, Strdays As String, StrExtdamt As String, StrE_days As String, StrE_tqty As String, StrE_extdamt As String, StrNotes As String)
On Error GoTo ErrH
Set cmd = New ADODB.Command
If Cn.State Then Cn.Close
Cn.Open
cmd.ActiveConnection = Cn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "spImportResourcesforEIC"
cmd.parameters("@bd_year").Value = StrYear
cmd.parameters("@bd_projectkey").Value = strProjectKey
cmd.parameters("@bd_resccode").Value = StrResccode
cmd.parameters("@bd_spread").Value = StrSpread
cmd.parameters("@bd_jobcharge").Value = StrJobcharge
cmd.parameters("@bd_costcode").Value = StrCostcode
cmd.parameters("@startDate").Value = CDate(strStartDate)
cmd.parameters("@endDate").Value = CDate(strEndDate)
cmd.parameters("@bd_type").Value = StrType
cmd.parameters("@bd_uom").Value = StrUom
cmd.parameters("@bd_curr").Value = StrCurr
cmd.parameters("@bd_unitrate").Value = CDbl(StrUnitrate)
cmd.parameters("@bd_xchg").Value = CDbl(StrXchg)
cmd.parameters("@bd_qty").Value = CDbl(StrQty)
cmd.parameters("@bd_days").Value = CDbl(Strdays)
cmd.parameters("@bd_tqty").Value = CDbl(StrQty) * CDbl(Strdays)
cmd.parameters("@bd_extdamt").Value = (CDbl(StrUnitrate) * CDbl(StrXchg) * (CDbl(StrQty) * CDbl(Strdays)))
cmd.parameters("@bd_e_days").Value = CDbl(StrE_days)
cmd.parameters("@bd_e_tqty").Value = CDbl(StrE_tqty)
cmd.parameters("@bd_e_extdamt").Value = CDbl(StrE_extdamt)
cmd.parameters("@bd_notes").Value = StrNotes
cmd.parameters("@t_user").Value = main.Label2.Caption
cmd.parameters("@result").Value = 0
cmd.Execute
intRowsInserted = intRowsInserted + cmd("@result")
Set cmd.ActiveConnection = Nothing
Exit Function
ErrH:
MsgBox "A Database Related Error has occured when updating Row " & strProjectKey & " : " & StrResccode & ".Proceeding to next row.", vbInformation, App.Title
Resume Next
End Function
Public Function InsertNewResourceBudget(StrYear As String, strProjectKey As String, StrResccode As String, StrSpread As String, StrJobcharge As String, StrCostcode As String, StrType As String, StrUom As String, StrCurr As String, StrUnitrate As String, StrXchg As String, StrQty As String, Strdays As String, StrDownTime As String, StrWorkComp As String, StrEscl As String, strBCWPAmt As String, StrNotes As String)
'On Error GoTo ErrH
Set cmd = New ADODB.Command
If Cn.State Then Cn.Close
Cn.Open
cmd.ActiveConnection = Cn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "spImportResourcesforBudget"
cmd.parameters("@bd_year").Value = StrYear
cmd.parameters("@bd_projectkey").Value = strProjectKey
cmd.parameters("@bd_resccode").Value = StrResccode
cmd.parameters("@bd_spread").Value = StrSpread
cmd.parameters("@bd_jobcharge").Value = StrJobcharge
cmd.parameters("@bd_costcode").Value = StrCostcode
cmd.parameters("@bd_type").Value = StrType
cmd.parameters("@bd_uom").Value = StrUom
cmd.parameters("@bd_curr").Value = StrCurr
cmd.parameters("@bd_unitrate").Value = CDbl(StrUnitrate)
cmd.parameters("@bd_xchg").Value = StrXchg
cmd.parameters("@bd_qty").Value = CDbl(StrQty)
cmd.parameters("@bd_days").Value = CDbl(Strdays)
cmd.parameters("@bd_tqty").Value = CDbl(StrQty) * CDbl(Strdays)
cmd.parameters("@bd_extdamt").Value = CDbl(StrUnitrate) * CDbl(StrXchg) * (CDbl(StrQty) * CDbl(Strdays))
cmd.parameters("@bd_downtime").Value = CDbl(StrDownTime)
cmd.parameters("@bd_wrkcomp").Value = CDbl(StrWorkComp)
cmd.parameters("@bd_escl").Value = CDbl(StrEscl)
cmd.parameters("@bd_bcwpamt").Value = CDbl(strBCWPAmt)
cmd.parameters("@bd_notes").Value = StrNotes
cmd.parameters("@t_user").Value = main.Label2.Caption
cmd.parameters("@result").Value = 0
cmd.Execute
intRowsInserted = intRowsInserted + cmd("@result")
Set cmd.ActiveConnection = Nothing
Exit Function
ErrH:
MsgBox "A Database Related Error has occured when updating Row " & strProjectKey & " : " & StrResccode & ".Proceeding to next row.", vbInformation, App.Title
Resume Next
End Function
Public Sub data_checkforBudget()
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
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 12)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 14)) = False Then
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
'    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 13)) = False Then
'    asd = 1
'    MSFlexGrid1.Row = ch
'    MSFlexGrid1.Col = 13
'    MSFlexGrid1.CellBackColor = vbRed
'    Else
'    chf.Open "select * from currencymaster where cur_xchgrate='" & MSFlexGrid1.TextMatrix(ch, 13) & "' ", Cn, 3, 2
'    If Not chf.EOF Then
'    Else
'    asd = 1
'    MSFlexGrid1.Row = ch
'    MSFlexGrid1.Col = 13
'    MSFlexGrid1.CellBackColor = vbRed
'    End If
'    End If
'    chf.Close
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
Public Sub data_checkforEIC()
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
    If chf.State Then chf.Close
    chf.Open "select * from resourcedetails where dresc_code='" & MSFlexGrid1.TextMatrix(ch, 3) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 3
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    If MSFlexGrid1.TextMatrix(ch, 4) <> "NA" Then
    chf.Open "select * from progressdurationdetails where prgs_spread_code='" & MSFlexGrid1.TextMatrix(ch, 4) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    chf.Open "select * from progressdurationdetails where prgs_job_key='" & MSFlexGrid1.TextMatrix(ch, 5) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 5
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    End If
    chf.Open "select * from costcode where cc_code='" & MSFlexGrid1.TextMatrix(ch, 6) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 6
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    If (MSFlexGrid1.TextMatrix(ch, 7)) = "" Or IsDate(MSFlexGrid1.TextMatrix(ch, 7)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If (MSFlexGrid1.TextMatrix(ch, 8)) = "" Or IsDate(MSFlexGrid1.TextMatrix(ch, 8)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 8
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If (MSFlexGrid1.TextMatrix(ch, 9)) = "" Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 9
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    chf.Open "select * from resourcemaster where resc_uom='" & MSFlexGrid1.TextMatrix(ch, 10) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 10
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    chf.Open "select * from currencymaster where cur_currency='" & MSFlexGrid1.TextMatrix(ch, 11) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If chf.State Then chf.Close
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 12)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 12
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 13)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 13
    MSFlexGrid1.CellBackColor = vbRed
    Else
    If chf.State Then chf.Close
    chf.Open "select * from currencymaster where cur_xchgrate='" & MSFlexGrid1.TextMatrix(ch, 13) & "' ", Cn, 3, 2
    If Not chf.EOF Then
    Else
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 13
    MSFlexGrid1.CellBackColor = vbRed
    End If
    End If
    If chf.State Then chf.Close
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
    'Skip to Column 22 - ACWP
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 22)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 22
    MSFlexGrid1.CellBackColor = vbRed
    End If
    If IsNumeric(MSFlexGrid1.TextMatrix(ch, 23)) = False Then
    asd = 1
    MSFlexGrid1.Row = ch
    MSFlexGrid1.Col = 23
    MSFlexGrid1.CellBackColor = vbRed
    End If
    End If
Next ch
MsgBox "Check Completed!", vbInformation, App.Title
End Sub

Private Sub optImportBudget_Click()
asd = 1
End Sub

Private Sub optImportEIC_Click()
asd = 1
End Sub

Private Sub txtEditFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            txtEditFlexGrid.Visible = False
            MSFlexGrid1.SetFocus
        Case vbKeyReturn
            ' Finish editing.
            MSFlexGrid1.SetFocus
        Case vbKeyDown
            ' Move down 1 row.
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            End If
        Case vbKeyUp
            ' Move up 1 row.
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
                MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            End If
    End Select
End Sub
' Do not beep on Return or Escape.
Private Sub txtEditFlexGrid_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Then KeyAscii = 0
End Sub
Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Col > 6 Then GridEdit Asc(" ")
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If MSFlexGrid1.Col > 6 Then GridEdit KeyAscii
End Sub
Private Sub MSFlexGrid1_LeaveCell()
    If txtEditFlexGrid.Visible Then
        MSFlexGrid1.Text = txtEditFlexGrid.Text
        txtEditFlexGrid.Visible = False
    End If
End Sub
Private Sub MSFlexGrid1_GotFocus()
    If txtEditFlexGrid.Visible Then
        MSFlexGrid1.Text = txtEditFlexGrid.Text
        txtEditFlexGrid.Visible = False
    End If
End Sub
Private Sub GridEdit(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    txtEditFlexGrid.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    txtEditFlexGrid.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txtEditFlexGrid.Width = MSFlexGrid1.CellWidth
    txtEditFlexGrid.Height = MSFlexGrid1.CellHeight
    txtEditFlexGrid.Visible = True
    txtEditFlexGrid.SetFocus
    Select Case KeyAscii
        Case 0 To Asc(" ")
            txtEditFlexGrid.Text = MSFlexGrid1.Text
            txtEditFlexGrid.SelStart = Len(txtEditFlexGrid.Text)
        Case Else
            txtEditFlexGrid.Text = Chr$(KeyAscii)
            txtEditFlexGrid.SelStart = 1
            Command3.Enabled = False
    End Select
End Sub
