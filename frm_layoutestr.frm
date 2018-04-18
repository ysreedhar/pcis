VERSION 5.00
Begin VB.Form frm_layoutestr 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DC7E5A&
      Height          =   615
      Left            =   3120
      Picture         =   "frm_layoutestr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frm_layoutestr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Select All" Then

Dim i As Integer
i = 0
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next i
Command1.Caption = "DeSelect"
ElseIf Command1.Caption = "DeSelect" Then

Dim j As Integer
j = 0
For j = 0 To List1.ListCount - 1
List1.Selected(j) = False
Next j
Command1.Caption = "Select All"
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 5
Me.Left = 5
Me.Width = 4320
Me.Height = 6690
        List1.AddItem "Type"
        
        List1.AddItem "Spread "
       
        List1.AddItem "JobCharge"
      
        List1.AddItem "OBS"
      
        List1.AddItem "CostCode"
      
        List1.AddItem "Start Date"
      
        List1.AddItem "End Date"
      
        List1.AddItem "Qty"
 
        List1.AddItem "A Days"
        
        List1.AddItem "Tot Qty"
     
        List1.AddItem "UOM "
    
        List1.AddItem "Curcy "
   
        List1.AddItem "UnitRate"
       
        List1.AddItem "Xrate"
       
        List1.AddItem "ACWP Amount"
      
        List1.AddItem "E Days"
 
        List1.AddItem "Tot Qty"
     
        List1.AddItem "ECTC Amount"
 
        List1.AddItem "Notes"
   
        
End Sub

Private Sub List1_Click()
If List1.SelCount >= 1 Then
Command1.Caption = "Select All"
End If
End Sub
