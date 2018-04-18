VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flex_jobcharge 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   16325
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call flex_title
End Sub
Public Sub flex_title()


    With flex_jobcharge
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Job Charge"
        .ColWidth(1) = 2500
        .TextMatrix(0, 2) = "Description"
        .ColWidth(2) = 4960
        
        
    End With
End Sub
