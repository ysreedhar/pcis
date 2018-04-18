VERSION 5.00
Begin VB.Form frm_resourcemapping 
   BackColor       =   &H00FFFFFF&
   Caption         =   "RESOURCE MAP TO PROJECTKEY"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   14955
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4560
         TabIndex        =   5
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton cmd_map 
            BackColor       =   &H00DC7E5A&
            Caption         =   "<<<<<        MAP      >>> >>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            ToolTipText     =   "Click to Map Resouce to Project"
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8190
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   1320
         Width           =   7455
      End
      Begin VB.ComboBox cbo_proj 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resource Code  - Description"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key - Description"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frm_resourcemapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_proj_Click()
 
Dim t As Integer
For t = 0 To List1.ListCount - 1
List1.Selected(t) = False
Next t
mm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
Dim rr As New ADODB.Recordset
If rr.State Then rr.Close
rr.Open "select * from resourcedetails where dresc_proj='" & mm(0) & "' ", Cn, 3, 2
While Not rr.EOF
For k = 0 To List1.ListCount - 1
nn = Split(List1.List(k), "  -  ", Len(List1.List(k)), vbTextCompare)

If nn(0) = rr!dresc_code Then
List1.Selected(k) = True
End If

Next k
rr.MoveNext
Wend
End Sub

Private Sub cmd_close_Click()
'Unload Me
End Sub

Private Sub cmd_map_Click()
hg = Year(Date)
nm = Split(cbo_proj.Text, "  -  ", Len(cbo_proj.Text), vbTextCompare)
pp = Split(cbo_proj.Text, "20", Len(cbo_proj.Text), vbTextCompare)
Dim hgg As String
hgg = Mid(pp(1), 1, 2)
'''Dim dl As New ADODB.Recordset
'''    If dl.State Then dl.Close
    Cn.Execute "delete from resourcedetails where dresc_proj='" & nm(0) & "'"
Dim l As Integer
l = 0
For l = 0 To List1.ListCount - 1
If List1.Selected(l) = True Then
nn = Split(List1.List(l), "  -  ", Len(List1.List(l)), vbTextCompare)
                        If l = 0 Then
                        Cn.Execute "delete from resourcedetails where dresc_proj='" & nm(0) & "' and  dresc_code <> '" & nn(0) & "'"
                        End If
                Dim sv As New ADODB.Recordset
                If sv.State Then sv.Close
                sv.Open "select * from resourcedetails where dresc_code='" & nn(0) & "' and dresc_proj='" & nm(0) & "'", Cn, 3, 2
                If sv.EOF Then
                sv.AddNew
                sv!dresc_proj = nm(0)
                sv!dresc_code = nn(0)
                sv!dresc_year = "20" & hgg
                sv!dresc_curcy = "RM"
                sv!dresc_rate = 0
                sv!dresc_ratetype = "BR"
                sv!dresc_notes = "-"
                Dim rd As New ADODB.Recordset
                If rd.State Then rd.Close
                rd.Open "select * from resourcemaster where resc_code='" & nn(0) & "' ", Cn, 3, 2
                If Not rd.EOF Then
                sv!resc_id = rd!resc_id
                End If
                sv!t_date = Format(Date, "dd/MM/yyyy")
                sv!u_date = Now
                sv!t_user = main.Label2.Caption
                sv.Update
                sv.Close
                End If
                
                

End If
Next l
MsgBox "Mapped"
Dim i As Integer
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "RESOURCE MAP TO PROJECTKEY"
Me.Top = 10
Me.Left = 10

Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(p.proj_key),p.proj_desc from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2 & "' order by p.proj_key", Cn, 3, 2
While Not rs.EOF
cbo_proj.AddItem rs(0) & "  -  " & rs(1)
rs.MoveNext
Wend
rs.Close
Dim lst As New ADODB.Recordset
If lst.State Then lst.Close
lst.Open "select DISTINCT(resc_code),resc_desc from resourcemaster order by resc_code", Cn, 3, 2
While Not lst.EOF
List1.AddItem lst(0) & "  -  " & lst(1)
lst.MoveNext
Wend
lst.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
