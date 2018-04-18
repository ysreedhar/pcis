VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_budgettransferfull 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_year 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox cbo_pproj 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   635
      ButtonWidth     =   6906
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Generate EIC Transactions from BC Transactions"
            Key             =   "ar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComCtl2.DTPicker dtpdefault 
         Height          =   375
         Left            =   9000
         TabIndex        =   5
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy H:mm:ss"
         Format          =   64684035
         CurrentDate     =   38140
      End
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frm_budgettransferfull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''::::::::::::::::::::::::::::::::::::::::::
'Dim btra As New ADODB.Recordset
'If btra.State Then btra.Close
'btra.Open "select * from cost where bd_year='" & cbo_year.Text & "' and bd_projeckkey='" & cbo_pproj.Text & "'", Cn, 3, 2
'While Not btra.EOF
'        btra!bd_year = cbo_year.Text
'        btra!bd_resccode = ab(0)
'        btra!bd_rescname = ab(1)
'        btra!bd_brate = Format(fl!dresc_rate, "###,###,##0.00")
'        btra!bd_vendor = fl!resc_vendorcode
'        btra!bd_respcode = fl!resc_respcode
'        btra!bd_respname = rr(0)
'        btra!bd_crate = 0
'        btra!bd_projectkey = nh(0)
'        btra!bd_projectdesc = nh(1)
'        btra!bd_costtype = "E"
'        btra!bd_cuttdate = main.DTPcutdate1.Value
'        btra!bd_spread = ac(0)
'        btra!bd_tranx = flex_grid.TextMatrix(i, 2)
'        btra!bd_jobcharge = ng(0)
'        btra!bd_costcode = ad(0)
'        btra!bd_qty = flex_grid.TextMatrix(i, 7)
'        btra!bd_days = flex_grid.TextMatrix(i, 8)
'        btra!bd_tqty = flex_grid.TextMatrix(i, 9)
'        btra!bd_uom = flex_grid.TextMatrix(i, 10)
'        btra!bd_curr = flex_grid.TextMatrix(i, 11)
'        btra!bd_unitrate = (((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))))
'        btra!bd_xchg = flex_grid.TextMatrix(i, 12)
'        btra!bd_downtime = flex_grid.TextMatrix(i, 14)
'        btra!bd_escl = flex_grid.TextMatrix(i, 15)
'        btra!bd_extdamt = CDbl(flex_grid.TextMatrix(i, 12)) * CDbl((((CDbl(100 + CDbl(flex_grid.TextMatrix(i, 15))) / 100) * CDbl(flex_grid.TextMatrix(i, 13))))) * CDbl(flex_grid.TextMatrix(i, 9))
'        btra!bd_wrkcomp = flex_grid.TextMatrix(i, 17)
'        btra!bd_bcwpamt = flex_grid.TextMatrix(i, 18)
'        btra!bd_e_days = 0
'        btra!bd_e_tqty = 0
'        btra!bd_e_extdamt = 0
'        Dim pl As Double
'        pl = 0
'        pl = CDbl(btra!bd_days) + CDbl(btra!bd_e_days)
'        btra!bd_chk = 1
'        If btra!bd_spread = "NA" Then
'        btra!bd_sdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
'        If pl = 0 Then
'        btra!bd_edate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
'        Else
'        btra!bd_edate = Format(DateAdd("d", CDbl(pl), dtpdefault.Value), "dd/MM/yyyy H:mm:ss")
'        End If
'        Else
'        btra!bd_sdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
'        btra!bd_edate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
'        End If
'        btra!bd_inv = "-"
'        btra!bd_invdate = Format(dtpdefault.Value, "dd/MM/yyyy H:mm:ss")
'        btra!bd_type = "-"
'        btra!bd_notes = "-"
'        btra!t_date = flex_grid.TextMatrix(i, 20)
'        btra!u_date = Now
'        btra!t_user = main.Label2.Caption
'        btra!bd_obs = flex_grid.TextMatrix(i, 5)
'        btra!bd_idd = flex_grid.TextMatrix(i, 0)
'btra.MoveNext
'Wend
'':::::::::::::::::::::::::::::::::::::::::
'End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
