VERSION 5.00
Begin VB.Form frm_replaceresc 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Replace Resource"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DC7E5A&
         Height          =   615
         Left            =   9120
         Picture         =   "frm_replaceresc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.ComboBox cbo_projcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Key - Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Replace Resource"
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11175
      Begin VB.CommandButton a0 
         Enabled         =   0   'False
         Height          =   255
         Index           =   19
         Left            =   9840
         TabIndex        =   128
         Top             =   7920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   19
         Left            =   840
         TabIndex        =   125
         Top             =   7920
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   19
         Left            =   5880
         TabIndex        =   124
         Top             =   7920
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   19
         Left            =   10560
         TabIndex        =   123
         Top             =   7920
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   18
         Left            =   840
         TabIndex        =   120
         Top             =   7560
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   18
         Left            =   5880
         TabIndex        =   119
         Top             =   7560
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   18
         Left            =   10560
         TabIndex        =   118
         Top             =   7560
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   18
         Left            =   9840
         TabIndex        =   117
         Top             =   7560
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   17
         Left            =   840
         TabIndex        =   114
         Top             =   7200
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   17
         Left            =   5880
         TabIndex        =   113
         Top             =   7200
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   17
         Left            =   10560
         TabIndex        =   112
         Top             =   7200
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   17
         Left            =   9840
         TabIndex        =   111
         Top             =   7200
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   16
         Left            =   840
         TabIndex        =   108
         Top             =   6840
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   16
         Left            =   5880
         TabIndex        =   107
         Top             =   6840
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   16
         Left            =   10560
         TabIndex        =   106
         Top             =   6840
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   16
         Left            =   9840
         TabIndex        =   105
         Top             =   6840
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   15
         Left            =   840
         TabIndex        =   102
         Top             =   6480
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   15
         Left            =   5880
         TabIndex        =   101
         Top             =   6480
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   15
         Left            =   10560
         TabIndex        =   100
         Top             =   6480
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   15
         Left            =   9840
         TabIndex        =   99
         Top             =   6480
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   14
         Left            =   840
         TabIndex        =   96
         Top             =   6120
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   14
         Left            =   5880
         TabIndex        =   95
         Top             =   6120
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   14
         Left            =   10560
         TabIndex        =   94
         Top             =   6120
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   14
         Left            =   9840
         TabIndex        =   93
         Top             =   6120
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   13
         Left            =   840
         TabIndex        =   90
         Top             =   5760
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   13
         Left            =   5880
         TabIndex        =   89
         Top             =   5760
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   13
         Left            =   10560
         TabIndex        =   88
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   13
         Left            =   9840
         TabIndex        =   87
         Top             =   5760
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   12
         Left            =   840
         TabIndex        =   84
         Top             =   5400
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   12
         Left            =   5880
         TabIndex        =   83
         Top             =   5400
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   12
         Left            =   10560
         TabIndex        =   82
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   12
         Left            =   9840
         TabIndex        =   81
         Top             =   5400
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   11
         Left            =   840
         TabIndex        =   78
         Top             =   5040
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   11
         Left            =   5880
         TabIndex        =   77
         Top             =   5040
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   11
         Left            =   10560
         TabIndex        =   76
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   11
         Left            =   9840
         TabIndex        =   75
         Top             =   5040
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   10
         Left            =   840
         TabIndex        =   72
         Top             =   4680
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   10
         Left            =   5880
         TabIndex        =   71
         Top             =   4680
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   10
         Left            =   10560
         TabIndex        =   70
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   10
         Left            =   9840
         TabIndex        =   69
         Top             =   4680
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   9
         Left            =   840
         TabIndex        =   66
         Top             =   4320
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   9
         Left            =   5880
         TabIndex        =   65
         Top             =   4320
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   9
         Left            =   10560
         TabIndex        =   64
         Top             =   4320
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   9
         Left            =   9840
         TabIndex        =   63
         Top             =   4320
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   8
         Left            =   840
         TabIndex        =   60
         Top             =   3960
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   8
         Left            =   5880
         TabIndex        =   59
         Top             =   3960
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   8
         Left            =   10560
         TabIndex        =   58
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   8
         Left            =   9840
         TabIndex        =   57
         Top             =   3960
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   7
         Left            =   840
         TabIndex        =   54
         Top             =   3600
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   7
         Left            =   5880
         TabIndex        =   53
         Top             =   3600
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   7
         Left            =   10560
         TabIndex        =   52
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   7
         Left            =   9840
         TabIndex        =   51
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   6
         Left            =   840
         TabIndex        =   48
         Top             =   3240
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   6
         Left            =   5880
         TabIndex        =   47
         Top             =   3240
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   6
         Left            =   10560
         TabIndex        =   46
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   6
         Left            =   9840
         TabIndex        =   45
         Top             =   3240
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   5
         Left            =   840
         TabIndex        =   42
         Top             =   2880
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   5
         Left            =   5880
         TabIndex        =   41
         Top             =   2880
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   5
         Left            =   10560
         TabIndex        =   40
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   5
         Left            =   9840
         TabIndex        =   39
         Top             =   2880
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   4
         Left            =   840
         TabIndex        =   36
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   4
         Left            =   5880
         TabIndex        =   35
         Top             =   2520
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   4
         Left            =   10560
         TabIndex        =   34
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   4
         Left            =   9840
         TabIndex        =   33
         Top             =   2520
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   3
         Left            =   840
         TabIndex        =   30
         Top             =   2160
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   29
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   3
         Left            =   10560
         TabIndex        =   28
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   3
         Left            =   9840
         TabIndex        =   27
         Top             =   2160
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   24
         Top             =   1800
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   2
         Left            =   5880
         TabIndex        =   23
         Top             =   1800
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   2
         Left            =   10560
         TabIndex        =   22
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   2
         Left            =   9840
         TabIndex        =   21
         Top             =   1800
         Width           =   495
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   18
         Top             =   1440
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   1
         Left            =   5880
         TabIndex        =   17
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   1
         Left            =   10560
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton a0 
         Caption         =   "Next"
         Height          =   255
         Index           =   0
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton c0 
         Caption         =   "Clear"
         Height          =   255
         Index           =   0
         Left            =   10560
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox txt_job2 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   4575
      End
      Begin VB.ComboBox TXT_RESOURCEN 
         Height          =   315
         Index           =   0
         Left            =   5880
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.ComboBox TXT_RESOURCEO 
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   3855
      End
      Begin VB.ComboBox txt_spreado1 
         Height          =   315
         Left            =   5040
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   19
         Left            =   5040
         TabIndex        =   127
         Top             =   7920
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   126
         Top             =   7920
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   18
         Left            =   5040
         TabIndex        =   122
         Top             =   7560
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   121
         Top             =   7560
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   17
         Left            =   5040
         TabIndex        =   116
         Top             =   7200
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   115
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   16
         Left            =   5040
         TabIndex        =   110
         Top             =   6840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   109
         Top             =   6840
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   15
         Left            =   5040
         TabIndex        =   104
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   103
         Top             =   6480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   14
         Left            =   5040
         TabIndex        =   98
         Top             =   6120
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   97
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   13
         Left            =   5040
         TabIndex        =   92
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   91
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   12
         Left            =   5040
         TabIndex        =   86
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   85
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   80
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   79
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   74
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   73
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   68
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   67
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   62
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   61
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   56
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   55
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   50
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   44
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   43
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   38
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "JobCharge"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Spread"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(Old)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Resc(New)"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_replaceresc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a0_Click(Index As Integer)
TXT_RESOURCEO(Index + 1).Visible = True
TXT_RESOURCEN(Index + 1).Visible = True
Label2(Index + 1).Visible = True
Label3(Index + 1).Visible = True
a0(Index + 1).Visible = True
c0(Index + 1).Visible = True
End Sub

Private Sub c0_Click(Index As Integer)

TXT_RESOURCEO(Index).Text = ""
TXT_RESOURCEN(Index).Text = ""
If Index <> 0 Then
TXT_RESOURCEO(Index).Visible = False
TXT_RESOURCEN(Index).Visible = False
Label2(Index).Visible = False
Label3(Index).Visible = False
a0(Index).Visible = False
c0(Index).Visible = False
End If
End Sub

Private Sub cbo_projcode_Click()
nm = Split(cbo_projcode.Text, "  -  ", Len(cbo_projcode.Text), vbTextCompare)
 Dim rc As New ADODB.Recordset
 
            


            
            
            
            Dim rc1 As New ADODB.Recordset
            If rc1.State Then rc1.Close
            rc1.Open "select DISTINCT(j.job_code),j.job_desc from cost c, jobcharge j where c.bd_jobcharge=j.job_code and j.job_proj_key = '" & nm(0) & "'  and bd_spread <> 'NA'  order by j.job_code", Cn, 3, 2
            While Not rc1.EOF
          
            txt_job2.AddItem rc1(0) & "  -  " & rc1(1)
            rc1.MoveNext
            Wend
            rc1.Close
End Sub



Private Sub Command2_Click()
If cbo_projcode.Text = "" Then
MsgBox "Select Project"
Exit Sub
End If
If txt_job2.Text = "" Then
MsgBox "Select JobCharge"
Exit Sub
End If
If txt_spreado1.Text = "" Then
MsgBox "Select Spread"
Exit Sub
End If
Dim k As Integer
k = 0
nn7 = Split(txt_job2.Text, "  -  ", Len(txt_job2.Text), vbTextCompare)
nn5 = Split(txt_spreado1.Text, "  -  ", Len(txt_spreado1.Text), vbTextCompare)
For k = 0 To 19
If TXT_RESOURCEO(k) <> "" Then
nn4 = Split(TXT_RESOURCEN(k).Text, "  -  ", Len(TXT_RESOURCEN(k).Text), vbTextCompare)
nn6 = Split(TXT_RESOURCEO(k).Text, "  -  ", Len(TXT_RESOURCEO(k).Text), vbTextCompare)

Cn.Execute "UPDATE COST SET BD_RESCCODE='" & nn4(0) & "' , bd_type='A' WHERE BD_SPREAD='" & nn5(0) & "' AND BD_RESCCODE='" & nn6(0) & "' AND BD_COSTTYPE='E'   and bd_jobcharge='" & nn7(0) & "'"
End If
Next k

MsgBox "DONE"
End Sub




Private Sub Command7_Click()
'Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 5
Me.Left = 5


Dim rv As New ADODB.Recordset
If rv.State Then rv.Close
rv.Open "select DISTINCT(p.proj_key),p.proj_title from projectmaster p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key", Cn, 3, 2
While Not rv.EOF
cbo_projcode.AddItem rv(0) & "  -  " & rv(1)
rv.MoveNext
Wend
Dim h As Integer
h = 0
For h = 1 To 19
TXT_RESOURCEO(h).Visible = False
TXT_RESOURCEN(h).Visible = False
Label2(h).Visible = False
Label3(h).Visible = False
a0(h).Visible = False
c0(h).Visible = False
Next h
End Sub

Private Sub txt_job2_Click()
nnw1 = Split(txt_job2.Text, "  -  ", Len(txt_job2.Text), vbTextCompare)
'txt_spreado1.Clear
Dim spr As New ADODB.Recordset
If spr.State Then spr.Close
spr.Open "select DISTINCT(s.spread_code),s.spread_desc from spreadmaster s , cost c where s.spread_code=c.bd_spread and c.bd_jobcharge='" & nnw1(0) & "' and s.spread_code <>'NA' order by s.spread_code", Cn, 3, 2
While Not spr.EOF
txt_spreado1.AddItem spr(0) & "  -  " & spr(1)
spr.MoveNext
Wend
End Sub

Private Sub txt_spreado1_Click()
nnw3 = Split(txt_spreado1.Text, "  -  ", Len(txt_spreado1.Text), vbTextCompare)
Dim i As Integer
Dim j As Integer
i = 0: j = 0

''    For i = 0 To 19
''    TXT_RESOURCEO(i).Clear
''    Next i
''
''    For j = 0 To 19
''    TXT_RESOURCEN(j).Clear
''    Next j

i = 0: j = 0
    Dim rcd As New ADODB.Recordset
    If rcd.State Then rcd.Close
    rcd.Open "select DISTINCT(r.resc_code),r.resc_desc from resourcemaster r , cost c where c.bd_costtype='E' and r.resc_code=c.bd_resccode and bd_spread='" & nnw3(0) & "' order by r.resc_code", Cn, 3, 2
    While Not rcd.EOF
    For i = 0 To 19
    TXT_RESOURCEO(i).AddItem rcd(0) & "  -  " & rcd(1)
    Next i
    rcd.MoveNext
    Wend
    
    
     Dim rcd1 As New ADODB.Recordset
    If rcd1.State Then rcd1.Close
    rcd1.Open "select DISTINCT(resc_code),resc_desc from resourcemaster   order by resc_code", Cn, 3, 2
    While Not rcd1.EOF
    For j = 0 To 19
    TXT_RESOURCEN(j).AddItem rcd1(0) & "  -  " & rcd1(1)
    Next j
    rcd1.MoveNext
    Wend
End Sub
