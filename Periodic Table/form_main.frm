VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodic Table Of Elements"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_hide 
      Caption         =   "Hide"
      Height          =   255
      Left            =   120
      TabIndex        =   159
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label lbl_noble_gases 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "noble gases"
      Height          =   255
      Left            =   4920
      TabIndex        =   158
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_other_metals 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "other metals"
      Height          =   255
      Left            =   4920
      TabIndex        =   157
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl_actinide_series 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "actinide series"
      Height          =   255
      Left            =   3360
      TabIndex        =   156
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_lanthanide_series 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lanthanide series"
      Height          =   255
      Left            =   3360
      TabIndex        =   155
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbl_transition_metals 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "transition metals"
      Height          =   255
      Left            =   4920
      TabIndex        =   154
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl_alkaline_earth_metals 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "alkaline earth metals"
      Height          =   255
      Left            =   3360
      TabIndex        =   153
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl_alkali_metals 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "alkali metals"
      Height          =   255
      Left            =   3360
      TabIndex        =   152
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl_nonmetals 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "nonmetals"
      Height          =   255
      Left            =   4920
      TabIndex        =   151
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbl_18 
      Alignment       =   2  'Center
      Caption         =   "18"
      Height          =   255
      Left            =   10920
      TabIndex        =   150
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl_17 
      Alignment       =   2  'Center
      Caption         =   "17"
      Height          =   255
      Left            =   10320
      TabIndex        =   149
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_16 
      Alignment       =   2  'Center
      Caption         =   "16"
      Height          =   255
      Left            =   9720
      TabIndex        =   148
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_15 
      Alignment       =   2  'Center
      Caption         =   "15"
      Height          =   255
      Left            =   9120
      TabIndex        =   147
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_14 
      Alignment       =   2  'Center
      Caption         =   "14"
      Height          =   255
      Left            =   8520
      TabIndex        =   146
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_13 
      Alignment       =   2  'Center
      Caption         =   "13"
      Height          =   255
      Left            =   7920
      TabIndex        =   145
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_12 
      Alignment       =   2  'Center
      Caption         =   "12"
      Height          =   255
      Left            =   7320
      TabIndex        =   144
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_11 
      Alignment       =   2  'Center
      Caption         =   "11"
      Height          =   255
      Left            =   6720
      TabIndex        =   143
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_10 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   255
      Left            =   6120
      TabIndex        =   142
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_9 
      Alignment       =   2  'Center
      Caption         =   "9"
      Height          =   255
      Left            =   5520
      TabIndex        =   141
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_8 
      Alignment       =   2  'Center
      Caption         =   "8"
      Height          =   255
      Left            =   4920
      TabIndex        =   140
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_7 
      Alignment       =   2  'Center
      Caption         =   "7"
      Height          =   255
      Left            =   4320
      TabIndex        =   139
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_6 
      Alignment       =   2  'Center
      Caption         =   "6"
      Height          =   255
      Left            =   3720
      TabIndex        =   138
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_5 
      Alignment       =   2  'Center
      Caption         =   "5"
      Height          =   255
      Left            =   3120
      TabIndex        =   137
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_4 
      Alignment       =   2  'Center
      Caption         =   "4"
      Height          =   255
      Left            =   2520
      TabIndex        =   136
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_3 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Left            =   1920
      TabIndex        =   135
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl_2 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   255
      Left            =   840
      TabIndex        =   134
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl_1 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   133
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbl_0 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   10920
      TabIndex        =   132
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbl_VIIa 
      Alignment       =   2  'Center
      Caption         =   "VIIa"
      Height          =   255
      Left            =   10320
      TabIndex        =   131
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_VIa 
      Alignment       =   2  'Center
      Caption         =   "VIa"
      Height          =   255
      Left            =   9720
      TabIndex        =   130
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_Va 
      Alignment       =   2  'Center
      Caption         =   "Va"
      Height          =   255
      Left            =   9120
      TabIndex        =   129
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IVa 
      Alignment       =   2  'Center
      Caption         =   "IVa"
      Height          =   255
      Left            =   8520
      TabIndex        =   128
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IIIa 
      Alignment       =   2  'Center
      Caption         =   "IIIa"
      Height          =   255
      Left            =   7920
      TabIndex        =   127
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_IIb 
      Alignment       =   2  'Center
      Caption         =   "IIb"
      Height          =   255
      Left            =   7320
      TabIndex        =   126
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_Ib 
      Alignment       =   2  'Center
      Caption         =   "Ib"
      Height          =   255
      Left            =   6720
      TabIndex        =   125
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_right 
      Alignment       =   2  'Center
      Caption         =   "------------|"
      Height          =   255
      Left            =   6120
      TabIndex        =   124
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIIIb 
      Alignment       =   2  'Center
      Caption         =   "VIIIb"
      Height          =   255
      Left            =   5520
      TabIndex        =   123
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_left 
      Alignment       =   2  'Center
      Caption         =   "|-------------"
      Height          =   255
      Left            =   4920
      TabIndex        =   122
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIIb 
      Alignment       =   2  'Center
      Caption         =   "VIIb"
      Height          =   255
      Left            =   4320
      TabIndex        =   121
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_VIb 
      Alignment       =   2  'Center
      Caption         =   "VIb"
      Height          =   255
      Left            =   3720
      TabIndex        =   120
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_Vb 
      Alignment       =   2  'Center
      Caption         =   "Vb"
      Height          =   255
      Left            =   3120
      TabIndex        =   119
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IVb 
      Alignment       =   2  'Center
      Caption         =   "IVb"
      Height          =   255
      Left            =   2520
      TabIndex        =   118
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IIIb 
      Alignment       =   2  'Center
      Caption         =   "IIIb"
      Height          =   255
      Left            =   1920
      TabIndex        =   117
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lbl_IIa 
      Alignment       =   2  'Center
      Caption         =   "IIa"
      Height          =   255
      Left            =   840
      TabIndex        =   116
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lbl_Ia 
      Alignment       =   2  'Center
      Caption         =   "Ia"
      Height          =   255
      Left            =   240
      TabIndex        =   115
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   81
      Left            =   10920
      TabIndex        =   114
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   10920
      TabIndex        =   113
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Xe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   79
      Left            =   10920
      TabIndex        =   112
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   78
      Left            =   10920
      TabIndex        =   111
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   10920
      TabIndex        =   110
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ne"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   76
      Left            =   10920
      TabIndex        =   109
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "He"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   75
      Left            =   10920
      TabIndex        =   108
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl_At 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "At"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   107
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl_I 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   106
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Te"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   71
      Left            =   9720
      TabIndex        =   105
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbl_Br 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Br"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   104
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Se"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   70
      Left            =   9720
      TabIndex        =   103
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "As"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   65
      Left            =   9120
      TabIndex        =   102
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl_Cl 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   101
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   69
      Left            =   9720
      TabIndex        =   100
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   64
      Left            =   9120
      TabIndex        =   99
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Si"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   59
      Left            =   8520
      TabIndex        =   98
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl_F 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   97
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   68
      Left            =   9720
      TabIndex        =   96
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   63
      Left            =   9120
      TabIndex        =   95
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   58
      Left            =   8520
      TabIndex        =   94
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   7920
      TabIndex        =   93
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   73
      Left            =   9720
      TabIndex        =   92
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   8640
      TabIndex        =   91
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Po"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   72
      Left            =   9720
      TabIndex        =   90
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   67
      Left            =   9120
      TabIndex        =   89
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   62
      Left            =   8520
      TabIndex        =   88
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   57
      Left            =   7920
      TabIndex        =   87
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   66
      Left            =   9120
      TabIndex        =   86
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   61
      Left            =   8520
      TabIndex        =   85
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   56
      Left            =   7920
      TabIndex        =   84
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   60
      Left            =   8520
      TabIndex        =   83
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   7920
      TabIndex        =   82
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Al"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   54
      Left            =   7920
      TabIndex        =   81
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   7320
      TabIndex        =   80
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uuu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   6720
      TabIndex        =   79
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Uun"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   6120
      TabIndex        =   78
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   5520
      TabIndex        =   77
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   4920
      TabIndex        =   76
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   4320
      TabIndex        =   75
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   3720
      TabIndex        =   74
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Db"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   3120
      TabIndex        =   73
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   2520
      TabIndex        =   72
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1920
      TabIndex        =   71
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   7320
      TabIndex        =   70
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Au"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   6720
      TabIndex        =   69
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   6120
      TabIndex        =   68
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   5520
      TabIndex        =   67
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Os"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   4920
      TabIndex        =   66
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Re"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   4320
      TabIndex        =   65
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   3720
      TabIndex        =   64
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   3120
      TabIndex        =   63
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2520
      TabIndex        =   62
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   1920
      TabIndex        =   61
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   7320
      TabIndex        =   60
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   6720
      TabIndex        =   59
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   42
      Left            =   6120
      TabIndex        =   58
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   5520
      TabIndex        =   57
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ru"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   4920
      TabIndex        =   56
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   4320
      TabIndex        =   55
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   3720
      TabIndex        =   54
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   3120
      TabIndex        =   53
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2520
      TabIndex        =   52
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   1920
      TabIndex        =   51
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   7320
      TabIndex        =   50
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   6720
      TabIndex        =   49
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ni"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   6120
      TabIndex        =   48
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Co"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   5520
      TabIndex        =   47
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   4920
      TabIndex        =   46
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   4320
      TabIndex        =   45
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   3720
      TabIndex        =   44
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3120
      TabIndex        =   43
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ti"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   2520
      TabIndex        =   42
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1920
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.Line Line6 
      X1              =   1560
      X2              =   2520
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      X1              =   1560
      X2              =   1560
      Y1              =   3840
      Y2              =   5040
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   1560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   2520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   1800
      Y1              =   3360
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1800
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   108
      Left            =   10320
      TabIndex        =   40
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Md"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   107
      Left            =   9720
      TabIndex        =   39
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   105
      Left            =   9120
      TabIndex        =   38
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Es"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   103
      Left            =   8520
      TabIndex        =   37
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   101
      Left            =   7920
      TabIndex        =   36
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   99
      Left            =   7320
      TabIndex        =   35
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   97
      Left            =   6720
      TabIndex        =   34
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   95
      Left            =   6120
      TabIndex        =   33
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   93
      Left            =   5520
      TabIndex        =   32
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Np"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   91
      Left            =   4920
      TabIndex        =   31
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   89
      Left            =   4320
      TabIndex        =   30
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   87
      Left            =   3720
      TabIndex        =   29
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Th"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   85
      Left            =   3120
      TabIndex        =   28
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ac"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   83
      Left            =   2520
      TabIndex        =   27
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   109
      Left            =   10320
      TabIndex        =   26
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   106
      Left            =   9720
      TabIndex        =   25
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Er"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   104
      Left            =   9120
      TabIndex        =   24
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   102
      Left            =   8520
      TabIndex        =   23
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   100
      Left            =   7920
      TabIndex        =   22
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   98
      Left            =   7320
      TabIndex        =   21
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   96
      Left            =   6720
      TabIndex        =   20
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Eu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   94
      Left            =   6120
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   92
      Left            =   5520
      TabIndex        =   18
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   90
      Left            =   4920
      TabIndex        =   17
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   88
      Left            =   4320
      TabIndex        =   16
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   86
      Left            =   3720
      TabIndex        =   15
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ce"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   84
      Left            =   3120
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "La"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   82
      Left            =   2520
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   840
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   840
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   840
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   840
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Be"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   840
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Na"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Li"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSym 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conPT As ADODB.Connection
Public rstPT As ADODB.Recordset
Private Sub Form_Load()
Set conPT = New ADODB.Connection
With conPT
 .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ptoe.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
 .Open
End With
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblSym_Click(Index As Integer)
    Set rstPT = New ADODB.Recordset
    With rstPT
            .Open "select * from data_list where  symbol = '" & lblSym(Index).Caption & "'", conPT
        While .EOF = False
            frm_show.txt_atom_num.Text = !atom_num
            frm_show.txt_name.Text = !Name
            frm_show.txt_atom_mass.Text = !atom_mass
            frm_show.txt_atom_rad.Text = !atom_rad
            frm_show.txt_orbitals.Text = !orbitals
            frm_show.txt_elec_shell.Text = !elec_shell
            frm_show.txt_melt_boil.Text = !melt_boil
            frm_show.txt_density.Text = !density
            frm_show.txt_isotopes.Text = !isotopes
            frm_show.txt_electronegativity = !electronegativity
            frm_show.txt_oxidation.Text = !oxidation
            frm_show.txt_discovery.Text = !discovery
            .MoveNext
        
        Wend
    End With
frm_show.Show 1
End Sub
