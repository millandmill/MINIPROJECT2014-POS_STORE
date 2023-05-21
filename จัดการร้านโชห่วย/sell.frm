VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ระบบขายสินค้า"
   ClientHeight    =   11895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17535
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11895
   ScaleWidth      =   17535
   Begin VB.CommandButton Print_pa 
      Caption         =   "Print ใบเสร็จ"
      BeginProperty Font 
         Name            =   "AngsanaUPC"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7680
      Picture         =   "sell.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "ตารางการทอนเงิน"
      Height          =   2175
      Left            =   120
      TabIndex        =   165
      Top             =   9600
      Width           =   17295
      Begin VB.Label Label124 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14760
         TabIndex        =   187
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label123 
         Caption         =   "0.25 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14760
         TabIndex        =   186
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label122 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   185
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Line Line13 
         X1              =   14520
         X2              =   14520
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Label Label121 
         Caption         =   "0.50 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         TabIndex        =   184
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line12 
         X1              =   12960
         X2              =   12960
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line11 
         X1              =   10560
         X2              =   10560
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line10 
         X1              =   11760
         X2              =   11760
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Label Label120 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11760
         TabIndex        =   183
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Line Line9 
         X1              =   9360
         X2              =   9360
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line8 
         X1              =   8040
         X2              =   8040
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line7 
         X1              =   6720
         X2              =   6720
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line6 
         X1              =   5400
         X2              =   5400
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line5 
         X1              =   3960
         X2              =   3960
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line4 
         X1              =   2400
         X2              =   2400
         Y1              =   360
         Y2              =   1920
      End
      Begin VB.Line Line3 
         X1              =   840
         X2              =   16200
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label110 
         Caption         =   "1 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   173
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label109 
         Caption         =   "5 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   172
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label108 
         Caption         =   "10 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   171
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label107 
         Caption         =   "20 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   170
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label106 
         Caption         =   "50 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   169
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label104 
         Caption         =   "500 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   167
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label103 
         Caption         =   "1000 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   166
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label112 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   175
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label113 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   176
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label114 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   177
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label115 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   178
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label116 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   179
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label117 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   180
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label118 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   181
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label119 
         Caption         =   "2 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         TabIndex        =   182
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label105 
         Caption         =   "100 บาท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   168
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label111 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   174
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.TextBox mon_in 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      MaxLength       =   13
      TabIndex        =   162
      Text            =   "0.00"
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton s15 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   155
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton s14 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   154
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton s13 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   153
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton s12 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   152
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton s11 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   151
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton s10 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   150
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton s9 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   149
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton s8 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   148
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton s7 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   147
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton s6 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   146
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton s5 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   145
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton s4 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   144
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton s3 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   143
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton s2 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   142
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา15 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   141
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา14 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   140
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา13 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   139
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา12 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   138
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา11 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   137
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา10 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   136
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา9 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   135
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา8 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   134
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา7 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   133
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา6 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   132
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา5 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   131
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา4 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   130
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา3 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   129
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา2 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   128
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton ค้นหา1 
      Caption         =   "check"
      Height          =   255
      Left            =   5400
      TabIndex        =   127
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton s1 
      Caption         =   "ขาย"
      Height          =   255
      Left            =   1080
      TabIndex        =   126
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "เปิดรายการขายใหม่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   125
      Top             =   7680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ปิดรายการขาย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   124
      Top             =   7680
      Width           =   2775
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   101
      Text            =   "0"
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   98
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   94
      Text            =   "0"
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   91
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   87
      Text            =   "0"
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   84
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   80
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   77
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   73
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   70
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   66
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   63
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   59
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   56
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   52
      Text            =   "0"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   49
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   45
      Text            =   "0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   42
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   38
      Text            =   "0"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   35
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   31
      Text            =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   28
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   24
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "0"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   12960
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      MaxLength       =   13
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label125 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   189
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label102 
      Caption         =   "ทอน 0 บาท"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   164
      Top             =   8880
      Width           =   4695
   End
   Begin VB.Label Label98 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   159
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label101 
      Alignment       =   2  'Center
      Caption         =   "บาท"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   163
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label Label100 
      Caption         =   "รับเงินมา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   161
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label99 
      Alignment       =   2  'Center
      Caption         =   "ชิ้น"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   160
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "จำนวนทั้งหมด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   158
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label97 
      Alignment       =   2  'Center
      Caption         =   "รายการ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   157
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label96 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   156
      Top             =   7680
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   15480
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label95 
      Caption         =   "รายการสินค้าที่ขาย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   123
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label94 
      Caption         =   "ราคาทั้งหมด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   122
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label93 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   121
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label92 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   120
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label91 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   119
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label Label90 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   118
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label Label89 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   117
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label88 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   116
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label87 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   115
      Top             =   10320
      Width           =   3375
   End
   Begin VB.Label Label86 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   114
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label85 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   113
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label84 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   112
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label83 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   111
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label82 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   110
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label81 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   109
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label80 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   108
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label79 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   6960
      TabIndex        =   107
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label77 
      Alignment       =   1  'Right Justify
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   9600
      TabIndex        =   105
      Top             =   480
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   9480
      X2              =   14760
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label76 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   104
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label75 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   103
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label74 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   102
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label73 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   100
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label72 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   99
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label71 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   97
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label70 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   96
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label69 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   95
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label68 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   93
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label67 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   92
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label66 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   90
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label65 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   89
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label64 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   88
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label63 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   86
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label62 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   85
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label61 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   83
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label60 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   82
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label59 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   81
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label58 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   79
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label57 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   78
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label56 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   76
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label55 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   75
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label54 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   74
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label53 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   72
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label52 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   71
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label51 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   69
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label50 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   68
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label49 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   67
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label48 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   65
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label47 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   64
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label46 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   62
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label45 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   61
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label44 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   60
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label43 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   58
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label42 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   57
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label41 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   55
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label40 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   54
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label39 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   53
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label38 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   51
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label37 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   50
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label36 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   48
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label35 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   47
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label34 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   46
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label33 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   44
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label32 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   43
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label31 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   41
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label30 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   40
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label29 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   39
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label28 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   37
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label27 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   36
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label26 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   34
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label25 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   33
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label24 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   32
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label23 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   30
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   27
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label20 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   26
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label19 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   25
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   23
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   20
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   19
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   18
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   16
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "รวมราคา 0.00 บาท"
      Height          =   255
      Left            =   14160
      TabIndex        =   6
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "ราคาชิ้นละ 0.00 บาท"
      Height          =   255
      Left            =   10440
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   13680
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label label2 
      Caption         =   "จำนวน"
      Height          =   255
      Left            =   12240
      TabIndex        =   2
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "รหัสbarcode"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label78 
      Caption         =   "฿"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14160
      TabIndex        =   106
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, num, num_s As Integer
Dim p_all As Currency
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
'ปิดรายการขาย
Private Sub Command1_Click()
    Unload Form4
End Sub
'เปิดรายการขายใหม่
Private Sub Command2_Click()
Unload Form4
Form4.Show
End Sub

Private Sub Form_Load()
p_all = 0
num = 0
num_s = 0
mon_in.Enabled = False
p1 = 0
p2 = 0
p3 = 0
p4 = 0
p5 = 0
p6 = 0
p7 = 0
p8 = 0
p9 = 0
p10 = 0
p11 = 0
p12 = 0
p13 = 0
p14 = 0
p15 = 0

'เปิดฐานข้อมูล
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With

'เปิดตาราง
With RC
        SQL = "SELECT * FROM สินค้า"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
End With

Text2.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text8.Enabled = False
Text10.Enabled = False
Text12.Enabled = False
Text14.Enabled = False
Text16.Enabled = False
Text18.Enabled = False
Text20.Enabled = False
Text22.Enabled = False
Text24.Enabled = False
Text26.Enabled = False
Text28.Enabled = False
Text30.Enabled = False

p_all = 0
num = 0
num_s = 0

Label77.Caption = p_all
Label96.Caption = num
Label98.Caption = num_s
End Sub


Private Sub Text1_Change()
  If IsNumeric(Text1.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text1.text = ""
  End If
End Sub

Private Sub Text3_Change()
  If IsNumeric(Text3.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text3.text = ""
  End If
End Sub

Private Sub Text5_Change()
  If IsNumeric(Text5.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text5.text = ""
  End If
End Sub

Private Sub Text7_Change()
  If IsNumeric(Text7.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text7.text = ""
  End If
End Sub

Private Sub Text9_Change()
  If IsNumeric(Text9.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text9.text = ""
  End If
End Sub

Private Sub Text11_Change()
  If IsNumeric(Text11.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text11.text = ""
  End If
End Sub

Private Sub Text13_Change()
  If IsNumeric(Text13.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text13.text = ""
  End If
End Sub

Private Sub Text15_Change()
  If IsNumeric(Text15.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text15.text = ""
  End If
End Sub

Private Sub Text17_Change()
  If IsNumeric(Text17.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text17.text = ""
  End If
End Sub

Private Sub Text19_Change()
  If IsNumeric(Text19.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text19.text = ""
  End If
End Sub

Private Sub Text21_Change()
  If IsNumeric(Text21.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text21.text = ""
  End If
End Sub

Private Sub Text23_Change()
  If IsNumeric(Text23.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text23.text = ""
  End If
End Sub

Private Sub Text25_Change()
  If IsNumeric(Text25.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text25.text = ""
  End If
End Sub

Private Sub Text27_Change()
  If IsNumeric(Text27.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text27.text = ""
  End If
End Sub

Private Sub Text29_Change()
  If IsNumeric(Text29.text) = False Then
    MsgBox ("กรุณาใส่รหัส barcode เป็นตัวเลข")
    Text29.text = ""
  End If
End Sub

'ขายสินค้าที่1
Private Sub s1_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text2.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text2.text & " WHERE barcode = " & Text1.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s1.Enabled = False
        ค้นหา1.Enabled = False
        Text2.Enabled = False
        Text1.Enabled = False
        p_all = p_all + p1
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text2.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label79, 12) & "' , '" & Mid(Label4, 12) & "'" & ",'" & Text2.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text2.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub
'จำนวนชิ้น1
Private Sub Text2_Change()
With RC
'        If (Text4.Text >= 0) And (Text4.Text <> "") Then
        On Error GoTo e1
        p1 = .Fields("ราคา") * Text2.text
        Label7.Caption = "รวมราคา " & p1 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา1
Private Sub ค้นหา1_Click()
With RC
If (Text1.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text1.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label79.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label4.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text2.Enabled = True
            Text1.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text1.text = ""
End If
End Sub







'ขายสินค้าที่2
Private Sub s2_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text4.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text4.text & " WHERE barcode = " & Text3.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s2.Enabled = False
        ค้นหา2.Enabled = False
        Text4.Enabled = False
        Text3.Enabled = False
        p_all = p_all + p2
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text4.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label80, 12) & "' , '" & Mid(Label10, 12) & "'" & ",'" & Text4.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text4.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น2
Private Sub Text4_Change()
With RC
'        If (Text4.Text >= 0) And (Text4.Text <> "") Then
        On Error GoTo e1
        p2 = .Fields("ราคา") * Text4.text
        Label11.Caption = "รวมราคา " & p2 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา2
Private Sub ค้นหา2_Click()
With RC
If (Text3.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text3.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label80.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label10.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text4.Enabled = True
            Text3.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text3.text = ""
End If
End Sub





'ขายสินค้าที่3
Private Sub s3_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text6.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text6.text & " WHERE barcode = " & Text5.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s3.Enabled = False
        ค้นหา3.Enabled = False
        Text6.Enabled = False
        Text5.Enabled = False
        p_all = p_all + p3
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text6.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label81, 12) & "' , '" & Mid(Label15, 12) & "'" & ",'" & Text6.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text6.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น3
Private Sub Text6_Change()
With RC
'        If (Text6.Text >= 0) And (Text6.Text <> "") Then
        On Error GoTo e1
        p3 = .Fields("ราคา") * Text6.text
        Label16.Caption = "รวมราคา " & p3 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา3
Private Sub ค้นหา3_Click()
With RC
If (Text5.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text5.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label81.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label15.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text6.Enabled = True
            Text5.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text5.text = ""
End If
End Sub















'ขายสินค้าที่4
Private Sub s4_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text8.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text8.text & " WHERE barcode = " & Text7.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s4.Enabled = False
        ค้นหา4.Enabled = False
        Text8.Enabled = False
        Text7.Enabled = False
        p_all = p_all + p4
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text8.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label82, 12) & "' , '" & Mid(Label20, 12) & "'" & ",'" & Text8.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text8.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น4
Private Sub Text8_Change()
With RC
'        If (Text8.Text >= 0) And (Text8.Text <> "") Then
        On Error GoTo e1
        p4 = .Fields("ราคา") * Text8.text
        Label21.Caption = "รวมราคา " & p4 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา4
Private Sub ค้นหา4_Click()
With RC
If (Text7.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text7.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label82.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label20.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text8.Enabled = True
            Text7.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text7.text = ""
End If
End Sub






'ขายสินค้าที่5
Private Sub s5_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text10.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text10.text & " WHERE barcode = " & Text9.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s5.Enabled = False
        ค้นหา5.Enabled = False
        Text10.Enabled = False
        Text9.Enabled = False
        p_all = p_all + p5
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text10.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label83, 12) & "' , '" & Mid(Label25, 12) & "'" & ",'" & Text10.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text10.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น5
Private Sub Text10_Change()
With RC
'        If (Text10.Text >= 0) And (Text10.Text <> "") Then
        On Error GoTo e1
        p5 = .Fields("ราคา") * Text10.text
        Label26.Caption = "รวมราคา " & p5 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา5
Private Sub ค้นหา5_Click()
With RC
If (Text9.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text9.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label83.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label25.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text10.Enabled = True
            Text9.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text9.text = ""
End If
End Sub








'ขายสินค้าที่6
Private Sub s6_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text12.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text12.text & " WHERE barcode = " & Text11.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s6.Enabled = False
        ค้นหา6.Enabled = False
        Text12.Enabled = False
        Text11.Enabled = False
        p_all = p_all + p6
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text12.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label84, 12) & "' , '" & Mid(Label30, 12) & "'" & ",'" & Text12.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text12.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น6
Private Sub Text12_Change()
With RC
'        If (Text12.Text >= 0) And (Text12.Text <> "") Then
        On Error GoTo e1
        p6 = .Fields("ราคา") * Text12.text
        Label31.Caption = "รวมราคา " & p6 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา6
Private Sub ค้นหา6_Click()
With RC
If (Text11.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text11.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label84.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label30.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text12.Enabled = True
            Text11.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text11.text = ""
End If
End Sub









'ขายสินค้าที่7
Private Sub s7_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text14.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text14.text & " WHERE barcode = " & Text13.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s7.Enabled = False
        ค้นหา7.Enabled = False
        Text14.Enabled = False
        Text13.Enabled = False
        p_all = p_all + p7
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text14.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label85, 12) & "' , '" & Mid(Label35, 12) & "'" & ",'" & Text14.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text14.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น7
Private Sub Text14_Change()
With RC
'        If (Text14.Text >= 0) And (Text14.Text <> "") Then
        On Error GoTo e1
        p7 = .Fields("ราคา") * Text14.text
        Label36.Caption = "รวมราคา " & p7 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา7
Private Sub ค้นหา7_Click()
With RC
If (Text13.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text13.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label85.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label35.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text14.Enabled = True
            Text13.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text13.text = ""
End If
End Sub








'ขายสินค้าที่8
Private Sub s8_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text16.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text16.text & " WHERE barcode = " & Text15.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s8.Enabled = False
        ค้นหา8.Enabled = False
        Text16.Enabled = False
        Text15.Enabled = False
        p_all = p_all + p8
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text16.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label86, 12) & "' , '" & Mid(Label40, 12) & "'" & ",'" & Text16.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text16.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น8
Private Sub Text16_Change()
With RC
'        If (Text16.Text >= 0) And (Text16.Text <> "") Then
        On Error GoTo e1
        p8 = .Fields("ราคา") * Text16.text
        Label41.Caption = "รวมราคา " & p8 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา8
Private Sub ค้นหา8_Click()
With RC
If (Text15.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text15.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label86.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label40.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text16.Enabled = True
            Text15.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text15.text = ""
End If
End Sub








'ขายสินค้าที่9
Private Sub s9_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text18.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text18.text & " WHERE barcode = " & Text17.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s9.Enabled = False
        ค้นหา9.Enabled = False
        Text18.Enabled = False
        Text17.Enabled = False
        p_all = p_all + p9
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text18.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label125, 12) & "' , '" & Mid(Label45, 12) & "'" & ",'" & Text18.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text18.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น9
Private Sub Text18_Change()
With RC
'        If (Text18.Text >= 0) And (Text18.Text <> "") Then
        On Error GoTo e1
        p9 = .Fields("ราคา") * Text18.text
        Label46.Caption = "รวมราคา " & p9 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา9
Private Sub ค้นหา9_Click()
With RC
If (Text17.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text17.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label125.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label45.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text18.Enabled = True
            Text17.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text17.text = ""
End If
End Sub












'ขายสินค้าที่10
Private Sub s10_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text20.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text20.text & " WHERE barcode = " & Text19.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s10.Enabled = False
        ค้นหา10.Enabled = False
        Text20.Enabled = False
        Text19.Enabled = False
        p_all = p_all + p10
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text20.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label88, 12) & "' , '" & Mid(Label50, 12) & "'" & ",'" & Text20.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text20.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น10
Private Sub Text20_Change()
With RC
'        If (Text20.Text >= 0) And (Text20.Text <> "") Then
        On Error GoTo e1
        p10 = .Fields("ราคา") * Text20.text
        Label51.Caption = "รวมราคา " & p10 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา10
Private Sub ค้นหา10_Click()
With RC
If (Text19.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text19.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label88.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label50.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text20.Enabled = True
            Text19.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text19.text = ""
End If
End Sub






'ขายสินค้าที่11
Private Sub s11_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text22.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text22.text & " WHERE barcode = " & Text21.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s11.Enabled = False
        ค้นหา11.Enabled = False
        Text22.Enabled = False
        Text21.Enabled = False
        p_all = p_all + p11
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text22.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label89, 12) & "' , '" & Mid(Label55, 12) & "'" & ",'" & Text22.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text22.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น11
Private Sub Text22_Change()
With RC
'        If (Text22.Text >= 0) And (Text22.Text <> "") Then
        On Error GoTo e1
        p11 = .Fields("ราคา") * Text22.text
        Label56.Caption = "รวมราคา " & p11 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา11
Private Sub ค้นหา11_Click()
With RC
If (Text21.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text21.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label89.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label55.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text22.Enabled = True
            Text21.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text21.text = ""
End If
End Sub










'ขายสินค้าที่12
Private Sub s12_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text24.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text24.text & " WHERE barcode = " & Text23.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s12.Enabled = False
        ค้นหา12.Enabled = False
        Text24.Enabled = False
        Text23.Enabled = False
        p_all = p_all + p12
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text24.text
        Label98.Caption = num_s
        If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label90, 12) & "' , '" & Mid(Label60, 12) & "'" & ",'" & Text24.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text24.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น12
Private Sub Text24_Change()
With RC
'        If (Text24.Text >= 0) And (Text24.Text <> "") Then
        On Error GoTo e1
        p12 = .Fields("ราคา") * Text24.text
        Label61.Caption = "รวมราคา " & p12 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา12
Private Sub ค้นหา12_Click()
With RC
If (Text23.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text23.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label90.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label60.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text24.Enabled = True
            Text23.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text23.text = ""
End If
End Sub












'ขายสินค้าที่13
Private Sub s13_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text26.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text26.text & " WHERE barcode = " & Text25.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s13.Enabled = False
        ค้นหา13.Enabled = False
        Text26.Enabled = False
        Text25.Enabled = False
        p_all = p_all + p13
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text26.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label91, 12) & "' , '" & Mid(Label65, 12) & "'" & ",'" & Text26.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text26.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น13
Private Sub Text26_Change()
With RC
'        If (Text26.Text >= 0) And (Text26.Text <> "") Then
        On Error GoTo e1
        p13 = .Fields("ราคา") * Text26.text
        Label66.Caption = "รวมราคา " & p13 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา13
Private Sub ค้นหา13_Click()
With RC
If (Text25.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text25.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label91.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label65.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text26.Enabled = True
            Text25.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text25.text = ""
End If
End Sub












'ขายสินค้าที่14
Private Sub s14_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text28.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text28.text & " WHERE barcode = " & Text27.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s14.Enabled = False
        ค้นหา14.Enabled = False
        Text28.Enabled = False
        Text27.Enabled = False
        p_all = p_all + p14
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text28.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label92, 12) & "' , '" & Mid(Label70, 12) & "'" & ",'" & Text28.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
   
        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text28.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น14
Private Sub Text28_Change()
With RC
'        If (Text28.Text >= 0) And (Text28.Text <> "") Then
        On Error GoTo e1
        p14 = .Fields("ราคา") * Text28.text
        Label71.Caption = "รวมราคา " & p14 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา14
Private Sub ค้นหา14_Click()
With RC
If (Text27.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text27.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label92.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label70.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text28.Enabled = True
            Text27.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text27.text = ""
End If
End Sub














'ขายสินค้าที่15
Private Sub s15_Click()
With RC
        On Error GoTo e1
        If (.Fields("จำนวนสินค้า") - Text30.text >= 0) Then
        SQL = "UPDATE สินค้า  SET จำนวนสินค้า = จำนวนสินค้า -" & Text30.text & " WHERE barcode = " & Text29.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        s15.Enabled = False
        ค้นหา15.Enabled = False
        Text30.Enabled = False
        Text29.Enabled = False
        p_all = p_all + p15
        Label77.Caption = p_all
        num = num + 1
        Label96.Caption = num
        num_s = num_s + Text30.text
        Label98.Caption = num_s
         If (Label77 > 0) Then mon_in.Enabled = True
        'บันทึกไปยัง ตารางใบเสร็จ
        SQL = "INSERT INTO ใบเสร็จ (ชื่อสินค้า,ราคาชิ้นล่ะ,จำนวน) VALUES ('" & Mid(Label93, 12) & "' , '" & Mid(Label75, 12) & "'" & ",'" & Text30.text & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3

        Else
            MsgBox ("จำนวนสินค้าไม่พอสำหรับขายอยู่ " & Text30.text - .Fields("จำนวนสินค้า") & " ชิ้น")
        End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่รหัส barcode แล้วกด check และใส่จำนวนสินค้าที่ต้องการขาย")
End If
End Sub

'จำนวนชิ้น15
Private Sub Text30_Change()
With RC
'        If (Text30.Text >= 0) And (Text30.Text <> "") Then
        On Error GoTo e1
        p15 = .Fields("ราคา") * Text30.text
        Label76.Caption = "รวมราคา " & p15 & " บาท"
 '       End If
End With
If Error = 1 Then
e1:
    MsgBox ("กรุณาใส่จำนวนสินค้าเป็นตัวเลขให้ถูกต้อง")
End If
End Sub
'ค้นหา15
Private Sub ค้นหา15_Click()
With RC
If (Text29.text <> "") Then
        SQL = "SELECT * FROM สินค้า  WHERE barcode = " & Text29.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        On Error GoTo e1
        Label93.Caption = "ชื่อสินค้า " & .Fields("ชื่อสินค้า")
        Label75.Caption = "ราคาชิ้นละ " & .Fields("ราคา") & " บาท"
        
        If .Fields("ชื่อสินค้า") <> "" Then
            Text30.Enabled = True
            Text29.Enabled = False
        End If
Else
            MsgBox ("กรุณาใส่รหัส barcode")
End If
End With
If Error = 1 Then
e1:
    MsgBox ("ไม่มีสินค้าชนิดนี้อยู่ในร้าน")
    Text29.text = ""
End If
End Sub
Private Sub mon_in_Change()
Dim Money, money1 As Currency
Dim pay(8) As Currency
If (mon_in.text <> "") And (IsNumeric(mon_in) = True) Then
   Label102.Caption = "ทอน " & mon_in.text - p_all & " บาท"
    Money = mon_in.text - p_all
    If (Money >= 0) Then

    money1 = Int(Money)
    Label111.Caption = Int(money1 / 1000)
    Label112.Caption = Int((money1 Mod 1000) / 500)
    Label113.Caption = Int(((money1 Mod 1000) Mod 500) / 100)
    Label114.Caption = Int((((money1 Mod 1000) Mod 500) Mod 100) / 50)
    Label115.Caption = Int(((((money1 Mod 1000) Mod 500) Mod 100) Mod 50) / 20)
    Label116.Caption = Int((((((money1 Mod 1000) Mod 500) Mod 100) Mod 50) Mod 20) / 10)
    Label117.Caption = Int(((((((money1 Mod 1000) Mod 500) Mod 100) Mod 50) Mod 20) Mod 10) / 5)
    Label118.Caption = Int((((((((money1 Mod 1000) Mod 500) Mod 100) Mod 50) Mod 20) Mod 10) Mod 5) / 2)
    Label120.Caption = Int(((((((((money1 Mod 1000) Mod 500) Mod 100) Mod 50) Mod 20) Mod 10) Mod 5) Mod 2) / 1)
    
    'คำนวณเศษสตางค์
   Dim money2, money3 As Integer
   money2 = Money - money1
   money3 = money2 * 100
    Label122.Caption = Int(money3 / 50)
    Label124.Caption = Int((money3 Mod 50) / 25)
    End If
    Else
        MsgBox ("กรุณาใส่เงินที่ลูกค้าจ่ายตัวเลข")
        mon_in.text = "0.00"
End If
End Sub

Private Sub Print_pa_Click()
If (p_all > 0) And ((mon_in.text > p_all) Or (mon_in.text = p_all)) Then
        'ปิดการทำงานของปุ่ม print ใบเสร็จชั่วคราว
        Print_pa.Enabled = False
        On Error GoTo next1
        'ป้องกันการเปลี่ยนที่อยู่ของโปรแกรม เพื่อไม่ให้ DataEnvironment เกิด ERROR
        DataEnvironment1.Connection1.ConnectionString = App.Path & "\database\goods.mdb"
next1:
        'บอกราคารวมเข้าไปใน ใบเสร็จ
          DataReport1.Sections("Section3").Controls("Label6").Caption = Label77
          On Error GoTo makepaper
        'สร้างแฟ้มแยกเป็นวันที่ปัจจุบัน
          MkDir (App.Path & "\ใบเสร็จ\" & Format(Now(), "d mmm yyyy") & "\")
makepaper:
        'บันทึกใบเสร็จเป็น .txt ไว้ตามที่อยู่ที่กำหนดไว้โดยแยกแฟ้มเป็นวัน
         DataReport1.ExportReport rptKeyText, App.Path & "\ใบเสร็จ\" & Format(Now(), "d mmm yyyy") & "\" & "ใบเสร็จออก ณ วันที่ " & Format(Now(), "d mmm yyyy") & " เวลา " & Format(Now(), "hh - mm - ss") & ".txt", True
With RC
        'บันทึกไปยัง ตารางรวมเงิน_ใบเสร็จ
        SQL = "INSERT INTO รวมเงิน_ใบเสร็จ (รวมเงิน,เวลา) VALUES ('" & Label77 & "' , '" & Now() & "'" & ")"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
 End With
    DataReport1.Show
   MsgBox ("ถ้าคุณปิด หน้าต่าง แสดงตัวอย่างก่อนพิมพ์ใบเสร็จ ถ้าคุณต้องการจะพิมพ์ใบเสร็จนี้ใหม่กรุณาไปสั่งพิมพ์ที่เมนู ระบบค้นหาใบเสร็จ")
       With RC
        SQL = "DELETE * FROM ใบเสร็จ"
        If .State = 1 Then .Close
        .CursorLocation = 3
       .Open SQL, conn, 2, 3
       End With
 Else
    If (mon_in < p_all) Then MsgBox ("ลูกค้ายังชำระเงินไม่เต็มจำนวน ไม่สามารถออกใบเสร็จได้ กรุณาให้ลูกค้าชำระเต็มจำนวน")
    If (mon_in >= p_all) Then MsgBox ("รายการขายของท่านยังเป็น 0.00 บาท ถ้าคุณต้องการจะพิมพ์ใบเสร็จรายการขายก่อนหน้านี้ใหม่กรุณาไปสั่งพิมพ์ที่เมนู ระบบค้นหาใบเสร็จ")
End If
End Sub
