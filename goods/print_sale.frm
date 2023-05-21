VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "แสดงตัวอย่างใบเสร็จที่พิมพ์"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "Form8"
   ScaleHeight     =   4590
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   6855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   6855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   6855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   6855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   6855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label Label63_1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cordia New"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   6855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label1.Caption = Form9.Label1.Caption
Label2.Caption = Form9.Label2.Caption
Label3.Caption = Form9.Label3.Caption
Label4.Caption = Form9.Label4.Caption
Label5.Caption = Form9.Label5.Caption
Label6.Caption = Form9.Label6.Caption
Label7.Caption = Form9.Label7.Caption
Label8.Caption = Form9.Label8.Caption
Label9.Caption = Form9.Label9.Caption
Label10.Caption = Form9.Label10.Caption
Label11.Caption = Form9.Label11.Caption
Label12.Caption = Form9.Label12.Caption
Label13.Caption = Form9.Label13.Caption
Label14.Caption = Form9.Label14.Caption
Label15.Caption = Form9.Label15.Caption
Label16.Caption = Form9.Label16.Caption
Label17.Caption = Form9.Label17.Caption
Label63_1.Caption = Form9.Label63_1.Caption

Form8.PrintForm
End Sub


