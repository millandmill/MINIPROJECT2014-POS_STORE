VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ѡ"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13575
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   13575
   Begin VB.CommandButton Command13 
      Caption         =   "�к����ҫѾ���������"
      Height          =   855
      Left            =   3000
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "�к������Թ���"
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "�к������ʹ��������ѹ"
      Height          =   855
      Left            =   3000
      TabIndex        =   10
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "��§ҹ�ʴ���������´�ͧ�Ѻ���������"
      Height          =   855
      Left            =   9480
      TabIndex        =   9
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "��§ҹ�ʴ���¡���Թ��ҷ�����������ҹ"
      Height          =   855
      Left            =   5880
      TabIndex        =   8
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�к����������"
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�͡�ҡ�����"
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   12615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�к��ʴ��Թ��ҷ��������"
      Height          =   855
      Left            =   9480
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�к������Թ��ҷ������������ص�����͹�"
      Height          =   855
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�к�����Թ���"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�Ѵ��� �Ѿ���������"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Ѵ��û������Թ���"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ѵ����Թ����к�"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
End Sub

Private Sub Command10_Click()
        On Error GoTo next1
        '��ͧ�ѹ�������¹�������ͧ����� ���������� DataEnvironment �Դ ERROR
        DataEnvironment1.Connection1.ConnectionString = App.Path & "\database\goods.mdb"
next1:
        DataReport3.Show
End Sub

Private Sub Command11_Click()
    Form10.Show
End Sub

Private Sub Command12_Click()
    Form11.Show
End Sub

Private Sub Command13_Click()
    Form12.Show
End Sub

Private Sub Command2_Click()
    Form2.Show
End Sub

Private Sub Command3_Click()
    Form3.Show
End Sub

Private Sub Command4_Click()
    Form4.Show
End Sub

Private Sub Command5_Click()
    Form5.Show
End Sub

Private Sub Command6_Click()
    Form6.Show
End Sub

Private Sub Command7_Click()
    End
End Sub

Private Sub Command8_Click()
    Form9.Show
End Sub

Private Sub Command9_Click()
        On Error GoTo next1
        '��ͧ�ѹ�������¹�������ͧ����� ���������� DataEnvironment �Դ ERROR
        DataEnvironment1.Connection1.ConnectionString = App.Path & "\database\goods.mdb"
next1:
        DataReport2.Show
End Sub
