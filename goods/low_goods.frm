VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�к��ʴ��Թ��ҷ��������"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8430
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8430
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   300
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "��ѧ�������á"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   300
      Left            =   720
      TabIndex        =   16
      ToolTipText     =   "��ѧ�����š�͹˹�ҹ��"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   300
      Left            =   1200
      TabIndex        =   15
      ToolTipText     =   "��ѧ�����ŵ���"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   300
      Left            =   1680
      TabIndex        =   14
      ToolTipText     =   "��ѧ�������ش����"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "�óը���觫����Թ�������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   5055
      Begin VB.CommandButton Command6 
         Caption         =   "������š�ä���"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "���ҫѾ���������"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "0000000000000"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "�������Ѿ���Шӷ��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "�������Ѿ����Ͷ��"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "���ʫѾ���������"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "����barcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton �͡����� 
      Caption         =   "�͡"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "low_goods.frx":0000
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label �Ѻ 
      Caption         =   "�ըӹǹ�Թ��ҷ�������������ӹǹ����˹�  XX ���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   6495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL, SQL1 As String
Dim d_m, m_m, y_m, d_o, m_o, y_o, number_goods As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"

Private Sub Form_Load()
'�Դ�ҹ������
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With
'�ӹǳ�ѹ��������
With RC
        
       SQL = "SELECT barcode,�����Թ���,�ӹǹ�Թ���, �������͹ FROM �Թ��� WHERE �ӹǹ�Թ��� <= �������͹ "
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
   �Ѻ.Caption = "���Թ��ҷ���������������͹䢡�ä��ҷ����� " & .RecordCount & " ��¡��"
End With
End Sub
Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
End If
End With
End Sub
Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
End If
End With
End Sub
Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
End If
End With
End Sub
'���ҫѾ���������
Private Sub Command5_Click()
With RC
 Dim a, b As String
        On Error GoTo e1
       SQL1 = "SELECT * FROM �Ѻ��������� ,�Թ��� WHERE barcode =  " & Text8.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL1, conn, 2, 3
        
         Label5.Caption = .Fields("�Ѻ���������.���ʫѺ���������")
         Label3.Caption = .Fields("���Ѿ����Ͷ��")
         Label4.Caption = .Fields("���Ѿ��")
Call Form_Load
End With
If Error = 1 Then
e1:
    MsgBox ("��辺���� barcode �����к�")
    Text8.text = "0000000000000"
    Call Form_Load
End If
End Sub
'������š�ä���
Private Sub Command6_Click()
        Label5.Caption = "-"
         Label3.Caption = "-"
         Label4.Caption = "-"
         Text8.text = "0000000000000"
End Sub
Private Sub �͡�����_Click()
    Unload Form6
End Sub
