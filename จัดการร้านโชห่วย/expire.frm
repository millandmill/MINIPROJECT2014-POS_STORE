VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�к��ʴ��Թ��ҷ������������ط����¡��� 1 ��͹"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10125
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10125
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   300
      Left            =   1680
      TabIndex        =   9
      ToolTipText     =   "��ѧ�������ش����"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   300
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "��ѧ�����ŵ���"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   300
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "��ѧ�����š�͹˹�ҹ��"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "��ѧ�������á"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   20.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   4
      Text            =   "30"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton �͡����� 
      Caption         =   "�͡"
      Height          =   615
      Left            =   8400
      Picture         =   "expire.frx":0000
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
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
      Caption         =   "���Թ��ҷ���������������͹䢡�ä��ҷ����� 0 ��¡��"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   20.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "�ѹ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   20.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label �Թ��� 
      Caption         =   "���Թ��ҷ�����������������"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   20.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
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

        SQL = "UPDATE �Թ��� SET �ѹ�������� =  �ѹ������� - �ѹ��Ե"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        
       SQL = "SELECT barcode,�����Թ���,�ѹ��Ե,�ѹ�������,�ӹǹ�Թ��� FROM �Թ��� WHERE �ѹ�������� <= 30"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ.Caption = "���Թ��ҷ���������������͹䢡�ä��ҷ����� " & .RecordCount & " ��¡��"
End With
End Sub

Private Sub Text1_Change()
If (IsNumeric(Text1.text) = True) Then
With RC
        SQL = "SELECT barcode,�����Թ���,�ѹ��Ե,�ѹ�������,�ӹǹ�Թ���  FROM �Թ��� WHERE �ѹ�������� <= " & Text1.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ.Caption = "���Թ��ҷ���������������͹䢡�ä��ҷ����� " & .RecordCount & " ��¡��"
End With
Else
    MsgBox "��س�����ѹ�繵���Ţ"
    Text1.text = ""
End If
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
Private Sub �͡�����_Click()
    Unload Form5
End Sub
