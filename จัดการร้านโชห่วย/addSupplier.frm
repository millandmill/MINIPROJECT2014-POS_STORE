VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ѵ��ëѾ���������"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16515
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   16515
   Begin VB.TextBox email 
      Height          =   285
      Left            =   3480
      TabIndex        =   26
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton ������ 
      Caption         =   "������˹�Ҩ�"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":0000
      TabIndex        =   25
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox com 
      Height          =   735
      Left            =   2520
      TabIndex        =   23
      Text            =   "-"
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      ToolTipText     =   "��ѧ�������ش����"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      ToolTipText     =   "��ѧ�����ŵ���"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      ToolTipText     =   "��ѧ�����š�͹˹�ҹ��"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      ToolTipText     =   "��ѧ�������á"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox phone 
      Height          =   285
      Left            =   3480
      MaxLength       =   9
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox mobile 
      Height          =   285
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   15
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������"
      Height          =   1815
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   5295
      Begin VB.TextBox zip 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "00000"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox prov 
         Height          =   315
         ItemData        =   "addSupplier.frx":2CFA
         Left            =   1320
         List            =   "addSupplier.frx":2CFC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox addr 
         Height          =   525
         Left            =   1320
         TabIndex        =   9
         Text            =   "00/00"
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "������ɳ���"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "�ѧ��Ѵ"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "�������"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox namep 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton �����ѹ�֡ 
      Caption         =   "����/�ѹ�֡"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":2CFE
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton ��䢺ѹ�֡ 
      Caption         =   "���/�ѹ�֡"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":39C8
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton ź������ 
      Caption         =   "ź������"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":66C2
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton �͡����� 
      Caption         =   "�͡"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":9044
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "addSupplier.frx":B9C6
      Height          =   5775
      Left            =   7320
      TabIndex        =   24
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Caption         =   "�Ѻ�Ѿ���������"
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   6240
      Width           =   3735
   End
   Begin VB.Label Label9 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "�����˵�"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "�������Ѿ���Шӷ��"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "�������Ѿ����Ͷ��"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "���ͺ���ѷ/�ؤ��"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "���ʫѾ��������� :"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
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

'�Դ���ҧ
With RC
        SQL = "SELECT * FROM �Ѻ���������  ORDER BY ���ʫѺ��������� ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ.Caption = "�Ѿ����������շ����� " & .RecordCount & " ���"
End With

prov.AddItem "��к��"
prov.AddItem "��ا෾��ҹ��"
prov.AddItem "�ҭ������"
prov.AddItem "����Թ���"
prov.AddItem "��ᾧྪ�"
prov.AddItem "�͹��"
prov.AddItem "�ѹ�����"
prov.AddItem "���ԧ���"
prov.AddItem "�ź���"
prov.AddItem "��¹ҷ"
prov.AddItem "�������"
prov.AddItem "�����"
prov.AddItem "��§���"
prov.AddItem "��§����"
prov.AddItem "��ѧ"
prov.AddItem "��Ҵ"
prov.AddItem "�ҡ"
prov.AddItem "��ù�¡"
prov.AddItem "��û��"
prov.AddItem "��þ��"
prov.AddItem "����Ҫ����"
prov.AddItem "�����ո����Ҫ"
prov.AddItem "������ä�"
prov.AddItem "�������"
prov.AddItem "��Ҹ����"
prov.AddItem "��ҹ"
prov.AddItem "���������"
prov.AddItem "�����ҹ�"
prov.AddItem "��ШǺ���բѹ��"
prov.AddItem "��Ҩչ����"
prov.AddItem "�ѵ�ҹ�"
prov.AddItem "��й�������ظ��"
prov.AddItem "�����"
prov.AddItem "�ѧ��"
prov.AddItem "�ѷ�ا"
prov.AddItem "�ԨԵ�"
prov.AddItem "��ɳ��š"
prov.AddItem "ྪú���"
prov.AddItem "ྪú�ó�"
prov.AddItem "���"
prov.AddItem "����"
prov.AddItem "�����ä��"
prov.AddItem "�ء�����"
prov.AddItem "�����ͧ�͹"
prov.AddItem "��ʸ�"
prov.AddItem "����"
prov.AddItem "�������"
prov.AddItem "�йͧ"
prov.AddItem "���ͧ"
prov.AddItem "�Ҫ����"
prov.AddItem "ž����"
prov.AddItem "���"
prov.AddItem "�ӻҧ"
prov.AddItem "�Ӿٹ"
prov.AddItem "�������"
prov.AddItem "ʡŹ��"
prov.AddItem "ʧ���"
prov.AddItem "ʵ��"
prov.AddItem "��طû�ҡ��"
prov.AddItem "��ط�ʧ����"
prov.AddItem "��ط��Ҥ�"
prov.AddItem "������"
prov.AddItem "��к���"
prov.AddItem "�ԧ�����"
prov.AddItem "��⢷��"
prov.AddItem "�ؾ�ó����"
prov.AddItem "����ɮ��ҹ�"
prov.AddItem "���Թ���"
prov.AddItem "˹ͧ���"
prov.AddItem "˹ͧ�������"
prov.AddItem "��ҧ�ͧ"
prov.AddItem "�ӹҨ��ԭ"
prov.AddItem "�شøҹ�"
prov.AddItem "�صôԵ��"
prov.AddItem "�ط�¸ҹ�"
prov.AddItem "�غ��Ҫ�ҹ�"

prov.text = "��ا෾��ҹ��"
End Sub

Private Sub zip_Change()
    If (IsNumeric(zip.text) = False) Then
        MsgBox ("��س����������ɳ����繵���Ţ �ӹǹ 5 ��ѡ")
        zip.text = "00000"
    End If
End Sub
Private Sub ������_Click()
            Label1.Caption = "���ʫѾ��������� : "
            namep.text = ""
            addr.text = ""
            zip.text = "00000"
            mobile.text = ""
            phone.text = ""
            email.text = ""
            com.text = ""
End Sub
Private Sub ��䢺ѹ�֡_Click()
With RC
If .RecordCount = 0 Then
    MsgBox "�������ö���/�ѹ�֡�����ͧ�ҡ ����իѺ����������������", vbInformation, "���/�ѹ�֡"
ElseIf (namep.text <> "") And (addr.text <> "") And Len(zip.text) = 5 Then
            .Fields("����") = namep.text
            .Fields("�������") = addr.text
            .Fields("�ѧ��Ѵ") = prov.text
            .Fields("������ɳ���") = zip.text
            .Fields("���Ѿ����Ͷ��") = mobile.text
            .Fields("���Ѿ��") = phone.text
            .Fields("������������") = email.text
            .Fields("�����˵�") = com.text
           .Update
            MsgBox "���/�ѹ�֡�����ūѺ���������������º��������", vbInformation, "���/�ѹ�֡"
Else
            If (Len(zip.text) < 5) Then MsgBox "��س����������ɳ��� 5 ��ѡ"
            MsgBox "��س��������ͺ���ѷ/�ؤ�� ������� ������ɳ���  �������Ѿ��������º����", vbInformation, "���/�ѹ�֡"
End If
End With
If Error = 1 Then
error5:
    Unload Form3
    Form3.Show
End If
End Sub
Private Sub �����ѹ�֡_Click()
With RC
If (namep.text <> "") And (addr.text <> "") And Len(zip.text) = 5 Then
            On Error GoTo error1
            .AddNew
            .Fields("����") = namep.text
            .Fields("�������") = addr.text
            .Fields("�ѧ��Ѵ") = prov.text
            .Fields("������ɳ���") = zip.text
            .Fields("���Ѿ����Ͷ��") = mobile.text
            .Fields("���Ѿ��") = phone.text
            .Fields("������������") = email.text
            .Fields("�����˵�") = com.text
            .Fields("�ӹǹ�Թ���") = 0
        .Update
        �Ѻ.Caption = "�Ѿ����������շ����� " & .RecordCount & " ���"
        MsgBox "���������ūѺ����������������º����", vbInformation, "���������ūѺ���������"
        Call Command1_Click
Else
            If (Len(zip.text) < 5) Then MsgBox "��س����������ɳ��� 5 ��ѡ"
            MsgBox "��س��������ͺ���ѷ/�ؤ�� ������� ������ɳ���  �������Ѿ��������º����", vbInformation, "���/�ѹ�֡"
End If
End With
If Error = 1 Then
error1:
        MsgBox "��سҺѹ�֡�����ūѾ����������ա���� �Ҩ���������բ����Ź��������к����� ���� ��á�͡�����żԴ�ٻẺ", vbInformation, "�����������Թ��ҼԴ��Ҵ"
         MsgBox "�к��зӡ�ûԴ�к��Ѵ��ëѾ��������� ��Ҽ�����ͧ������������ūѾ���������������� ��س�����к��Ѵ��ëѾ����������ա����", vbInformation, "��ͤ����ҡ�к�"
         Call Form_Load
         Unload Form3
End If
End Sub

Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
            Label1.Caption = "���ʫѾ��������� : " & .Fields("���ʫѺ���������")
            namep.text = .Fields("����")
            addr.text = .Fields("�������")
            prov.text = .Fields("�ѧ��Ѵ")
            zip.text = .Fields("������ɳ���")
            mobile.text = .Fields("���Ѿ����Ͷ��")
            phone.text = .Fields("���Ѿ��")
            email.text = .Fields("������������")
            com.text = .Fields("�����˵�")
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
            Label1.Caption = "���ʫѾ��������� : " & .Fields("���ʫѺ���������")
            namep.text = .Fields("����")
            addr.text = .Fields("�������")
            prov.text = .Fields("�ѧ��Ѵ")
            zip.text = .Fields("������ɳ���")
            mobile.text = .Fields("���Ѿ����Ͷ��")
            phone.text = .Fields("���Ѿ��")
            email.text = .Fields("������������")
            com.text = .Fields("�����˵�")
End If
End With
End Sub


Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
            Label1.Caption = "���ʫѾ��������� : " & .Fields("���ʫѺ���������")
            namep.text = .Fields("����")
            addr.text = .Fields("�������")
            prov.text = .Fields("�ѧ��Ѵ")
            zip.text = .Fields("������ɳ���")
            mobile.text = .Fields("���Ѿ����Ͷ��")
            phone.text = .Fields("���Ѿ��")
            email.text = .Fields("������������")
            com.text = .Fields("�����˵�")
End If
End With
End Sub

Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
            Label1.Caption = "���ʫѾ��������� : " & .Fields("���ʫѺ���������")
            namep.text = .Fields("����")
            addr.text = .Fields("�������")
            prov.text = .Fields("�ѧ��Ѵ")
            zip.text = .Fields("������ɳ���")
            mobile.text = .Fields("���Ѿ����Ͷ��")
            phone.text = .Fields("���Ѿ��")
            email.text = .Fields("������������")
            com.text = .Fields("�����˵�")
End If
End With
End Sub

Private Sub ź������_Click()
With RC
On Error GoTo del
If (.RecordCount > 0) And (.Fields("�ӹǹ�Թ���") = 0) Then
        Dim sid As String
        sid = .Fields("���ʫѺ���������")
        .Delete
        .Requery
        �Ѻ.Caption = "�Ѿ����������շ����� " & .RecordCount & " ���"
        MsgBox "ź�����ūѺ������������� " & sid & " ���º��������", vbInformation, "ź�����ūѺ���������"
Else
del:
        MsgBox "�������öź�����ūѺ��������������ͧ�ҡ����իѺ������������ź ���� �Ѻ������������ͧ��è�ź�ѧ���Թ��ҷ������Ǣ�ͧ���� ", vbInformation, "ź�����ūѺ���������"
End If
End With
Call Command1_Click
End Sub
Private Sub �͡�����_Click()
    Unload Form3
End Sub
