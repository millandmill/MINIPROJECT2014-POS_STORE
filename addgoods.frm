VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ѵ����Թ��Ңͧ�к�"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10425
   Begin VB.ComboBox yy 
      Height          =   315
      ItemData        =   "addgoods.frx":0000
      Left            =   4920
      List            =   "addgoods.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox mm 
      Height          =   315
      ItemData        =   "addgoods.frx":0004
      Left            =   3960
      List            =   "addgoods.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox dd 
      Height          =   315
      ItemData        =   "addgoods.frx":0008
      Left            =   3000
      List            =   "addgoods.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox y 
      Height          =   315
      ItemData        =   "addgoods.frx":000C
      Left            =   4920
      List            =   "addgoods.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox m 
      Height          =   315
      ItemData        =   "addgoods.frx":0010
      Left            =   3960
      List            =   "addgoods.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox d 
      Height          =   315
      ItemData        =   "addgoods.frx":0014
      Left            =   3000
      List            =   "addgoods.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox p 
      Height          =   285
      Left            =   3720
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox nameg 
      Height          =   285
      Left            =   3000
      TabIndex        =   30
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton ������ 
      Caption         =   "������˹�Ҩ�"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":0018
      TabIndex        =   29
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      ToolTipText     =   "��ѧ�������á"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3120
      TabIndex        =   25
      ToolTipText     =   "��ѧ�����š�͹˹�ҹ��"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      ToolTipText     =   "��ѧ�����ŵ���"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      ToolTipText     =   "��ѧ�������ش����"
      Top             =   3840
      Width           =   855
   End
   Begin VB.ComboBox sup 
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox lessgood 
      Height          =   285
      Left            =   3720
      TabIndex        =   19
      Text            =   "0"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox numgood 
      Height          =   285
      Left            =   3720
      TabIndex        =   16
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox id 
      Height          =   285
      Left            =   3000
      MaxLength       =   13
      TabIndex        =   13
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox ������ 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "�ô���͡������"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton �����ѹ�֡ 
      Caption         =   "����/�ѹ�֡"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":2D12
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton ��䢺ѹ�֡ 
      Caption         =   "���/�ѹ�֡"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":39DC
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton ź������ 
      Caption         =   "ź������"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":66D6
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton �͡����� 
      Caption         =   "�͡"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":9058
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "addgoods.frx":B9DA
      Height          =   1215
      Left            =   240
      TabIndex        =   27
      Top             =   4440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2143
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
   Begin VB.Label Label14 
      Caption         =   "�ѹ / ��͹ / �� �.�."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   36
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "�ѹ / ��͹ / �� �.�."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "�ҷ"
      Height          =   255
      Left            =   5040
      TabIndex        =   34
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "�Ҥ��Թ���"
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "�����Թ���"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label �Ѻ�Թ��� 
      Caption         =   "�Ѻ�ӹǹ�Թ���"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label12 
      Caption         =   "�Թ��Ҫ�鹹����觨ҡ�Ѿ�������������"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "����������͹"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "��Ҩӹǹ�Թ��ҹ��¡���"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "���"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "�ӹǹ�Թ��ҷ���������ѧ"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "���� barcode ��ͧ�� 13 ��ѡ��ҹ��"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "����barcode"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "/"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "/"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "�ѹ�������"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "/"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "/"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "�ѹ��Ե"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�������Թ���"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Dim namep, ������ź, �Ѻ��������� As String
Dim flag As Integer
Dim d_m, m_m, y_m, d_o, m_o, y_o, number_goods, del, aa, i As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
Private Sub id_Change()
On Error GoTo error1
If IsNumeric(id.text) = False Then
        MsgBox "��س������͹����Ե�Թ������١��ͧ ����繵���Ţ 13 ��ѡ", vbInformation, "����͹"
error1:
        id.text = ""
End If
End Sub
Private Sub Form_Load()

'�ѹ��Ե
For i = 1 To 31
d.AddItem i
Next
d = 1

'��͹��Ե
For i = 1 To 12
m.AddItem i
Next
m = 1

'�ռ�Ե
For i = Year(Now()) - 10 To Year(Now()) + 10
y.AddItem i
Next
y = Year(Now())

'�ѹ�������
For i = 1 To 31
dd.AddItem i
Next
dd = 1

'��͹�������
For i = 1 To 12
mm.AddItem i
Next
mm = 1

'���������
For i = Year(Now()) - 10 To Year(Now()) + 10
yy.AddItem i
Next
yy = Year(Now())

'�Դ�ҹ������
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With

'�Դ���ҧ ������ ������� ������ ��� Combo box
With RC
        SQL = "SELECT * FROM ������"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Do While Not RC.EOF
            ������.AddItem .Fields("������")
            RC.MoveNext
        Loop
        RC.Close
End With

'�Դ���ҧ �Ѻ��������� ������� ID ��� Combo box
With RC
        SQL = "SELECT * FROM �Ѻ���������"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Do While Not RC.EOF
            sup.AddItem .Fields("���ʫѺ���������")
            RC.MoveNext
        Loop
        RC.Close
End With

'�Դ���ҧ
With RC
        SQL = "SELECT * FROM �Թ���  ORDER BY �ѹ������� ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ�Թ���.Caption = "���Թ�����к������� " & .RecordCount & "  ��¡�� "
End With
End Sub

Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
            sup.Enabled = False
            ������.Enabled = False
            �����ѹ�֡.Enabled = False
            id.text = .Fields("barcode")
            nameg.text = .Fields("�����Թ���")
            ������.text = .Fields("������")
            '�ѹ��Ե
            d.text = Day(.Fields("�ѹ��Ե"))
            m.text = Month(.Fields("�ѹ��Ե"))
            y.text = Year(.Fields("�ѹ��Ե"))
            '�ѹ�������
            dd.text = Day(.Fields("�ѹ�������"))
            mm.text = Month(.Fields("�ѹ�������"))
            yy.text = Year(.Fields("�ѹ�������"))
        
            numgood.text = .Fields("�ӹǹ�Թ���")
            lessgood.text = .Fields("�������͹")
            p.text = .Fields("�Ҥ�")
            sup.text = .Fields("���ʫѺ���������")
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
            sup.Enabled = False
            ������.Enabled = False
            �����ѹ�֡.Enabled = False
            id.text = .Fields("barcode")
            nameg.text = .Fields("�����Թ���")
            ������.text = .Fields("������")
            '�ѹ��Ե
            d.text = Day(.Fields("�ѹ��Ե"))
            m.text = Month(.Fields("�ѹ��Ե"))
            y.text = Year(.Fields("�ѹ��Ե"))
            '�ѹ�������
            dd.text = Day(.Fields("�ѹ�������"))
            mm.text = Month(.Fields("�ѹ�������"))
            yy.text = Year(.Fields("�ѹ�������"))
        
            numgood.text = .Fields("�ӹǹ�Թ���")
            lessgood.text = .Fields("�������͹")
            p.text = .Fields("�Ҥ�")
            sup.text = .Fields("���ʫѺ���������")
End If
End With
End Sub


Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
            sup.Enabled = False
            ������.Enabled = False
            �����ѹ�֡.Enabled = False
           id.text = .Fields("barcode")
            nameg.text = .Fields("�����Թ���")
            ������.text = .Fields("������")
            '�ѹ��Ե
            d.text = Day(.Fields("�ѹ��Ե"))
            m.text = Month(.Fields("�ѹ��Ե"))
            y.text = Year(.Fields("�ѹ��Ե"))
            '�ѹ�������
            dd.text = Day(.Fields("�ѹ�������"))
            mm.text = Month(.Fields("�ѹ�������"))
            yy.text = Year(.Fields("�ѹ�������"))
        
            numgood.text = .Fields("�ӹǹ�Թ���")
            lessgood.text = .Fields("�������͹")
            p.text = .Fields("�Ҥ�")
            sup.text = .Fields("���ʫѺ���������")
End If
End With
End Sub

Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
            sup.Enabled = False
            ������.Enabled = False
            �����ѹ�֡.Enabled = False
           id.text = .Fields("barcode")
            nameg.text = .Fields("�����Թ���")
            ������.text = .Fields("������")
            '�ѹ��Ե
            d.text = Day(.Fields("�ѹ��Ե"))
            m.text = Month(.Fields("�ѹ��Ե"))
            y.text = Year(.Fields("�ѹ��Ե"))
            '�ѹ�������
            dd.text = Day(.Fields("�ѹ�������"))
            mm.text = Month(.Fields("�ѹ�������"))
            yy.text = Year(.Fields("�ѹ�������"))
        
            numgood.text = .Fields("�ӹǹ�Թ���")
            lessgood.text = .Fields("�������͹")
            p.text = .Fields("�Ҥ�")
            sup.text = .Fields("���ʫѺ���������")
End If
End With

End Sub

Public Sub ������_Click()
            sup.Enabled = True
            ������.Enabled = True
            �����ѹ�֡.Enabled = True
            id.text = "0000000000000"
            nameg.text = ""
            d.text = "1"
            m.text = "1"
            y.text = Year(Now())
            dd.text = "1"
            mm.text = "1"
            yy.text = Year(Now())
             numgood.text = "0"
            lessgood.text = "0"
            p.text = "0.00"
End Sub

Private Sub �����ѹ�֡_Click()
With RC
Dim flag As Integer
Dim d1, d2 As Date

            flag = 0
            If (������.text = "") Then
                MsgBox "��س��������Ż�����"
                flag = 1
            End If
            
            If (sup.text = "") Then
                MsgBox "��س�������ʫѾ���������"
                flag = 1
            End If
            
            If (nameg.text = "") Then
                MsgBox "��س��������Թ���"
                flag = 1
             End If
             
            If (id.text = "") Or Len(id.text) < 13 Then
                MsgBox "��س�������� barcode �繵���Ţ���ú 13 ���"
                flag = 1
            End If
            
            On Error GoTo e2
            If (p.text <= 0) Then
e2:
                MsgBox "��س�����Ҥ��Թ����ҡ���� 0 �ҷ"
                flag = 1
            End If
            
            '�� � / �/ �
            If (yy.text < y.text) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text < m.text)) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text = m.text) And (dd.text < d.text)) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            
            
           If (flag = 0) Then
            .AddNew
            .Fields("barcode") = id.text
            .Fields("�����Թ���") = nameg.text
            .Fields("������") = ������.text
            '�ѹ��Ե
            d_m = d.text
            m_m = m.text
            y_m = y.text
            .Fields("�ѹ��Ե") = DateSerial(y_m, m_m, d_m)
            '�ѹ�������
            d_o = dd.text
            m_o = mm.text
            y_o = yy.text
            .Fields("�ѹ�������") = DateSerial(y_o, m_o, d_o)
            
            .Fields("�ӹǹ�Թ���") = numgood.text
            .Fields("�������͹") = lessgood.text
             .Fields("�Ҥ�") = p.text
            .Fields("���ʫѺ���������") = sup.text
        On Error GoTo error1
        .Update
        MsgBox "�����������Թ����������º����", vbInformation, "�����������Թ���"
        End If
End With

        '�ѹ�֡�Թ��� ��ѧ������ ���ͺѹ�֡�ӹǹ
        
With RC
        If (flag = 0) Then
        SQL = "UPDATE ������ SET �ӹǹ�Թ��һ�������� = �ӹǹ�Թ��һ�������� +1 WHERE ������ = " & "'" & ������.text & "'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        End If
        
        '�ѹ�֡�Թ��� ��ѧ�Ѻ��������� ���ͺѹ�֡�ӹǹ
        If (flag = 0) Then
        SQL = "UPDATE �Ѻ��������� SET �ӹǹ�Թ��� = �ӹǹ�Թ��� +1 WHERE ���ʫѺ��������� = " & sup.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        End If
        
        '�Դ���ҧ�����ա����
        SQL = "SELECT * FROM �Թ���  ORDER BY �ѹ������� ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ�Թ���.Caption = "���Թ�����к������� " & .RecordCount & "  ��¡�� "
End With
Call Command1_Click
        
If Error = 0 Then
error1:
        flag = 1
        MsgBox "��سҺѹ�֡�Թ��������ա���� �Ҩ���������Թ��Ҫ�鹹��������к����� ���� ��á�͡�����żԴ�ٻẺ", vbInformation, "�����������Թ��ҼԴ��Ҵ"
         Call Form_Load
         Unload Form1
         Form1.Show
End If

End Sub
Private Sub ��䢺ѹ�֡_Click()
With RC
Dim flag As Integer
flag = 0
If .RecordCount = 0 Then
    MsgBox "�������ö���/�ѹ�֡�����ͧ�ҡ ������Թ���������", vbInformation, "���/�ѹ�֡"
    flag = 1
End If
            If (������.text = "") Then
                MsgBox "��س��������Ż�����"
                flag = 1
            End If
            
            If (sup.text = "") Then
                MsgBox "��س�������ʫѾ���������"
                flag = 1
            End If
            
            If (nameg.text = "") Then
                MsgBox "��س��������Թ���"
                flag = 1
             End If
             
             
             If (id.text = "") Or Len(id.text) < 13 Then
                MsgBox "��س�������� barcode �繵���Ţ���ú 13 ���"
                flag = 1
            End If
            
            On Error GoTo e2
            If (p.text <= 0) Then
e2:
                MsgBox "��س�����Ҥ��Թ����ҡ���� 0 �ҷ"
                flag = 1
            End If
            
            
             '�� � / �/ �
            If (yy.text < y.text) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text < m.text)) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text = m.text) And (dd.text < d.text)) Then
                MsgBox "��س����ѹ������آͧ�Թ������١��ͧ ��� �ѹ������ص�ͧ��������͹ �ѹ��Ե"
                flag = 1
            End If
            
           If (flag = 0) Then
            .Fields("barcode") = id.text
            .Fields("�����Թ���") = nameg.text
            .Fields("������") = ������.text
            '�ѹ��Ե
            d_m = d.text
            m_m = m.text
            y_m = y.text
            .Fields("�ѹ��Ե") = DateSerial(y_m, m_m, d_m)
            '�ѹ�������
            d_o = dd.text
            m_o = mm.text
            y_o = yy.text
            .Fields("�ѹ�������") = DateSerial(y_o, m_o, d_o)
            
            .Fields("�ӹǹ�Թ���") = numgood.text
            .Fields("�������͹") = lessgood.text
            .Fields("�Ҥ�") = p.text
            .Fields("���ʫѺ���������") = sup.text
            On Error GoTo error1
           .Update
           Call Command1_Click
            MsgBox "���/�ѹ�֡�Թ���������º��������", vbInformation, "���/�ѹ�֡"
            �Ѻ�Թ���.Caption = "���Թ�����к������� " & .RecordCount & "  ��¡�� "
End If
End With
If Error = 0 Then
error1:
        MsgBox "��سҺѹ�֡�Թ��������ա���� �Ҩ���������Թ��Ҫ�鹹��������к����� ���� ��á�͡�����żԴ�ٻẺ", vbInformation, "�����������Թ��ҼԴ��Ҵ"
         Call Form_Load
         Unload Form1
         Form1.Show
End If
End Sub
Private Sub ź������_Click()
Dim flag As Integer
flag = 0
With RC
If .RecordCount > 0 Then
        flag = 1
        namep = .Fields("�����Թ���")
        ������ź = .Fields("������")
        �Ѻ��������� = .Fields("���ʫѺ���������")
        .Delete
        .Requery
        MsgBox "ź������˹ѧ��ͪ��� " & namep & " ���º��������", vbInformation, "ź������˹ѧ���"
Else
        MsgBox "�������öź������˹ѧ��������ͧ�ҡ�����˹ѧ������ź", vbInformation, "ź������˹ѧ���"
End If
End With
       '�ѹ�֡�Թ��� ��ѧ������ ����ź�ӹǹ
       If (flag = 1) Then
       flag = 0
        With RC
        SQL = "UPDATE ������ SET �ӹǹ�Թ��һ�������� = �ӹǹ�Թ��һ�������� -1 WHERE ������ = " & "'" & ������ź & "'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        ������ź = ""
        
        '�ѹ�֡�Թ��� ��ѧ�Ѻ��������� ����ź�ӹǹ
        SQL = "UPDATE �Ѻ��������� SET �ӹǹ�Թ��� = �ӹǹ�Թ��� -1 WHERE ���ʫѺ��������� = " & �Ѻ���������
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
  
        '�Դ���ҧ�����ա����
        SQL = "SELECT * FROM �Թ���  ORDER BY �ѹ������� ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
         '��ͧ�ѹ��û�͹��������ѧ DataGrid
        DataGrid1.AllowUpdate = False
        �Ѻ�Թ���.Caption = "���Թ�����к������� " & .RecordCount & "  ��¡�� "
        End With
        End If
End Sub
Private Sub �͡�����_Click()
    Unload Form1
End Sub

