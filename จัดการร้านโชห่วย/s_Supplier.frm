VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ระบบค้นหาซัพพลายเออร์"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command13 
      Caption         =   "เปิดการค้นหาใหม่"
      Height          =   255
      Left            =   1800
      Picture         =   "s_Supplier.frx":0000
      TabIndex        =   25
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "ค้นหาจากสินค้าที่เกี่ยวข้อง"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
      Height          =   255
      Left            =   240
      Picture         =   "s_Supplier.frx":2982
      TabIndex        =   23
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   300
      Left            =   6240
      TabIndex        =   22
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   300
      Left            =   5760
      TabIndex        =   21
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   300
      Left            =   5280
      TabIndex        =   20
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   300
      Left            =   4800
      TabIndex        =   19
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   300
      Left            =   6240
      TabIndex        =   18
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   300
      Left            =   5760
      TabIndex        =   17
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      Height          =   300
      Left            =   5280
      TabIndex        =   16
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "|<"
      Height          =   300
      Left            =   4800
      TabIndex        =   15
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "สินค้าที่เกี่ยวข้องกับซัพพลายเออร์ที่เลือก"
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Width           =   6495
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2566
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
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "ผลการค้นหาซัพพลายเออร์ที่ตรงกับการค้นหา"
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   6495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2566
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "การค้นหา"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ค้นหาจากชื่อ"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ค้นหาจากรหัส"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "เคลียร์ผลการค้นหา"
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "ชื่อบริษัท/บุคคล"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "รหัสซัพพลายเออร์"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "รหัสซัพพลายเออร์"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "จากผลการค้นหาเลือกซัพพลายเออร์เพื่อดูความเกี่ยวข้องกับสินค้าของระบบ"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   5175
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Dim flag As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"

Private Sub Command3_Click()
Text1.text = ""
Text2.text = ""
Call Form_Load
End Sub

Private Sub Form_Load()
'เปิดฐานข้อมูล
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With
'เปิดตาราง
With RC
        SQL = "SELECT * FROM ซับพลายเออร์  ORDER BY รหัสซับพลายเออร์ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
End With
Frame2.Caption = "รายชื่อซัพพลายเออร์ในระบบทั้งหมด"
Combo1.Enabled = False
Command12.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
DataGrid2.Enabled = False
End Sub
Private Sub Command1_Click()
'ค้นหาจากชื่อบริษัท/บุคคล
If (Text1.text <> "") Then
Text2.text = ""
With RC
        SQL = "SELECT * FROM ซับพลายเออร์ WHERE ชื่อ  LIKE '%" & Text1.text & "%' ORDER BY ชื่อ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        
        'ใส่รหัสซับพลายเออร์ ใน combo1
        Do While Not RC.EOF
            Combo1.AddItem .Fields("รหัสซับพลายเออร์")
            RC.MoveNext
        Loop

        'เปิดผลลัพธ์เป้นแสดงใน data กริด อีกครั้ง
        SQL = "SELECT * FROM ซับพลายเออร์ WHERE ชื่อ  LIKE '%" & Text1.text & "%' ORDER BY ชื่อ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        Label5.Caption = "มีผลการค้นหาที่ตรงกับเงื่อนไขทั้งหมด " & .RecordCount & " รายการ"
        If .RecordCount >= 1 Then
        Combo1.Enabled = True
        Command12.Enabled = True
        DataGrid2.Enabled = True
        Command8.Enabled = True
        Command9.Enabled = True
        Command10.Enabled = True
        Command11.Enabled = True
        End If
End With
Else
    MsgBox "กรุณาใส่คำค้นหาชื่อบริษัท/บุคคล"
End If
End Sub
Private Sub Command2_Click()
'ค้นหาจากบาร์โค้ด
If IsNumeric(Text2.text) = True Then
Text1.text = ""
With RC
        SQL = "SELECT * FROM ซับพลายเออร์ WHERE รหัสซับพลายเออร์ = " & Text2.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
  
        'ใส่รหัสซับพลายเออร์ ใน combo1
        Do While Not RC.EOF
            Combo1.AddItem .Fields("รหัสซับพลายเออร์")
            RC.MoveNext
        Loop
        
        'เปิดผลลัพธ์เป้นแสดงใน data กริด อีกครั้ง
        SQL = "SELECT * FROM ซับพลายเออร์ WHERE รหัสซับพลายเออร์ = " & Text2.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        Label5.Caption = "มีผลการค้นหาที่ตรงกับเงื่อนไขทั้งหมด " & .RecordCount & " รายการ"
        If .RecordCount >= 1 Then
        Combo1.Enabled = True
        Command12.Enabled = True
        DataGrid2.Enabled = True
        Command8.Enabled = True
        Command9.Enabled = True
        Command10.Enabled = True
        Command11.Enabled = True
        End If
End With
Else
        MsgBox "กรุณาใส่คำค้น บาร์โค้ด เป็นตัวเลข"
        Text2.text = ""
End If
End Sub



Public Sub Command7_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
End If
End With
End Sub

Private Sub Command6_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
End If
End With
End Sub
Private Sub Command5_Click()
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

Private Sub Command12_Click()
If (Combo1) <> "" Then
DataGrid1.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Label5.Caption = ""
'ค้นหาสินค้าที่เกี่ยวข้องจากรหัส
With RC
        SQL = "SELECT * FROM สินค้า WHERE รหัสซับพลายเออร์ =" & Combo1.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid2.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid2.AllowUpdate = False
        Label6.Caption = "มีสินค้าที่เกี่ยวข้องกับรหัสซัพพลายเออร์ " & Combo1.text & " ทั้งหมด " & .RecordCount & " รายการ"
End With
End If
End Sub
Public Sub Command8_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
End If
End With
End Sub

Private Sub Command9_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
End If
End With
End Sub
Private Sub Command10_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
End If
End With
End Sub
Private Sub Command11_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
End If
End With
End Sub
Private Sub ออกโปรแกรม_Click()
    Unload Form12
End Sub
Private Sub Command13_Click()
    Unload Form12
    Form12.Show
End Sub
