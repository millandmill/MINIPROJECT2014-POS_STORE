VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ระบบค้นหาสินค้า"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "|<"
      Height          =   300
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      Height          =   300
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   300
      Left            =   5640
      TabIndex        =   11
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   300
      Left            =   6120
      TabIndex        =   10
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
      Height          =   375
      Left            =   120
      Picture         =   "s_goods.frx":0000
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "การค้นหา"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "เคลียร์ผลการค้นหา"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ค้นหาจากบาร์โค้ด"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ค้นหาจากชื่อสินค้า"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ชื่อสินค้า"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
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
   Begin VB.Label Label4 
      Caption         =   "แสดงรายชื่อสินค้าที่มีอยู่ในร้านทั้งหมด"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   3840
      Width           =   2895
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Dim flag As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
Private Sub Form_Load()
'เปิดฐานข้อมูล
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With
'เปิดตาราง
With RC
        SQL = "SELECT barcode,ชื่อสินค้า,วันผลิต,วันหมดอายุ,ประเภท,จำนวนสินค้า,เหลือเตือน,รหัสซับพลายเออร์ FROM สินค้า  ORDER BY วันหมดอายุ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
End With
End Sub
Private Sub Command1_Click()
'ค้นหาจากชื่อสินค้า
If (Text1.text <> "") Then
With RC
        SQL = "SELECT barcode,ชื่อสินค้า,วันผลิต,วันหมดอายุ,ประเภท,จำนวนสินค้า,เหลือเตือน,รหัสซับพลายเออร์  FROM สินค้า  WHERE ชื่อสินค้า LIKE '" & Text1.text & "%' ORDER BY ชื่อสินค้า ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        Label3.Caption = "มีผลการค้นหาตรงกันหรือใกล้เคียงทั้งหมด " & .RecordCount & " รายการ"
        Label4.Caption = "ผลลัพธ์จากการค้นหาจากชื่อสินค้า"
End With
Else
    MsgBox "กรุณาใส่คำค้นหาชื่อสินค้า"
End If
End Sub
Private Sub Command2_Click()
'ค้นหาจากบาร์โค้ด
If IsNumeric(Text2.text) = True Then
With RC
        SQL = "SELECT barcode,ชื่อสินค้า,วันผลิต,วันหมดอายุ,ประเภท,จำนวนสินค้า,เหลือเตือน,รหัสซับพลายเออร์  FROM สินค้า  WHERE barcode LIKE '" & Text2.text & "%' ORDER BY barcode ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        Label3.Caption = "มีผลการค้นหาตรงกันหรือใกล้เคียงทั้งหมด " & .RecordCount & " รายการ"
        Label4.Caption = "ผลลัพธ์จากการค้นหาจาก barcode"
End With
Else
        MsgBox "กรุณาใส่คำค้น บาร์โค้ด เป็นตัวเลข"
        Text2.text = ""
End If
End Sub

Private Sub Text2_Change()
    If (IsNumeric(Text2.text) = False) And (flag <> 1) Then
    MsgBox "กรุณาใส่คำค้น บาร์โค้ด เป็นตัวเลข"
    Text2.text = ""
    End If
End Sub
'เคลียร์ผลการค้นหา
Private Sub Command3_Click()
flag = 1
    Text1.text = ""
    Text2.text = ""
    Label4.Caption = "แสดงรายชื่อสินค้าที่มีอยู่ในร้านทั้งหมด"
    Label3.Caption = ""
flag = 0
'แสดงสินค้าทั้งหมดในระบบ
With RC
        SQL = "SELECT barcode,ชื่อสินค้า,วันผลิต,วันหมดอายุ,ประเภท,จำนวนสินค้า,เหลือเตือน,รหัสซับพลายเออร์ FROM สินค้า  ORDER BY วันหมดอายุ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
End With
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
Private Sub ออกโปรแกรม_Click()
    Unload Form11
End Sub

