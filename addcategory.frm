VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "จัดการประเภทสินค้า"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13365
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   13365
   Begin VB.CommandButton Command5 
      Caption         =   "เคลียร์หน้าจอ"
      Height          =   855
      Left            =   0
      Picture         =   "addcategory.frx":0000
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton เคลียร์ 
      Caption         =   "เคลียร์"
      Height          =   855
      Left            =   -2640
      Picture         =   "addcategory.frx":2CFA
      TabIndex        =   14
      Top             =   -120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox co 
      Height          =   1095
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox namec 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
      Height          =   855
      Left            =   0
      Picture         =   "addcategory.frx":59F4
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton ลบข้อมูล 
      Caption         =   "ลบข้อมูล"
      Height          =   855
      Left            =   0
      Picture         =   "addcategory.frx":8376
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton แก้ไขบันทึก 
      Caption         =   "แก้ไข/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addcategory.frx":ACF8
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton เพิ่ม 
      Caption         =   "เพิ่ม/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addcategory.frx":D9F2
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "addcategory.frx":E6BC
      Height          =   2655
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
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
   Begin VB.Label นับ 
      Caption         =   "นับประเภทสินค้า"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label4 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "คำอธิบาย"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ชื่อประเภทสินค้า"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
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
        SQL = "SELECT * FROM ประเภท  ORDER BY จำนวนสินค้าประเภทนี้ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End With
End Sub

Private Sub เพิ่มบันทึก_Click()
With RC
            .AddNew
            .Fields("ประเภท") = namec.text
            .Fields("คำอธิบาย") = co.text
        .Update
        MsgBox "เพิ่มประเภทสินค้าชนิดใหม่เรียบร้อยแล้ว", vbInformation, "เพิ่มประเภทสินค้า"
        นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End With
End Sub


Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
            If (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then
            namec.Enabled = False
            co.Enabled = False
            แก้ไขบันทึก.Enabled = False
            Label4.Caption = "ไม่สามารถแก้ไขชื่อประเภทสินค้าได้เนื่องจาก ในประเภทสินค้านี้ มีรายการสินค้ามากกว่า 0 รายการ"
            End If
            
            If (.Fields("จำนวนสินค้าประเภทนี้") = 0) Then
            namec.Enabled = True
            co.Enabled = True
            แก้ไขบันทึก.Enabled = True
            Label4.Caption = ""
            End If
            
            namec.text = .Fields("ประเภท")
            co.text = .Fields("คำอธิบาย")
       'นับจำนวนสินค้าในประเภท
       Label3.Caption = "ประเภทสินค้านี้มีสินค้าทั้งหมด " & .Fields("จำนวนสินค้าประเภทนี้") & " รายการ"
       นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
            If (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then
            namec.Enabled = False
            co.Enabled = False
            แก้ไขบันทึก.Enabled = False
            Label4.Caption = "ไม่สามารถแก้ไขชื่อประเภทสินค้าได้เนื่องจาก ในประเภทสินค้านี้ มีรายการสินค้ามากกว่า 0 รายการ"
            End If
            
            If (.Fields("จำนวนสินค้าประเภทนี้") = 0) Then
            namec.Enabled = True
            co.Enabled = True
            แก้ไขบันทึก.Enabled = True
            End If
            
            namec.text = .Fields("ประเภท")
            co.text = .Fields("คำอธิบาย")
       'นับจำนวนสินค้าในประเภท
       Label3.Caption = "ประเภทสินค้านี้มีสินค้าทั้งหมด " & .Fields("จำนวนสินค้าประเภทนี้") & " รายการ"
       นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End If
End With
End Sub
Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
            If (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then
            namec.Enabled = False
            co.Enabled = False
            แก้ไขบันทึก.Enabled = False
            Label4.Caption = "ไม่สามารถแก้ไขชื่อประเภทสินค้าได้เนื่องจาก ในประเภทสินค้านี้ มีรายการสินค้ามากกว่า 0 รายการ"
            End If
            
            If (.Fields("จำนวนสินค้าประเภทนี้") = 0) Then
            namec.Enabled = True
            co.Enabled = True
            แก้ไขบันทึก.Enabled = True
            Label4.Caption = ""
            End If
            
            namec.text = .Fields("ประเภท")
            co.text = .Fields("คำอธิบาย")
           'นับจำนวนสินค้าในประเภท
       Label3.Caption = "ประเภทสินค้านี้มีสินค้าทั้งหมด " & .Fields("จำนวนสินค้าประเภทนี้") & " รายการ"
       นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End If
End With
End Sub

Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
            If (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then
            namec.Enabled = False
            co.Enabled = False
            แก้ไขบันทึก.Enabled = False
            Label4.Caption = "ไม่สามารถแก้ไขชื่อประเภทสินค้าได้เนื่องจาก ในประเภทสินค้านี้ มีรายการสินค้ามากกว่า 0 รายการ"
            End If
            
            If (.Fields("จำนวนสินค้าประเภทนี้") = 0) Then
            namec.Enabled = True
            co.Enabled = True
            แก้ไขบันทึก.Enabled = True
            Label4.Caption = ""
            End If
            
            namec.text = .Fields("ประเภท")
            co.text = .Fields("คำอธิบาย")
                   'นับจำนวนสินค้าในประเภท
       Label3.Caption = "ประเภทสินค้านี้มีสินค้าทั้งหมด " & .Fields("จำนวนสินค้าประเภทนี้") & " รายการ"
       นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
End If
End With
End Sub

Private Sub namec_Change()
If (flag <> 1) And ((namec.text = "") Or (namec.text = " ")) Then
        MsgBox "กรุณาป้อนชื่อประเภทสินค้า", vbInformation, "เพิ่มข้อมูลประเภทของสินค้า"
        flag = 0
End If
End Sub

Private Sub เพิ่ม_Click()
With RC
            If (namec.text = "") Or (namec.text = " ") Then
            MsgBox "กรุณาป้อนชื่อประเภทสินค้า"
            Else
            On Error GoTo error1
            .AddNew
            .Fields("ประเภท") = namec.text
            .Fields("คำอธิบาย") = co.text
            .Update
            นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
            End If
End With
If Error = 0 Then
error1:
        MsgBox "กรุณาบันทึกประเภทของสินค้าใหม่อีกครั้ง อาจจะเป็นเพราะประเภทของสินค้าชี้นนี้อยู่ในระบบแล้ว หรือ การกรอกข้อมูลผิดรูปแบบ", vbInformation, "เพิ่มข้อมูลประเภทของสินค้าผิดพลาด"
         Call Form_Load
End If
End Sub

Private Sub แก้ไขบันทึก_Click()
With RC
If .RecordCount = 0 Then
    MsgBox "ไม่สามารถแก้ไข/บันทึกได้เนื่องจาก ไม่มีประเภทของสินค้าใหแก้ไข้", vbInformation, "แก้ไข/บันทึก"
ElseIf (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then
    MsgBox "ไม่สามารถแก้ไข/บันทึกได้เนื่องจากยังมีสินค้าที่เกี่ยวข้องกับประเภทสินค้านี้อยู่"
ElseIf (.RecordCount > 0) Then
            .Fields("ประเภท") = namec.text
            .Fields("คำอธิบาย") = co.text
            .Update
            นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
    MsgBox "แก้ไขประเภทสินค้าเรียบร้อยแล้ว", vbInformation, "แก้ไข/บันทึก"
End If
End With
End Sub
'เคลียร์
Private Sub Command5_Click()
    flag = 1
    namec.text = ""
    flag = 0
    co.text = ""
    Label4.Caption = ""
    namec.Enabled = True
    Label3.Caption = ""
    Call Form_Load
End Sub
Private Sub ลบข้อมูล_Click()
With RC
On Error GoTo e1
If .RecordCount > 0 Or (.Fields("จำนวนสินค้าประเภทนี้") = 0) Then
        Dim cname As String
        cname = .Fields("ประเภท")
        .Delete
        .Requery
        นับ.Caption = "มีประเภทสินค้าอยู่ในระบบทั้งหมด " & .RecordCount & " ประเภท"
        MsgBox "ลบข้อมูลประเภทของสินค้าชื่อ " & cname & " เรียบร้อยแล้ว"
Else
        If (.Fields("จำนวนสินค้าประเภทนี้") > 0) Then MsgBox "ไม่สามารถแก้ไข/บันทึกได้เนื่องจากยังมีสินค้าที่เกี่ยวข้องกับประเภทสินค้านี้อยู่"
e1:
        If (.RecordCount = 0) Then MsgBox "ไม่สามารถลบข้อมูลประเภทของสินค้าได้เนื่องจากไม่มีรายการให้ลบ"
End If
End With
Call Command1_Click
End Sub
Private Sub ออกโปรแกรม_Click()
    Unload Form2
End Sub
