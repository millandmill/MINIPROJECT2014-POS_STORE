VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "จัดการสินค้าของระบบ"
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
   Begin VB.CommandButton เคลียร์ 
      Caption         =   "เคลียร์หน้าจอ"
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
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3120
      TabIndex        =   25
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
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
   Begin VB.ComboBox ประเภท 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "โปรดเลือกประเภท"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton เพิ่มบันทึก 
      Caption         =   "เพิ่ม/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":2D12
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton แก้ไขบันทึก 
      Caption         =   "แก้ไข/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":39DC
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton ลบข้อมูล 
      Caption         =   "ลบข้อมูล"
      Height          =   855
      Left            =   0
      Picture         =   "addgoods.frx":66D6
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
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
      Caption         =   "วัน / เดือน / ปี ค.ศ."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   36
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "วัน / เดือน / ปี ค.ศ."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "บาท"
      Height          =   255
      Left            =   5040
      TabIndex        =   34
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "ราคาสินค้า"
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "ชื่อสินค้า"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label นับสินค้า 
      Caption         =   "นับจำนวนสินค้า"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label12 
      Caption         =   "สินค้าชิ้นนี้สั่งจากซัพพลายเออร์รหัส"
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "ชี้นให้แจ้งเตือน"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "ถ้าจำนวนสินค้าน้อยกว่า"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "ชี้น"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "จำนวนสินค้าที่มีอยู่คลัง"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "รหัส barcode ต้องมี 13 หลักเท่านั้น"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "รหัสbarcode"
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
      Caption         =   "วันหมดอายุ"
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
      Caption         =   "วันผลิต"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ประเภทสินค้า"
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
Dim namep, ประเภทลบ, ซับพลายเออร์ As String
Dim flag As Integer
Dim d_m, m_m, y_m, d_o, m_o, y_o, number_goods, del, aa, i As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
Private Sub id_Change()
On Error GoTo error1
If IsNumeric(id.text) = False Then
        MsgBox "กรุณาใส่เดือนที่ผลิตสินค้าให้ถูกต้อง ใส่เป็นตัวเลข 13 หลัก", vbInformation, "คำเตือน"
error1:
        id.text = ""
End If
End Sub
Private Sub Form_Load()

'วันผลิต
For i = 1 To 31
d.AddItem i
Next
d = 1

'เดือนผลิต
For i = 1 To 12
m.AddItem i
Next
m = 1

'ปีผลิต
For i = Year(Now()) - 10 To Year(Now()) + 10
y.AddItem i
Next
y = Year(Now())

'วันหมดอายุ
For i = 1 To 31
dd.AddItem i
Next
dd = 1

'เดือนหมดอายุ
For i = 1 To 12
mm.AddItem i
Next
mm = 1

'ปีหมดอายุ
For i = Year(Now()) - 10 To Year(Now()) + 10
yy.AddItem i
Next
yy = Year(Now())

'เปิดฐานข้อมูล
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With

'เปิดตาราง ประเภท เพื่อเอา ประเภท ใส่ Combo box
With RC
        SQL = "SELECT * FROM ประเภท"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Do While Not RC.EOF
            ประเภท.AddItem .Fields("ประเภท")
            RC.MoveNext
        Loop
        RC.Close
End With

'เปิดตาราง ซับพลายเออร์ เพื่อเอา ID ใส่ Combo box
With RC
        SQL = "SELECT * FROM ซับพลายเออร์"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Do While Not RC.EOF
            sup.AddItem .Fields("รหัสซับพลายเออร์")
            RC.MoveNext
        Loop
        RC.Close
End With

'เปิดตาราง
With RC
        SQL = "SELECT * FROM สินค้า  ORDER BY วันหมดอายุ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        นับสินค้า.Caption = "มีสินค้าในระบบทั้งหมด " & .RecordCount & "  รายการ "
End With
End Sub

Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
            sup.Enabled = False
            ประเภท.Enabled = False
            เพิ่มบันทึก.Enabled = False
            id.text = .Fields("barcode")
            nameg.text = .Fields("ชื่อสินค้า")
            ประเภท.text = .Fields("ประเภท")
            'วันผลิต
            d.text = Day(.Fields("วันผลิต"))
            m.text = Month(.Fields("วันผลิต"))
            y.text = Year(.Fields("วันผลิต"))
            'วันหมดอายุ
            dd.text = Day(.Fields("วันหมดอายุ"))
            mm.text = Month(.Fields("วันหมดอายุ"))
            yy.text = Year(.Fields("วันหมดอายุ"))
        
            numgood.text = .Fields("จำนวนสินค้า")
            lessgood.text = .Fields("เหลือเตือน")
            p.text = .Fields("ราคา")
            sup.text = .Fields("รหัสซับพลายเออร์")
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
            sup.Enabled = False
            ประเภท.Enabled = False
            เพิ่มบันทึก.Enabled = False
            id.text = .Fields("barcode")
            nameg.text = .Fields("ชื่อสินค้า")
            ประเภท.text = .Fields("ประเภท")
            'วันผลิต
            d.text = Day(.Fields("วันผลิต"))
            m.text = Month(.Fields("วันผลิต"))
            y.text = Year(.Fields("วันผลิต"))
            'วันหมดอายุ
            dd.text = Day(.Fields("วันหมดอายุ"))
            mm.text = Month(.Fields("วันหมดอายุ"))
            yy.text = Year(.Fields("วันหมดอายุ"))
        
            numgood.text = .Fields("จำนวนสินค้า")
            lessgood.text = .Fields("เหลือเตือน")
            p.text = .Fields("ราคา")
            sup.text = .Fields("รหัสซับพลายเออร์")
End If
End With
End Sub


Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
            sup.Enabled = False
            ประเภท.Enabled = False
            เพิ่มบันทึก.Enabled = False
           id.text = .Fields("barcode")
            nameg.text = .Fields("ชื่อสินค้า")
            ประเภท.text = .Fields("ประเภท")
            'วันผลิต
            d.text = Day(.Fields("วันผลิต"))
            m.text = Month(.Fields("วันผลิต"))
            y.text = Year(.Fields("วันผลิต"))
            'วันหมดอายุ
            dd.text = Day(.Fields("วันหมดอายุ"))
            mm.text = Month(.Fields("วันหมดอายุ"))
            yy.text = Year(.Fields("วันหมดอายุ"))
        
            numgood.text = .Fields("จำนวนสินค้า")
            lessgood.text = .Fields("เหลือเตือน")
            p.text = .Fields("ราคา")
            sup.text = .Fields("รหัสซับพลายเออร์")
End If
End With
End Sub

Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
            sup.Enabled = False
            ประเภท.Enabled = False
            เพิ่มบันทึก.Enabled = False
           id.text = .Fields("barcode")
            nameg.text = .Fields("ชื่อสินค้า")
            ประเภท.text = .Fields("ประเภท")
            'วันผลิต
            d.text = Day(.Fields("วันผลิต"))
            m.text = Month(.Fields("วันผลิต"))
            y.text = Year(.Fields("วันผลิต"))
            'วันหมดอายุ
            dd.text = Day(.Fields("วันหมดอายุ"))
            mm.text = Month(.Fields("วันหมดอายุ"))
            yy.text = Year(.Fields("วันหมดอายุ"))
        
            numgood.text = .Fields("จำนวนสินค้า")
            lessgood.text = .Fields("เหลือเตือน")
            p.text = .Fields("ราคา")
            sup.text = .Fields("รหัสซับพลายเออร์")
End If
End With

End Sub

Public Sub เคลียร์_Click()
            sup.Enabled = True
            ประเภท.Enabled = True
            เพิ่มบันทึก.Enabled = True
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

Private Sub เพิ่มบันทึก_Click()
With RC
Dim flag As Integer
Dim d1, d2 As Date

            flag = 0
            If (ประเภท.text = "") Then
                MsgBox "กรุณาใส่ข้อมูลประเภท"
                flag = 1
            End If
            
            If (sup.text = "") Then
                MsgBox "กรุณาใส่รหัสซัพพลายเออร์"
                flag = 1
            End If
            
            If (nameg.text = "") Then
                MsgBox "กรุณาใส่ชื่อสินค้า"
                flag = 1
             End If
             
            If (id.text = "") Or Len(id.text) < 13 Then
                MsgBox "กรุณาใส่รหัส barcode เป็นตัวเลขให้ครบ 13 ตัว"
                flag = 1
            End If
            
            On Error GoTo e2
            If (p.text <= 0) Then
e2:
                MsgBox "กรุณาใส่ราคาสินค้ามากกว่า 0 บาท"
                flag = 1
            End If
            
            'เช็ด ว / ด/ ป
            If (yy.text < y.text) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text < m.text)) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text = m.text) And (dd.text < d.text)) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            
            
           If (flag = 0) Then
            .AddNew
            .Fields("barcode") = id.text
            .Fields("ชื่อสินค้า") = nameg.text
            .Fields("ประเภท") = ประเภท.text
            'วันผลิต
            d_m = d.text
            m_m = m.text
            y_m = y.text
            .Fields("วันผลิต") = DateSerial(y_m, m_m, d_m)
            'วันหมดอายุ
            d_o = dd.text
            m_o = mm.text
            y_o = yy.text
            .Fields("วันหมดอายุ") = DateSerial(y_o, m_o, d_o)
            
            .Fields("จำนวนสินค้า") = numgood.text
            .Fields("เหลือเตือน") = lessgood.text
             .Fields("ราคา") = p.text
            .Fields("รหัสซับพลายเออร์") = sup.text
        On Error GoTo error1
        .Update
        MsgBox "เพิ่มข้อมูลสินค้าใหม่เรียบร้อย", vbInformation, "เพิ่มข้อมูลสินค้า"
        End If
End With

        'บันทึกสินค้า ไปยังประเภท เพื่อบันทึกจำนวน
        
With RC
        If (flag = 0) Then
        SQL = "UPDATE ประเภท SET จำนวนสินค้าประเภทนี้ = จำนวนสินค้าประเภทนี้ +1 WHERE ประเภท = " & "'" & ประเภท.text & "'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        End If
        
        'บันทึกสินค้า ไปยังซับพลายเออร์ เพื่อบันทึกจำนวน
        If (flag = 0) Then
        SQL = "UPDATE ซับพลายเออร์ SET จำนวนสินค้า = จำนวนสินค้า +1 WHERE รหัสซับพลายเออร์ = " & sup.text
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        End If
        
        'เปิดตารางใหม่อีกครั้ง
        SQL = "SELECT * FROM สินค้า  ORDER BY วันหมดอายุ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        นับสินค้า.Caption = "มีสินค้าในระบบทั้งหมด " & .RecordCount & "  รายการ "
End With
Call Command1_Click
        
If Error = 0 Then
error1:
        flag = 1
        MsgBox "กรุณาบันทึกสินค้าใหม่อีกครั้ง อาจจะเป็นเพราะสินค้าชี้นนี้อยู่ในระบบแล้ว หรือ การกรอกข้อมูลผิดรูปแบบ", vbInformation, "เพิ่มข้อมูลสินค้าผิดพลาด"
         Call Form_Load
         Unload Form1
         Form1.Show
End If

End Sub
Private Sub แก้ไขบันทึก_Click()
With RC
Dim flag As Integer
flag = 0
If .RecordCount = 0 Then
    MsgBox "ไม่สามารถแก้ไข/บันทึกได้เนื่องจาก ไม่มีสินค้าใหแก้ไข้", vbInformation, "แก้ไข/บันทึก"
    flag = 1
End If
            If (ประเภท.text = "") Then
                MsgBox "กรุณาใส่ข้อมูลประเภท"
                flag = 1
            End If
            
            If (sup.text = "") Then
                MsgBox "กรุณาใส่รหัสซัพพลายเออร์"
                flag = 1
            End If
            
            If (nameg.text = "") Then
                MsgBox "กรุณาใส่ชื่อสินค้า"
                flag = 1
             End If
             
             
             If (id.text = "") Or Len(id.text) < 13 Then
                MsgBox "กรุณาใส่รหัส barcode เป็นตัวเลขให้ครบ 13 ตัว"
                flag = 1
            End If
            
            On Error GoTo e2
            If (p.text <= 0) Then
e2:
                MsgBox "กรุณาใส่ราคาสินค้ามากกว่า 0 บาท"
                flag = 1
            End If
            
            
             'เช็ด ว / ด/ ป
            If (yy.text < y.text) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text < m.text)) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            If ((yy.text = y.text) And (mm.text = m.text) And (dd.text < d.text)) Then
                MsgBox "กรุณาใสวันหมดอายุของสินค้าให้ถูกต้อง คือ วันหมดอายุต้องไม่หมดก่อน วันผลิต"
                flag = 1
            End If
            
           If (flag = 0) Then
            .Fields("barcode") = id.text
            .Fields("ชื่อสินค้า") = nameg.text
            .Fields("ประเภท") = ประเภท.text
            'วันผลิต
            d_m = d.text
            m_m = m.text
            y_m = y.text
            .Fields("วันผลิต") = DateSerial(y_m, m_m, d_m)
            'วันหมดอายุ
            d_o = dd.text
            m_o = mm.text
            y_o = yy.text
            .Fields("วันหมดอายุ") = DateSerial(y_o, m_o, d_o)
            
            .Fields("จำนวนสินค้า") = numgood.text
            .Fields("เหลือเตือน") = lessgood.text
            .Fields("ราคา") = p.text
            .Fields("รหัสซับพลายเออร์") = sup.text
            On Error GoTo error1
           .Update
           Call Command1_Click
            MsgBox "แก้ไข/บันทึกสินค้าให้เรียบร้อยแล้ว", vbInformation, "แก้ไข/บันทึก"
            นับสินค้า.Caption = "มีสินค้าในระบบทั้งหมด " & .RecordCount & "  รายการ "
End If
End With
If Error = 0 Then
error1:
        MsgBox "กรุณาบันทึกสินค้าใหม่อีกครั้ง อาจจะเป็นเพราะสินค้าชี้นนี้อยู่ในระบบแล้ว หรือ การกรอกข้อมูลผิดรูปแบบ", vbInformation, "เพิ่มข้อมูลสินค้าผิดพลาด"
         Call Form_Load
         Unload Form1
         Form1.Show
End If
End Sub
Private Sub ลบข้อมูล_Click()
Dim flag As Integer
flag = 0
With RC
If .RecordCount > 0 Then
        flag = 1
        namep = .Fields("ชื่อสินค้า")
        ประเภทลบ = .Fields("ประเภท")
        ซับพลายเออร์ = .Fields("รหัสซับพลายเออร์")
        .Delete
        .Requery
        MsgBox "ลบข้อมูลหนังสือชื่อ " & namep & " เรียบร้อยแล้ว", vbInformation, "ลบข้อมูลหนังสือ"
Else
        MsgBox "ไม่สามารถลบข้อมูลหนังสือได้เนื่องจากไม่มีหนังสือให้ลบ", vbInformation, "ลบข้อมูลหนังสือ"
End If
End With
       'บันทึกสินค้า ไปยังประเภท เพื่อลบจำนวน
       If (flag = 1) Then
       flag = 0
        With RC
        SQL = "UPDATE ประเภท SET จำนวนสินค้าประเภทนี้ = จำนวนสินค้าประเภทนี้ -1 WHERE ประเภท = " & "'" & ประเภทลบ & "'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        ประเภทลบ = ""
        
        'บันทึกสินค้า ไปยังซับพลายเออร์ เพื่อลบจำนวน
        SQL = "UPDATE ซับพลายเออร์ SET จำนวนสินค้า = จำนวนสินค้า -1 WHERE รหัสซับพลายเออร์ = " & ซับพลายเออร์
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
  
        'เปิดตารางใหม่อีกครั้ง
        SQL = "SELECT * FROM สินค้า  ORDER BY วันหมดอายุ ASC"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
         'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        นับสินค้า.Caption = "มีสินค้าในระบบทั้งหมด " & .RecordCount & "  รายการ "
        End With
        End If
End Sub
Private Sub ออกโปรแกรม_Click()
    Unload Form1
End Sub

