VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "จัดการซัพพลายเออร์"
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
   Begin VB.CommandButton เคลียร์ 
      Caption         =   "เคลียร์หน้าจอ"
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
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      ToolTipText     =   "ไปยังข้อมูลแรก"
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
      Caption         =   "ที่อยู่"
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
         Caption         =   "รหัสไปรษณีย์"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "จังหวัด"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "ที่อยู่"
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
   Begin VB.CommandButton เพิ่มบันทึก 
      Caption         =   "เพิ่ม/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":2CFE
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton แก้ไขบันทึก 
      Caption         =   "แก้ไข/บันทึก"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":39C8
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton ลบข้อมูล 
      Caption         =   "ลบข้อมูล"
      Height          =   855
      Left            =   0
      Picture         =   "addSupplier.frx":66C2
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
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
   Begin VB.Label นับ 
      Caption         =   "นับซัพพลายเออร์"
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
      Caption         =   "หมายเหตุ"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "เบอร์โทรศัพท์ประจำที่"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "เบอร์โทรศัพท์มือถือ"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "ชื่อบริษัท/บุคคล"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "รหัสซัพพลายเออร์ :"
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
        นับ.Caption = "ซัพพลายเออร์มีทั้งหมด " & .RecordCount & " ราย"
End With

prov.AddItem "กระบี่"
prov.AddItem "กรุงเทพมหานคร"
prov.AddItem "กาญจนบุรี"
prov.AddItem "กาฬสินธุ์"
prov.AddItem "กำแพงเพชร"
prov.AddItem "ขอนแก่น"
prov.AddItem "จันทบุรี"
prov.AddItem "ฉะเชิงเทรา"
prov.AddItem "ชลบุรี"
prov.AddItem "ชัยนาท"
prov.AddItem "ชัยภูมิ"
prov.AddItem "ชุมพร"
prov.AddItem "เชียงราย"
prov.AddItem "เชียงใหม่"
prov.AddItem "ตรัง"
prov.AddItem "ตราด"
prov.AddItem "ตาก"
prov.AddItem "นครนายก"
prov.AddItem "นครปฐม"
prov.AddItem "นครพนม"
prov.AddItem "นครราชสีมา"
prov.AddItem "นครศรีธรรมราช"
prov.AddItem "นครสวรรค์"
prov.AddItem "นนทบุรี"
prov.AddItem "นราธิวาส"
prov.AddItem "น่าน"
prov.AddItem "บุรีรัมย์"
prov.AddItem "ปทุมธานี"
prov.AddItem "ประจวบคีรีขันธ์"
prov.AddItem "ปราจีนบุรี"
prov.AddItem "ปัตตานี"
prov.AddItem "พระนครศรีอยุธยา"
prov.AddItem "พะเยา"
prov.AddItem "พังงา"
prov.AddItem "พัทลุง"
prov.AddItem "พิจิตร"
prov.AddItem "พิษณุโลก"
prov.AddItem "เพชรบุรี"
prov.AddItem "เพชรบูรณ์"
prov.AddItem "แพร่"
prov.AddItem "ภูเก็ต"
prov.AddItem "มหาสารคาม"
prov.AddItem "มุกดาหาร"
prov.AddItem "แม่ฮ่องสอน"
prov.AddItem "ยโสธร"
prov.AddItem "ยะลา"
prov.AddItem "ร้อยเอ็ด"
prov.AddItem "ระนอง"
prov.AddItem "ระยอง"
prov.AddItem "ราชบุรี"
prov.AddItem "ลพบุรี"
prov.AddItem "เลย"
prov.AddItem "ลำปาง"
prov.AddItem "ลำพูน"
prov.AddItem "ศีรสะเกษ"
prov.AddItem "สกลนคร"
prov.AddItem "สงขลา"
prov.AddItem "สตูล"
prov.AddItem "สมุทรปราการ"
prov.AddItem "สมุทรสงคราม"
prov.AddItem "สมุทรสาคร"
prov.AddItem "สระแก้ว"
prov.AddItem "สระบุรี"
prov.AddItem "สิงห์บุรี"
prov.AddItem "สุโขทัย"
prov.AddItem "สุพรรณบุรี"
prov.AddItem "สุราษฎร์ธานี"
prov.AddItem "สุรินทร์"
prov.AddItem "หนองคาย"
prov.AddItem "หนองบัวลำภู"
prov.AddItem "อ่างทอง"
prov.AddItem "อำนาจเจริญ"
prov.AddItem "อุดรธานี"
prov.AddItem "อุตรดิตถ์"
prov.AddItem "อุทัยธานี"
prov.AddItem "อุบลราชธานี"

prov.text = "กรุงเทพมหานคร"
End Sub

Private Sub zip_Change()
    If (IsNumeric(zip.text) = False) Then
        MsgBox ("กรุณาใส่รหัสไปรษณีย์เป็นตัวเลข จำนวน 5 หลัก")
        zip.text = "00000"
    End If
End Sub
Private Sub เคลียร์_Click()
            Label1.Caption = "รหัสซัพพลายเออร์ : "
            namep.text = ""
            addr.text = ""
            zip.text = "00000"
            mobile.text = ""
            phone.text = ""
            email.text = ""
            com.text = ""
End Sub
Private Sub แก้ไขบันทึก_Click()
With RC
If .RecordCount = 0 Then
    MsgBox "ไม่สามารถแก้ไข/บันทึกได้เนื่องจาก ไม่มีซับพลายเออร์ให้แก้ไข้", vbInformation, "แก้ไข/บันทึก"
ElseIf (namep.text <> "") And (addr.text <> "") And Len(zip.text) = 5 Then
            .Fields("ชื่อ") = namep.text
            .Fields("ที่อยู่") = addr.text
            .Fields("จังหวัด") = prov.text
            .Fields("รหัสไปรษณีย์") = zip.text
            .Fields("โทรศัพท์มือถือ") = mobile.text
            .Fields("โทรศัพท์") = phone.text
            .Fields("ที่อยู่อีเมล") = email.text
            .Fields("หมายเหตุ") = com.text
           .Update
            MsgBox "แก้ไข/บันทึกข้อมูลซับพลายเออร์ให้เรียบร้อยแล้ว", vbInformation, "แก้ไข/บันทึก"
Else
            If (Len(zip.text) < 5) Then MsgBox "กรุณาใส่รหัสไปรษณีย์ 5 หลัก"
            MsgBox "กรุณาเพิ่มชื่อบริษัท/บุคคล ที่อยู่ รหัสไปรษณีย์  เบอร์โทรศัพท์ให้เรียบร้อย", vbInformation, "แก้ไข/บันทึก"
End If
End With
If Error = 1 Then
error5:
    Unload Form3
    Form3.Show
End If
End Sub
Private Sub เพิ่มบันทึก_Click()
With RC
If (namep.text <> "") And (addr.text <> "") And Len(zip.text) = 5 Then
            On Error GoTo error1
            .AddNew
            .Fields("ชื่อ") = namep.text
            .Fields("ที่อยู่") = addr.text
            .Fields("จังหวัด") = prov.text
            .Fields("รหัสไปรษณีย์") = zip.text
            .Fields("โทรศัพท์มือถือ") = mobile.text
            .Fields("โทรศัพท์") = phone.text
            .Fields("ที่อยู่อีเมล") = email.text
            .Fields("หมายเหตุ") = com.text
            .Fields("จำนวนสินค้า") = 0
        .Update
        นับ.Caption = "ซัพพลายเออร์มีทั้งหมด " & .RecordCount & " ราย"
        MsgBox "เพิ่มข้อมูลซับพลายเออร์ใหม่เรียบร้อย", vbInformation, "เพิ่มข้อมูลซับพลายเออร์"
        Call Command1_Click
Else
            If (Len(zip.text) < 5) Then MsgBox "กรุณาใส่รหัสไปรษณีย์ 5 หลัก"
            MsgBox "กรุณาเพิ่มชื่อบริษัท/บุคคล ที่อยู่ รหัสไปรษณีย์  เบอร์โทรศัพท์ให้เรียบร้อย", vbInformation, "แก้ไข/บันทึก"
End If
End With
If Error = 1 Then
error1:
        MsgBox "กรุณาบันทึกข้อมูลซัพพลายเออร์อีกครั้ง อาจจะเป็นเพราะมีข้อมูลนี้อยู่ในระบบแล้ว หรือ การกรอกข้อมูลผิดรูปแบบ", vbInformation, "เพิ่มข้อมูลสินค้าผิดพลาด"
         MsgBox "ระบบจะทำการปิดระบบจัดการซัพพลายเออร์ ถ้าผู้ใช้ต้องการเพิ่มข้อมูลซัพพลายเออร์รายใหม่ กรุณาเข้าระบบจัดการซัพพลายเออร์อีกครั้ง", vbInformation, "ข้อความจากระบบ"
         Call Form_Load
         Unload Form3
End If
End Sub

Public Sub Command1_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
            Label1.Caption = "รหัสซัพพลายเออร์ : " & .Fields("รหัสซับพลายเออร์")
            namep.text = .Fields("ชื่อ")
            addr.text = .Fields("ที่อยู่")
            prov.text = .Fields("จังหวัด")
            zip.text = .Fields("รหัสไปรษณีย์")
            mobile.text = .Fields("โทรศัพท์มือถือ")
            phone.text = .Fields("โทรศัพท์")
            email.text = .Fields("ที่อยู่อีเมล")
            com.text = .Fields("หมายเหตุ")
End If
End With
End Sub

Private Sub Command2_Click()
With RC
If .RecordCount > 0 Then
            .MovePrevious
            If .BOF = True Then .MoveLast
            Label1.Caption = "รหัสซัพพลายเออร์ : " & .Fields("รหัสซับพลายเออร์")
            namep.text = .Fields("ชื่อ")
            addr.text = .Fields("ที่อยู่")
            prov.text = .Fields("จังหวัด")
            zip.text = .Fields("รหัสไปรษณีย์")
            mobile.text = .Fields("โทรศัพท์มือถือ")
            phone.text = .Fields("โทรศัพท์")
            email.text = .Fields("ที่อยู่อีเมล")
            com.text = .Fields("หมายเหตุ")
End If
End With
End Sub


Private Sub Command3_Click()
With RC
If .RecordCount > 0 Then
            .MoveNext
        If .EOF = True Then .MoveFirst
            Label1.Caption = "รหัสซัพพลายเออร์ : " & .Fields("รหัสซับพลายเออร์")
            namep.text = .Fields("ชื่อ")
            addr.text = .Fields("ที่อยู่")
            prov.text = .Fields("จังหวัด")
            zip.text = .Fields("รหัสไปรษณีย์")
            mobile.text = .Fields("โทรศัพท์มือถือ")
            phone.text = .Fields("โทรศัพท์")
            email.text = .Fields("ที่อยู่อีเมล")
            com.text = .Fields("หมายเหตุ")
End If
End With
End Sub

Private Sub Command4_Click()
With RC
If .RecordCount > 0 Then
            .MoveLast
            Label1.Caption = "รหัสซัพพลายเออร์ : " & .Fields("รหัสซับพลายเออร์")
            namep.text = .Fields("ชื่อ")
            addr.text = .Fields("ที่อยู่")
            prov.text = .Fields("จังหวัด")
            zip.text = .Fields("รหัสไปรษณีย์")
            mobile.text = .Fields("โทรศัพท์มือถือ")
            phone.text = .Fields("โทรศัพท์")
            email.text = .Fields("ที่อยู่อีเมล")
            com.text = .Fields("หมายเหตุ")
End If
End With
End Sub

Private Sub ลบข้อมูล_Click()
With RC
On Error GoTo del
If (.RecordCount > 0) And (.Fields("จำนวนสินค้า") = 0) Then
        Dim sid As String
        sid = .Fields("รหัสซับพลายเออร์")
        .Delete
        .Requery
        นับ.Caption = "ซัพพลายเออร์มีทั้งหมด " & .RecordCount & " ราย"
        MsgBox "ลบข้อมูลซับพลายเออร์รหัส " & sid & " เรียบร้อยแล้ว", vbInformation, "ลบข้อมูลซับพลายเออร์"
Else
del:
        MsgBox "ไม่สามารถลบข้อมูลซับพลายเออร์ได้เนื่องจากไม่มีซับพลายเออร์ให้ลบ หรือ ซับพลายเออร์ที่ต้องการจะลบยังมีสินค้าที่เกี่ยวข้องอยู่ ", vbInformation, "ลบข้อมูลซับพลายเออร์"
End If
End With
Call Command1_Click
End Sub
Private Sub ออกโปรแกรม_Click()
    Unload Form3
End Sub
