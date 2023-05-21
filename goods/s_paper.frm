VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "ระบบค้นหาใบเสร็จ"
   ClientHeight    =   5385
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   5385
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "ค้นหาใหม่อีกครั้ง"
      Height          =   375
      Left            =   9000
      TabIndex        =   52
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ออก"
      Height          =   375
      Left            =   1800
      TabIndex        =   51
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "พิมพ์ใบเสร็จ"
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "กรองการค้นหา ตัวที่ 2"
      Height          =   2055
      Left            =   6960
      TabIndex        =   44
      Top             =   2400
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "ตกลง"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   720
         TabIndex        =   45
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label24 
         Caption         =   "ชช - นน - วว"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "จากใบเสร็จของผลกรองตัวที่ 1 ให้มากรอกเวลา"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label22 
         Caption         =   "จนผลลัพธ์เหลือแค่ 1 ใบเสร็จ"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label23 
         Caption         =   "เวลา"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "กรองการค้นหา ตัวที่ 1"
      Height          =   1575
      Left            =   6960
      TabIndex        =   38
      Top             =   840
      Width           =   3615
      Begin VB.ComboBox d 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox MCombo 
         Height          =   315
         ItemData        =   "s_paper.frx":0000
         Left            =   1320
         List            =   "s_paper.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox y 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "2557"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ค้นหาการกรองวันที่"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "วัน - เดือน - ปี พ.ศ."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "เจอใบเสร็จทั้งหมด 0 ใบ"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "วันที่"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label show_pa 
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   5160
      Width           =   9255
   End
   Begin VB.Label Label18 
      Caption         =   "ค้นหาใบเสร็จ/รายการซื้อ"
      Height          =   255
      Left            =   6960
      TabIndex        =   37
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   4560
   End
   Begin VB.Label Label63_1 
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   4320
      Width           =   6855
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   4560
      Y2              =   120
   End
   Begin VB.Label Label64 
      Caption         =   "Label64"
      Height          =   1455
      Left            =   4920
      TabIndex        =   35
      Top             =   14160
      Width           =   1575
   End
   Begin VB.Label Label63 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   15120
      Width           =   6855
   End
   Begin VB.Label Label62 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   14880
      Width           =   6855
   End
   Begin VB.Label Label61 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   14640
      Width           =   6855
   End
   Begin VB.Label Label60 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   14400
      Width           =   6855
   End
   Begin VB.Label Label59 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   14160
      Width           =   6855
   End
   Begin VB.Label Label58 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   13920
      Width           =   6855
   End
   Begin VB.Label Label57 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   13680
      Width           =   6855
   End
   Begin VB.Label Label56 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   13440
      Width           =   6855
   End
   Begin VB.Label Label55 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   13200
      Width           =   6855
   End
   Begin VB.Label Label54 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   12960
      Width           =   6855
   End
   Begin VB.Label Label53 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   12720
      Width           =   6855
   End
   Begin VB.Label Label52 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   12480
      Width           =   6855
   End
   Begin VB.Label Label51 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   12240
      Width           =   6855
   End
   Begin VB.Label Label50 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   12000
      Width           =   6855
   End
   Begin VB.Label Label49 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   11760
      Width           =   6855
   End
   Begin VB.Label Label48 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   11520
      Width           =   6855
   End
   Begin VB.Label Label47 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   11280
      Width           =   6855
   End
   Begin VB.Label Label46 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   11040
      Width           =   6855
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   6855
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   6855
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   6855
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   6855
   End
   Begin VB.Label Label13 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   6855
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   6855
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   6855
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim fso As New FileSystemObject
    Dim fld As Folder
    Dim fil As File
    Dim paper_n As String



Private Sub Form_Load()

'ปิดปุ่ม พิมพ์ใบเสร็จ กับ ตกลง ตอนเริ่มต้น
Command2.Enabled = False
Command3.Enabled = False

MCombo.AddItem "ม.ค."
MCombo.AddItem "ก.พ."
MCombo.AddItem "มี.ค."
MCombo.AddItem "เม.ย."
MCombo.AddItem "พ.ค."
MCombo.AddItem "มิ.ย."
MCombo.AddItem "ก.ค."
MCombo.AddItem "ส.ค."
MCombo.AddItem "ก.ย."
MCombo.AddItem "ต.ค."
MCombo.AddItem "พ.ย."
MCombo.AddItem "ธ.ค."
MCombo = "ม.ค."
For i = 1 To 31
d.AddItem i
Next
d = 1
End Sub


Private Sub Command1_Click()

Dim strPath As String
Dim strFile As String
Dim lngCount As Long
strPath = App.Path & "\ใบเสร็จ\" & d & " " & MCombo & " " & y & "\ใบเสร็จออก ณ วันที่ " & d & " " & MCombo & " " & y & " เวลา " & "*.txt"
strFile = Dir(strPath)
Do Until strFile = ""
lngCount = lngCount + 1
strFile = Dir
Loop
Label20.Caption = "เจอใบเสร็จทั้งหมด " & lngCount & " ใบ"
If (lngCount > 0) Then
    Command2.Enabled = True
    d.Enabled = False
    MCombo.Enabled = False
    y.Enabled = False
    Set fld = fso.GetFolder(App.Path & "\ใบเสร็จ\" & d & " " & MCombo & " " & y)
    For Each fil In fld.Files
      List1.AddItem Left(Right(fil.Name, 16), 12)
    Next
    
    Set fil = Nothing
    Set fld = Nothing
    Set fso = Nothing
    
End If
End Sub

Private Sub Command2_Click()

paper_n = "ใบเสร็จออก ณ วันที่ " & d & " " & MCombo & " " & y & " เวลา " & List1 & ".txt"
show_pa.Caption = "กำลังแสดงชื่อไฟล์ : " & paper_n

'แสดงตัวอย่างไฟล์
Open App.Path & "\ใบเสร็จ\" & d & " " & MCombo & " " & y & "\" & paper_n For Input As #1
Dim text As String
Line Input #1, text
Label1.Caption = text
Line Input #1, text
Line Input #1, text
Label2.Caption = text
Line Input #1, text
Line Input #1, text
Label3.Caption = text
Line Input #1, text
Label4.Caption = text
Line Input #1, text
Label5.Caption = text
Line Input #1, text
Label6.Caption = text
Line Input #1, text
Label7.Caption = text
Line Input #1, text
Label8.Caption = text
Line Input #1, text
Label9.Caption = text
Line Input #1, text
Label10.Caption = text
Line Input #1, text
Label11.Caption = text
Line Input #1, text
Label12.Caption = text
Line Input #1, text
Label13.Caption = text
Line Input #1, text
Label14.Caption = text
Line Input #1, text
Label15.Caption = text
Line Input #1, text
Label16.Caption = text
Line Input #1, text
Label17.Caption = text
Do Until EOF(1)
Line Input #1, text
If (text <> "") Then Label63_1.Caption = text
Loop
Command3.Enabled = True
Close #1
End Sub
Private Sub Command3_Click()
     Form8.Show
End Sub
Private Sub Command4_Click()
    Unload Form9
End Sub
Private Sub Command5_Click()
List1.Clear
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
d.Enabled = True
d = 1
MCombo.Enabled = True
MCombo = "ม.ค."
y.text = "2557"

Label20.Caption = "เจอใบเสร็จทั้งหมด 0 ใบ"
show_pa = ""
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Label63_1.Caption = ""
End Sub
