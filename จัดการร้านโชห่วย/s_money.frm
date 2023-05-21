VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ระบบค้นหายอดขาย"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "ค้นหายอดขาย"
      Height          =   1575
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "ค้นหา"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox y 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "2557"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox MCombo 
         Height          =   315
         ItemData        =   "s_money.frx":0000
         Left            =   1320
         List            =   "s_money.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox d 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "วันที่"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "เจอจำนวนการขายทั้งหมด 0 ครั้ง"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
   End
   Begin VB.CommandButton ออกโปรแกรม 
      Caption         =   "ออก"
      Height          =   375
      Left            =   7200
      Picture         =   "s_money.frx":0004
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">|"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "ไปยังข้อมูลสุดท้าย"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "ไปยังข้อมูลต่อไป"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   300
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "ไปยังข้อมูลก่อนหน้านี้"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "|<"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "ไปยังข้อมูลแรก"
      Top             =   2040
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
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
   Begin VB.Label Label1 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   2040
      Width           =   4815
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim RC As New ADODB.Recordset
Dim SQL As String
Dim money_sum As Currency
Dim i As Integer
Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
Private Sub Command1_Click()
'ค้นหา ยอดขายตามวันที่
With RC
      SQL = "SELECT เวลา , รวมเงิน FROM รวมเงิน_ใบเสร็จ WHERE เวลา LIKE '" & d & "/" & MCombo & "/" & y & " %'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
        Label20.Caption = "เจอจำนวนการขายทั้งหมด " & .RecordCount & " ครั้ง"
        Do Until .EOF = True
            money_sum = money_sum + .Fields("รวมเงิน")
            .MoveNext
       Loop
        Label1.Caption = "รวมยอดเงินจากการขาย ณ วันที่ " & d & "/" & MCombo & "/" & y & " คือ " & money_sum & " บาท"
        money_sum = 0
        SQL = "SELECT เวลา , รวมเงิน FROM รวมเงิน_ใบเสร็จ WHERE เวลา LIKE '" & d & "/" & MCombo & "/" & y & " %'"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
End With
End Sub

Private Sub Form_Load()
money_sum = 0

'เปิดฐานข้อมูล
With conn
        If .State = 1 Then .Close
        .ConnectionString = strConn & ";Data Source=" & App.Path & "\database\goods.mdb"
        .Open
End With
'เปิดตาราง
With RC
        SQL = "SELECT  เวลา , รวมเงิน FROM รวมเงิน_ใบเสร็จ"
        If .State = 1 Then .Close
        .CursorLocation = 3
        .Open SQL, conn, 2, 3
        Set DataGrid1.DataSource = RC
        'ป้องกันการป้อนข้อมูลไปยัง DataGrid
        DataGrid1.AllowUpdate = False
End With
Label20.Caption = ""
For i = 1 To 12
MCombo.AddItem i
Next
MCombo = 1
For i = 1 To 31
d.AddItem i
Next
d = 1
End Sub

Public Sub Command6_Click()
With RC
If .RecordCount > 0 Then
            .MoveFirst
End If
End With
End Sub

Private Sub Command5_Click()
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

Private Sub ออกโปรแกรม_Click()
    Unload Form10
End Sub
