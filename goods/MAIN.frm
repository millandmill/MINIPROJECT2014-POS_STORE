VERSION 5.00
Begin VB.MDIForm MAIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ระบบจัดการร้านโชห่วย"
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20355
   LinkTopic       =   "MDIForm1"
   Picture         =   "MAIN.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Form7.Show
End Sub
