VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "“浮雕”效果"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "→"
      Height          =   1575
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "←"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "新年快乐"
      ForeColor       =   &H8000000A&
      Height          =   378
      Left            =   2535
      TabIndex        =   1
      Top             =   365
      Width           =   980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "新年快乐"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   360
      Picture         =   "Ex_1_4.frx":0000
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label1.Left = Label1.Left - 50
    Label2.Left = Label2.Left - 50
End Sub

Private Sub Command2_Click()
    Label1.Left = Label1.Left + 50
    Label2.Left = Label2.Left + 50
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = Me.MousePointer + 1
End Sub

