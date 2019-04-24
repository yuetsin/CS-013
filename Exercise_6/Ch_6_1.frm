VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   5280
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "get min"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add number"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2760
      ItemData        =   "Ch_6_1.frx":0000
      Left            =   240
      List            =   "Ch_6_1.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "MinValue:"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a() As Integer
Dim minVal As Integer
Dim flag As Boolean
Private Sub Command1_Click()
    List1.AddItem (Val(Text1))
    If flag Then
        ReDim Preserve a(1) As Integer
        a(0) = Val(Text1)
    Else
        ReDim Preserve a(UBound(a)) As Integer
        a(UBound(a)) = Val(Text1)
    End If
    flag = False
End Sub

Private Sub Command2_Click()
        Call ProcMin(a, minVal)
        Label1 = "MinValue:" & minVal
End Sub

Public Function ProcMin(ByRef a() As Integer, ByRef mina As Integer)
    mina = 32767
    Dim i As Integer
    For i = LBound(a) To UBound(a)
        If a(i) < mina Then
        mina = a(i)
        End If
    Next
End Function

Private Sub Form_Load()
    flag = True
End Sub
