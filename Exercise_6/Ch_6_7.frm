VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Sitka Heading"
         Size            =   120
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   -480
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Dim num1 As Integer
    Dim num2 As Integer
    num1 = Val(Text1)
    num2 = Val(Text2)
    If num1 <= num2 Then
    Label2 = " = " & getC(num1, num2)
    Else
    Label2 = " = ???"
    End If
End Sub

Public Function getC(valA As Integer, valB As Integer)
    If valA = 0 Then
     getC = 1
     ElseIf valA = 1 Then
     getC = valB
     ElseIf valA = valB Then
     getC = 1
     Else
     getC = getC(valA, valB - 1) + getC(valA - 1, valB - 1)
     End If
End Function
