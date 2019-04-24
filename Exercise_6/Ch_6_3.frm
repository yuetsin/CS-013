VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get!"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "log"
      Height          =   4695
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Label1 = ""
    Dim x As Double
    x = Val(Text1)
    Form1.Caption = "Result = " & mySin(x)
End Sub

Public Function Factorial(ByVal i As Integer)
    Dim j As Integer
    Dim sumUp As Double
    sumUp = 1
    For j = 1 To i
        'Label1 = Label1 & "fact info: " & sumUp & vbCrLf
        sumUp = sumUp * j
        Next
    Factorial = sumUp
End Function

Public Function mySin(x As Double)
    Dim eachItem As Double
    Dim sumUp As Double
    Dim count As Integer
    count = 1
    sumUp = 0
    Do While True
        eachItem = (-1) ^ ((count + 3) / 2) * (x ^ count) / CDbl(Factorial(count))
        Label1 = Label1 & "get one item: " & eachItem & vbCrLf
        If Abs(eachItem) < 0.00001 Then
            Exit Do
        Else
            sumUp = sumUp + eachItem
            count = count + 2
        End If
    Loop
    mySin = sumUp
End Function

