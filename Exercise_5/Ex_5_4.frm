VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Trace"
      Height          =   855
      Left            =   7800
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "max(A)"
      Height          =   615
      Left            =   7800
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Transform"
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A + B"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   855
      Left            =   7920
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   3015
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   7095
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim MatrixA(16) As Integer
    Dim MatrixB(16) As Integer

Private Sub Command1_Click()
    Dim randVar As Integer
    Text1.Text = ""
    Text2.Text = ""
    For i = 0 To 3
        For j = 0 To 3
            Randomize
            randVar = Int(Rnd * 41) + 30
            MatrixA(i * 4 + j) = randVar
            Text1.Text = Text1.Text & randVar & " "
            Next
        Text1.Text = Text1.Text & vbCrLf
    Next
    For i = 0 To 3
        For j = 0 To 3
            Randomize
            randVar = Int(Rnd * 35) + 101
            MatrixB(i * 4 + j) = randVar
            Text2.Text = Text2.Text & randVar & " "
            Next
        Text2.Text = Text2.Text & vbCrLf
    Next
End Sub

Private Sub Command2_Click()
    Text3.Text = "A + B Matrix is: " & vbCrLf
    Dim var As Integer
    For i = 1 To 16
        var = MatrixA(i - 1) + MatrixB(i - 1)
        Text3.Text = Text3.Text & var & " "
        If i Mod 4 = 0 Then
            Text3.Text = Text3.Text & vbCrLf
        End If
    Next
    Text3.Text = Text3.Text & vbCrLf
End Sub

Private Sub Command3_Click()
    Text3.Text = "A's Transform is: " & vbCrLf
    Dim var As Integer
    For i = 0 To 3
        For j = 0 To 3
            var = MatrixA(j * 4 + i)
            Text3.Text = Text3.Text & var & " "
            Next
        Text3.Text = Text3.Text & vbCrLf
    Next
End Sub

Private Sub Command4_Click()
    Text3.Text = "The max value index of Matrix A is: " & vbCrLf
    Dim maxIndexX As Integer
    Dim maxIndexY As Integer
    Dim maxVal As Integer
    maxVal = 0
    For i = 0 To 3
        For j = 0 To 3
           If MatrixA(i * 4 + j) > Max Then
            maxIndexX = i
            maxIndexY = j
            Max = MatrixA(i * 4 + j)
           End If
        Next
    Next
    Text3.Text = Text3.Text & "(" & maxIndexX + 1 & ", " & maxIndexY + 1 & ") = " & Max
End Sub

Private Sub Command5_Click()
    Text3.Text = "The trace of Matrix A is: " & vbCrLf
    Dim sumUp As Integer
    sumUp = MatrixA(0) + MatrixA(5) + MatrixA(10) + MatrixA(15)
    Text3.Text = Text3.Text & sumUp
End Sub
