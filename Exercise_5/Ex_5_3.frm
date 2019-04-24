VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim subScore As Variant
    subScore = Array(3, 2, 3, 4, 1)
    Dim subMark As Variant
    subMark = Array(78, 98, 83, 68, 90)
    Label1.Caption = Label1.Caption & "成绩为："
    For i = 0 To 4
        Label1.Caption = Label1.Caption & subMark(i) & " "
    Next
        Label1.Caption = Label1.Caption & vbCrLf & "学分为："
    For i = 0 To 4
        Label1.Caption = Label1.Caption & subScore(i) & " "
    Next
       
        Dim GPA As Double
        Dim allIn As Double
        Dim allScore As Integer
        allIn = 0
        allScore = 0
    For i = 0 To 4
       allScore = allScore + subScore(i)
        If subMark(i) >= 90 Then
            allIn = allIn + 4 * subScore(i)
        ElseIf subMark(i) >= 80 Then
            allIn = allIn + 3 * subScore(i)
           ElseIf subMark(i) >= 70 Then
            allIn = allIn + 2 * subScore(i)
            ElseIf subMark(i) >= 60 Then
            allIn = allIn + 1 * subScore(i)
            End If
        Next
        GPA = allIn / allScore
         Label1.Caption = Label1.Caption & vbCrLf & "GPA 为：" & GPA
End Sub
