VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = Label1.Caption & "generated "
    Dim intArray(10) As Integer
    Dim i As Integer
    Dim sumUp As Integer
    Dim max As Integer
    max = 29
    Dim min As Integer
    min = 101
    For i = 0 To 9
        Randomize
        Dim a As Integer
        a = Int(Rnd * 71) + 30
        Label1.Caption = Label1.Caption & a & " "
        sumUp = sumUp + a
        If max < a Then
            max = a
        End If
        If min > a Then
            min = a
        End If
        Next
    Label2.Caption = "平均值 " & sumUp / 10 & " 最大值 " & max & "最小值 " & min
End Sub
