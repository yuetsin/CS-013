VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Height          =   4215
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim gradeSc(5) As Integer
    For i = 1 To 20
        Randomize
        Dim a As Integer
        a = Int(Rnd * 101)
        Label1.Caption = Label1.Caption & a & " "
        If i Mod 5 = 0 Then
        Label1.Caption = Label1.Caption & vbCrLf
        End If
        If a < 60 Then
        gradeSc(0) = gradeSc(0) + 1
        ElseIf a < 70 Then
        gradeSc(1) = gradeSc(1) + 1
        ElseIf a < 80 Then
        gradeSc(2) = gradeSc(2) + 1
        ElseIf a < 90 Then
        gradeSc(3) = gradeSc(3) + 1
        Else
        gradeSc(4) = gradeSc(4) + 1
        End If
        Next
Dim info As String
    For i = 0 To 4
        info = "s(" & i + 5 & ") ÈËÊýÎª " & gradeSc(i) & vbCrLf
        List1.AddItem (info)
        Next
End Sub

