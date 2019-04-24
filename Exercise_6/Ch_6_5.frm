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
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Text            =   "input your string here"
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strList() As String
Private Sub Command1_Click()
    Dim rawStr As String
    Dim index As Integer
    Dim LongestStr As String
    rawStr = Text1
    LongestStr = ""
    ReDim strList(UBound(Split(rawStr))) As String
    strList = Split(rawStr)
    For index = 0 To UBound(strList)
    If Len(strList(index)) <> 0 Then
        Combo1.AddItem (strList(index))
        If Len(LongestStr) < Len(strList(index)) Then
            LongestStr = strList(index)
        End If
    End If
        Next
     Print LongestStr
     Combo1.Text = "Longest Str: " & LongestStr
End Sub
