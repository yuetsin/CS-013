VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "↑↓"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   5640
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   6615
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "工资"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "姓名"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "工号"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type EmployeeType
    ID As Long
    name As String
    salary As Integer
End Type
Dim EmployeeData(5) As EmployeeType
Dim employIndex As Integer
Private Sub Command1_Click()
    If employIndex <> 4 Then
        If Int(Val(Text2.Text)) <> 0 Then
            If Int(Val(Text4.Text)) <> 0 Then
                If Text3.Text <> "" Then
                    EmployeeData(employIndex).ID = Int(Val(Text2.Text))
                    EmployeeData(employIndex).name = Text3.Text
                    EmployeeData(employIndex).salary = Int(Val(Text4.Text))
                    employIndex = employIndex + 1
                    MsgBox ("添加成功√")
                Else
                    MsgBox ("格式有误×")
                End If
            Else
                MsgBox ("格式有误×")
            End If
        Else
            MsgBox ("格式有误×")
        End If
    Else
        MsgBox ("已经达到添加上限")
    End If
End Sub

Private Sub Command2_Click()
    For i = 0 To employIndex
        For j = i To employIndex
            If EmployeeData(i).salary < EmployeeData(j).salary Then
                Dim tempObj As EmployeeType
                tempObj = EmployeeData(j)
                EmployeeData(j) = EmployeeData(i)
                EmployeeData(i) = tempObj
            End If
        Next
    Next
    Text1.Text = "姓名    ID    工资" & vbCrLf
    For i = 0 To employIndex - 1
        Text1.Text = Text1.Text & EmployeeData(i).name & " " & EmployeeData(i).ID & " " & EmployeeData(i).salary & vbCrLf
    Next
End Sub

Private Sub Form_Load()
    employIndex = 0
End Sub
