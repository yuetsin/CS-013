VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算圆面积与周长"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "周长"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "面积"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "半径"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Not IsNumeric(Text1.Text) Then
        MsgBox ("invalid input!")
        Text1.SetFocus
    Else
        Text2.Text = Format(3.141592653589 * Val(Text1.Text) * Val(Text1.Text), "#.##")
    End If
End Sub

Private Sub Command2_Click()
    If Not IsNumeric(Text1.Text) Then
        MsgBox ("invalid input!")
        Text1.SetFocus
    Else
        Text3.Text = Format((3.141592653589 * Val(Text1.Text) * 2), "#.##")
    End If
End Sub
