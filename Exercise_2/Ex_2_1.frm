VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "温度的转换"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4410
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "Auto"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-"
      Height          =   1215
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "->"
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "invalid value!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "fahrenheit"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "celcius"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim isAuto As Boolean

Private Sub Check1_Click()
    If Check1.Value = vbUnchecked Then
        Command1.Enabled = True
        Command2.Enabled = True
        isAuto = False
    Else
        Command1.Enabled = False
        Command2.Enabled = False
        isAuto = True
    End If
End Sub

Private Sub Command1_Click()
    Text2.Text = Str(Val(Text1.Text) * 9 / 5 + 32)
End Sub

Private Sub Command2_Click()
    Text1.Text = Str((Val(Text2.Text) - 32) * 5 / 9)
End Sub

Private Sub Form_Load()
    isAuto = False
End Sub

Private Sub Text1_Change()
    If Val(Text1.Text) < -273.15 Then
        Label3.Visible = True
    Else
        Label3.Visible = False
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If isAuto Then
        Text2.Text = Str(Val(Text1.Text) * 9 / 5 + 32)
    End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If isAuto Then
        Text2.Text = Str(Val(Text1.Text) * 9 / 5 + 32)
    End If
End Sub

Private Sub Text2_Change()
    If Val(Text2.Text) < -459.67 Then
        Label3.Visible = True
    Else
        Label3.Visible = False
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If isAuto Then
        Text1.Text = Str((Val(Text2.Text) - 32) * 5 / 9)
    End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If isAuto Then
        Text1.Text = Str((Val(Text2.Text) - 32) * 5 / 9)
    End If
End Sub
