VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "选电脑"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   8850
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "立即购买"
      Height          =   855
      Left            =   2640
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "软驱"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "鼠标"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Intel i9 7980XE"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Pantium III"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pantium II"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      IntegralHeight  =   0   'False
      ItemData        =   "Ch_7_1.frx":0000
      Left            =   240
      List            =   "Ch_7_1.frx":0002
      TabIndex        =   1
      Text            =   "美帝良心想"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "CPU"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "内存容量"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "品牌"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text1 <> "" Then
        Text2 = "Manifest here:" & vbCrLf
        Text2 = Text2 & Combo1.Text & vbCrLf & Text1 & vbCrLf
        If Check1.Value Then
         Text2 = Text2 & "鼠标" & vbCrLf
         End If
         If Check2.Value Then
         Text2 = Text2 & "软驱" & vbCrLf
         End If
         If Option1.Value Then
         Text2 = Text2 & "CPU: Pantium II" & vbCrLf
         ElseIf Option2.Value Then
         Text2 = Text2 & "CPU: Pantium III" & vbCrLf
         ElseIf Option3.Value Then
         Text2 = Text2 & "CPU: Intel i9 7980XE" & vbCrLf
         End If
    End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("美帝良心想")
Combo1.AddItem ("奸如磐石硕")
Combo1.AddItem ("偷工减料碁")
Combo1.AddItem ("做工渣渣船")
Combo1.AddItem ("人傻钱多戴")
Combo1.AddItem ("铁板熊掌普")
Combo1.AddItem ("宗教信仰果")
Combo1.AddItem ("专业贴牌尔")
Combo1.AddItem ("同方勇气多")
Combo1.AddItem ("散热缩水星")
Option1 = True
End Sub

Private Sub Text1_LostFocus()
    Dim str As String
    str = Text1
    If Right(str, 2) = "MB" And IsNumeric(Left(str, Len(str) - 2)) Then
        
    Else
        Text1 = ""
        End If
End Sub
