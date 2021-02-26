VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "×Ö·û¸ñÊ½"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   ScaleHeight     =   3330
   ScaleWidth      =   4785
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "15 °õ"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "25 °õ"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Á¥Êé"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ó×Ô²"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1217
      ButtonWidth     =   609
      ButtonHeight    =   1058
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "VB is powerful but needs efforts to learn"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.FontName = "Ó×Ô²"
End Sub

Private Sub Command2_Click()
    Text1.FontName = "Á¥Êé"
End Sub

Private Sub Command3_Click()
    Text1.FontSize = 25
End Sub

Private Sub Command4_Click()
    Text1.FontSize = 15
End Sub
