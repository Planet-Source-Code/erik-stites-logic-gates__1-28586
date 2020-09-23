VERSION 5.00
Begin VB.Form frmLogic 
   Caption         =   "Logic Gates"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Build Truth Table"
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox pbxGate 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "Nand"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "Xnor"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CheckBox chkB 
      Caption         =   "B"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CheckBox chkA 
      Caption         =   "A"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "Nor"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "Xor"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "And"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optLogic 
      Caption         =   "Or"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdLogic 
      Caption         =   "Compare A and B"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4200
      Y1              =   1920
      Y2              =   240
   End
   Begin VB.Label lblOut 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   25
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblOut 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   24
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblOut 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   23
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblOut 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblB 
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   20
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblA 
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   19
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblB 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   18
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblA 
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   17
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblB 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblA 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblB 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblA 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   4
      Left            =   4200
      Picture         =   "Form1.frx":2162
      Top             =   2760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   5
      Left            =   5040
      Picture         =   "Form1.frx":42C4
      Top             =   2760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   3
      Left            =   3360
      Picture         =   "Form1.frx":6426
      Top             =   2760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   2
      Left            =   5040
      Picture         =   "Form1.frx":8588
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   0
      Left            =   3360
      Picture         =   "Form1.frx":A6EA
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgGate 
      Height          =   795
      Index           =   1
      Left            =   4200
      Picture         =   "Form1.frx":C84C
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblX 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Output"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LogicType As Integer

Private Sub cmdLogic_Click()
    'uses absolute value because my comp would return a negative 1 instead of a 1 when the output was true
    lblX.Caption = Abs(Int(CompareLogic(LogicType, chkA.Value, chkB.Value)))
End Sub

Private Sub Command1_Click()
    Dim TableOut As Integer
    
    For TableOut = 0 To 3
        'uses absolute value because my comp would return a negative 1 instead of a 1 when the output was true
        lblOut(TableOut).Caption = Abs(Int(CompareLogic(LogicType, lblA(TableOut).Caption, lblB(TableOut).Caption)))
    Next TableOut
End Sub

Private Sub optLogic_Click(Index As Integer)
    LogicType = Index
    pbxGate.Picture = imgGate(Index).Picture
End Sub

Private Function CompareLogic(CompareType As Integer, In1 As Boolean, In2 As Boolean) As Boolean
    Select Case CompareType
        Case Is = 0
            CompareLogic = In1 Or In2
        Case Is = 1
            CompareLogic = In1 And In2
        Case Is = 2
            CompareLogic = In1 Xor In2
        Case Is = 3
            CompareLogic = Not (In1 Or In2)
        Case Is = 4
            CompareLogic = Not (In1 And In2)
        Case Is = 5
            CompareLogic = Not (In1 Xor In2)
    End Select
End Function

