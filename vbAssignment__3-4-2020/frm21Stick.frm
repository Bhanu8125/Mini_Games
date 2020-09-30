VERSION 5.00
Begin VB.Form frm21Stick 
   Caption         =   "21 STICKS"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTryAgain 
      Caption         =   "Try Again"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1140
      TabIndex        =   3
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox txtDisplayRandom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   2
      Top             =   1620
      Width           =   615
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H8000000A&
      Caption         =   "Check"
      Height          =   495
      Left            =   1140
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtEnterNumber 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      Caption         =   "Enter  Your  Number"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lblRandom 
      Caption         =   "Random Number"
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblSticks 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "Total Sticks 21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1020
      TabIndex        =   4
      Top             =   300
      Width           =   1560
   End
End
Attribute VB_Name = "frm21Stick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim U_Num, G_Num, Five, Sticks As Integer

Private Sub cmdCheck_Click()
    If InStr(txtEnterNumber.Text, ".") Or txtEnterNumber.Text = Empty Or U_Num < 1 Or U_Num > 4 Then
        MsgBox "Please Enter Numbers Between 1 to 4", vbOKOnly
        txtEnterNumber.Text = Empty
        txtEnterNumber.SetFocus
        Exit Sub
    End If
    G_Num = Five - U_Num
    txtDisplayRandom.Text = Val(G_Num)
    Sticks = Sticks - (U_Num + G_Num)
    lblSticks.Caption = "Remaining Sticks " & Sticks
    txtEnterNumber.Text = Empty
    txtEnterNumber.SetFocus
    If Sticks = 1 Then
        MsgBox " You Lose The Game ", vbOKOnly, "GameOver"
        cmdCheck.Enabled = False
        txtDisplayRandom.Text = Empty
        cmdTryAgain.Enabled = True
    End If
End Sub

Private Sub cmdTryAgain_Click()
    cmdCheck.Enabled = True
    cmdTryAgain.Enabled = False
    Sticks = 21
    lblSticks.Caption = "Total sticks 21"
End Sub

Private Sub Form_Load()
    Five = 5
    Sticks = 21
End Sub
Private Sub txtEnterNumber_Change()
    U_Num = Val(txtEnterNumber.Text)
End Sub
