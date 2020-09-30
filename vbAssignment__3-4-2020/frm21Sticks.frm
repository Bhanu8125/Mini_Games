VERSION 5.00
Begin VB.Form frm21Sticks 
   Caption         =   "21 STICKS"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnterNumber 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H8000000A&
      Caption         =   "Check"
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Top             =   2700
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdTryAgain 
      Caption         =   "Try Again"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1620
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblSticks 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "21 Sticks Remaining"
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
      TabIndex        =   6
      Top             =   540
      Width           =   2175
   End
   Begin VB.Label lblRandom 
      Caption         =   "Random Number"
      Height          =   255
      Left            =   900
      TabIndex        =   5
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label lblEnterNumber 
      AutoSize        =   -1  'True
      Caption         =   "Enter  Your  Number"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1380
      Width           =   1500
   End
End
Attribute VB_Name = "frm21Sticks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number, Sticks, Index, Temp As Integer
Dim Check As Boolean


Private Sub cmdCheck_Click()
If InStr(txtEnterNumber.Text, ".") Or Len(CStr(txtEnterNumber.Text)) > 1 Or txtEnterNumber.Text = Empty Then
        MsgBox "Please Enter Numbers Between 1 to 4", vbOKOnly
        txtEnterNumber.Text = Empty
        Exit Sub
    End If
    Dim Random As Integer
    If Sticks > 16 Then
        Random = 0
       Do Until Random > 0
Repeat1:     Random = Val(10 * (Rnd))
            If Random < 1 Or Random > 4 Then
                GoTo Repeat1
            End If
        Loop
        Sticks = Sticks - (Random + Number)
        txtDisplayRandom.Text = CStr(Random)
        lblSticks.Caption = Sticks & " Remaining Sticks"
    
    ElseIf Sticks > 10 Then
         Random = 0
       Do Until Random > 0
Repeat2:     Random = Val(10 * (Rnd))
            If Random < 1 Or Random > 3 Then
                GoTo Repeat2
            End If
'            Temp = Sticks - (Random + Number)
'            If Temp > 0 Then
'                Sticks = Temp
'            Else
'                GoTo Repeat1
'            End If
        Loop
        Sticks = Sticks - (Random + Number)
        txtDisplayRandom.Text = CStr(Random)
        If Sticks <= 6 Then
            lblSticks.BackColor = vbRed
        End If
        lblSticks.Caption = Sticks & " Remaining Sticks"
    Else
        Sticks = Sticks - Number
        For Index = 1 To 2
Repeat3:    Random = Val(10 * Rnd)
            If (Random < 1 Or Random > 4) Or (Sticks - Random < 0) Then
                GoTo Repeat3
            End If
            If Sticks - Random = 1 Then
                Check = True
                Exit For
            Else
                GoTo Repeat2
            End If
        Next
        If Check Then
            MsgBox "You Lose The Game", vbOKOnly, "GameOver"
        Else
           MsgBox "You Won The Game", vbOKOnly, "GameOver"
        End If
    End If
        txtEnterNumber.Text = ""
        txtEnterNumber.SetFocus
        txtDisplayRandom.Text = " "
End Sub

Private Sub cmdTryAgain_Click()
        cmdCheck.Enabled = True
        cmdTryAgain.Enabled = False
        txtEnterNumber.Text = ""
        txtEnterNumber.SetFocus
        txtDisplayRandom.Text = " "
End Sub

Private Sub Form_Load()
    Sticks = 21
End Sub

Private Sub txtEnterNumber_Change()
    Number = Val(txtEnterNumber.Text)
End Sub

