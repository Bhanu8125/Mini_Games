VERSION 5.00
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   Caption         =   "Paint Me"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   5055
      Left            =   780
      ScaleHeight     =   4995
      ScaleWidth      =   11235
      TabIndex        =   37
      Top             =   1740
      Width           =   11295
   End
   Begin VB.TextBox txtShapeOutput 
      Height          =   285
      Left            =   3240
      TabIndex        =   35
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cmbShapes 
      Height          =   315
      ItemData        =   "frmprakruthipaint.frx":0000
      Left            =   3240
      List            =   "frmprakruthipaint.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   34
      Text            =   "Select"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame frameColor 
      Caption         =   "Color Tray"
      Height          =   1095
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdMore 
         Caption         =   "More Colors"
         Height          =   615
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCurr 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H0039D8EE&
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2040
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00B6FEDF&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H0050B8B8&
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FAFCA5&
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F8B9FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblWhite 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -360
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblBlack 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblGray 
         BackColor       =   &H00808080&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblBrown 
         BackColor       =   &H00000080&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lclColor5 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblOrange 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblYellow 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblDGreen 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLBlue 
         BackColor       =   &H00F3D065&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblDBlue 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblPurple 
         BackColor       =   &H00D67EB9&
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   38
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   225
      Width           =   375
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "1"
      Top             =   225
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cu&t"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "&Copy"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Paste"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   10380
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblCircle 
      Height          =   375
      Left            =   10440
      TabIndex        =   36
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblShapes 
      Caption         =   "Shapes"
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Brush Width"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim drawing
Dim cX As Integer
Dim cY As Integer
Dim color
Dim index As Long

Private Sub cmbShapes_Click()
    txtShapeOutput.Text = cmbShapes.ListIndex
    If cmbShapes.Text = "None" Then
        drawing = True
    Else
        drawing = False
        rad = Abs(px - cX)
        'Picture1.Circle (px, py), rad, vbRed
    End If
End Sub

Private Sub cmdClear_Click()
    Picture1.Cls
    cX = 0
    cY = 0
    px = 0
    py = 0
    
End Sub

Private Sub cmdMore_Click()
    CommonDialog1.ShowColor
    Picture1.ForeColor = CommonDialog1.color
    lblCurr.BackColor = Picture1.ForeColor
    
End Sub

Private Sub cmdOkay_Click()
    Picture1.DrawWidth = txtWidth.Text
End Sub

Private Sub Form_Load()
    lblCurr.BackColor = Picture1.ForeColor
End Sub



Private Sub Label10_Click()
    Picture1.ForeColor = &H80FF80
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label11_Click()
    Picture1.ForeColor = &HC0FFFF
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label12_Click()
     Picture1.ForeColor = &H39D8EE
     lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label13_Click()
     Picture1.ForeColor = &H80C0FF
     lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()
    Picture1.ForeColor = &HF8B9FF
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label5_Click()
     Picture1.ForeColor = &HFAFCA5
     lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label6_Click()
    Picture1.ForeColor = &H50B8B8
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label7_Click()
    Picture1.ForeColor = &HFF8080
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label8_Click()
    Picture1.ForeColor = &HB6FEDF
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Label9_Click()
    Picture1.ForeColor = &HFFC0C0
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblBrown_Click()
    Picture1.ForeColor = &H80&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblDBlue_Click()
    Picture1.ForeColor = &HFF0000
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblDGreen_Click()
    Picture1.ForeColor = &HC000&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblGray_Click()
    Picture1.ForeColor = &H808080
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblGreen_Click()
     Picture1.ForeColor = &HFF00&
     lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblLBlue_Click()
    Picture1.ForeColor = &HF3D065
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblOrange_Click()
    Picture1.ForeColor = &H80FF&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblPurple_Click()
    Picture1.ForeColor = &HD67EB9
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblWhite_Click()
    Picture1.ForeColor = &HFFFFFF
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblBlack_Click()
    Picture1.ForeColor = &H0&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lblYellow_Click()
    Picture1.ForeColor = &HFFFF&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub lclColor5_Click()
    Picture1.ForeColor = &HFF&
    lblCurr.BackColor = Picture1.ForeColor
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmbShapes.ListIndex = 0 Or cmbShapes.ListIndex = 1 Or cmbShapes.ListIndex = 7 Or cmbShapes.ListIndex = 3 Or cmbShapes.ListIndex = 4 Or cmbShapes.ListIndex = 5 Or cmbShapes.ListIndex = 6 Then
        cX = X
        cY = Y
        drawing = False
    Else
        drawing = True
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If drawing = True Then
        Picture1.PSet (X, Y)
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drawing = False
    If cmbShapes.Text = "Circle" Then
        px = X
        py = Y
        rad = Abs(px - cX)
        Picture1.Circle (X, Y), rad, Picture1.ForeColor
    ElseIf cmbShapes.ListIndex = 1 Then
        Picture1.Line (X, Y)-(cX, cY), Picture1.ForeColor
    ElseIf cmbShapes.ListIndex = 7 Then
        Picture1.LinkSend
    End If
    
End Sub

