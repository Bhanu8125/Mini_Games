VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "Paint"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Height          =   975
      Left            =   7500
      Picture         =   "frmPaint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdEraser 
      Caption         =   "Eraser"
      Height          =   495
      Left            =   2460
      TabIndex        =   28
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdWidthSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   120
      Width           =   435
   End
   Begin VB.ComboBox comboBrushWidth 
      Height          =   315
      Left            =   2520
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox ComboShapes 
      Height          =   315
      Left            =   2520
      TabIndex        =   22
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      Height          =   6015
      Left            =   60
      MouseIcon       =   "frmPaint.frx":4FAA
      MousePointer    =   99  'Custom
      ScaleHeight     =   5955
      ScaleWidth      =   8355
      TabIndex        =   7
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Frame frameColourTray 
      Caption         =   "Colour Tray"
      Height          =   975
      Left            =   4380
      TabIndex        =   4
      Top             =   60
      Width           =   3075
      Begin VB.Label lblLightGrey 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblYellow 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   600
         Width           =   315
      End
      Begin VB.Label LeafGreen 
         BackColor       =   &H0000FF00&
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblWhite 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblBrown 
         BackColor       =   &H00000040&
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblDarkGrey 
         BackColor       =   &H00404040&
         Height          =   315
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblLightBlue 
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblLightGreen 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblGray 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblLavender 
         BackColor       =   &H00FF00FF&
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblOrange 
         BackColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblViolet 
         BackColor       =   &H00800080&
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblGreen 
         BackColor       =   &H00008000&
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblBlue 
         BackColor       =   &H00FF0000&
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblBlack 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5100
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   495
      Left            =   660
      TabIndex        =   3
      Top             =   1080
      Width           =   675
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cut"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   915
      Left            =   780
      Picture         =   "frmPaint.frx":53EC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   555
   End
   Begin VB.CommandButton cmdPencil 
      Caption         =   "Pencil"
      Height          =   915
      Left            =   60
      Picture         =   "frmPaint.frx":582E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lblShapes 
      Caption         =   "Shapes"
      Height          =   255
      Left            =   1500
      TabIndex        =   25
      Top             =   660
      Width           =   855
   End
   Begin VB.Label lblBrushWidth 
      Caption         =   "Brush Width"
      Height          =   255
      Left            =   1500
      TabIndex        =   23
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim px, py, cx, cy, rad, dx, dy As Integer
Dim PencilClicked, Drawing, Initial As Boolean
Dim color

Private Sub Label1_Click()

End Sub

Private Sub Combo1_Change()
    
End Sub

Private Sub cmdClear_Click()
   Picture1.Cls
    cx = 0
    cy = 0
    px = 0
    py = 0
End Sub

Private Sub cmdPencil_Click()
    Picture1.Enabled = True
    PencilClicked = True
    If PencilClicked Then
        Initial = False
        ComboShapes.ListIndex = 0
    End If
End Sub
Private Sub cmdWidthSet_Click()
    Picture1.DrawWidth = comboBrushWidth.Text
End Sub

Private Sub Command1_Click()
    CommonDialog1.ShowColor
    Picture1.ForeColor = CommonDialog1.color
End Sub

Private Sub Form_Load()
        comboBrushWidth.AddItem ("")
        comboBrushWidth.AddItem (1)
        comboBrushWidth.AddItem (5)
        comboBrushWidth.AddItem (10)
        comboBrushWidth.AddItem (15)
        comboBrushWidth.AddItem (20)
        comboBrushWidth.AddItem (25)
        comboBrushWidth.AddItem (30)
        comboBrushWidth.AddItem (35)
        comboBrushWidth.AddItem (40)
        comboBrushWidth.AddItem (45)
        comboBrushWidth.AddItem (50)
        comboBrushWidth.AddItem (55)
        comboBrushWidth.AddItem (60)
        ComboShapes.AddItem ("")
        ComboShapes.AddItem ("Circle")
        ComboShapes.AddItem ("Line")
        ComboShapes.AddItem ("Rectangle")
        ComboShapes.AddItem ("Square")
        ComboShapes.AddItem ("Oval")
        ComboShapes.AddItem ("Rounded Rectangle")
        ComboShapes.AddItem ("Rounded Square")
        Initial = True
End Sub

Private Sub lblBlack_Click()
    Picture1.ForeColor = &H0&
End Sub

Private Sub lblBlue_Click()
    Picture1.ForeColor = &HFF0000
End Sub

Private Sub lblBrown_Click()
    Picture1.ForeColor = &H40&
End Sub
Private Sub lblDarkGrey_Click()
    Picture1.ForeColor = &H404040
End Sub

Private Sub lblGray_Click()
    Picture1.ForeColor = &H808080
End Sub

Private Sub lblGreen_Click()
    Picture1.ForeColor = &H8000&
End Sub

Private Sub lblLavender_Click()
      Picture1.ForeColor = &HFF00FF
End Sub

Private Sub lblLightBlue_Click()
    Picture1.ForeColor = &HFF8080
End Sub

Private Sub lblLightGreen_Click()
Picture1.ForeColor = &HFFFF80
End Sub

Private Sub lblLightGrey_Click()
    Picture1.ForeColor = &HE0E0E0
End Sub

Private Sub lblOrange_Click()
    Picture1.ForeColor = &H80FF&
End Sub

Private Sub lblRed_Click()
Picture1.ForeColor = &HFF&
End Sub

Private Sub lblViolet_Click()
    Picture1.ForeColor = &H800080
End Sub

Private Sub lblWhite_Click()
    Picture1.ForeColor = &HFFFFFF
End Sub

Private Sub lblYellow_Click()
    Picture1.ForeColor = &HFFFF&
End Sub

Private Sub LeafGreen_Click()
   Picture1.ForeColor = &HFF00&
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (ComboShapes.ListIndex = 1 Or ComboShapes.ListIndex = 2 Or ComboShapes.ListIndex = 3 Or ComboShapes.ListIndex = 4 Or ComboShapes.ListIndex = 5 Or ComboShapes.ListIndex = 6 Or ComboShapes.ListIndex = 7 Or ComboShapes.ListIndex = 8) And ComboShapes.Text <> "" Then
        cx = X
        cy = Y
        Initial = True
    End If
    If Initial = True Then
        Drawing = False
    Else
        Drawing = True
'        dx = X
'        dy = Y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Drawing = True Then
        Picture1.PSet (X, Y)
    End If
' If Drawing = True Then
'        Picture1.Line (X, Y)-(dx, dy)
'    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Drawing = False
    If ComboShapes.Text = "Circle" Then
        px = X
        py = Y
        rad = Abs(px - cx)
        Picture1.Circle (X, Y), rad, Picture1.ForeColor
        cx = 0
        px = 0
        cy = 0
        py = 0
        ElseIf ComboShapes.Text = "Line" Then
        Picture1.Line (X, Y)-(cx, cy), Picture1.ForeColor
    End If
End Sub
