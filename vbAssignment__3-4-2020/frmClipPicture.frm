VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmClipPicture 
   Caption         =   "Clip Picture"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Image imgCapture 
      Height          =   3075
      Left            =   540
      Stretch         =   -1  'True
      Top             =   420
      Width           =   5415
   End
End
Attribute VB_Name = "frmClipPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path1, Path2 As String

Private Sub cmdBrowse_Click()
    Dim filename As String
    CommonDialog1.Filter = "All Images (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    If CommonDialog1.FilterIndex = 1 Then
        filename1 = CommonDialog1.filename
        imgCapture.Picture = LoadPicture(filename1)
    Else
        MsgBox "Invalid file,please choose current file"
        Exit Sub
    End If
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
     Clipboard.SetData imgCapture.Picture
End Sub

Private Sub cmdPaste_Click()
    imgCapture.Picture = Clipboard.GetData
End Sub

Private Sub Image1_Click()
     
End Sub

Private Sub Form_Load()
    Path1 = "C:\Users\EVV RAMANA\Desktop\1.jpg"
    'Path2 = "C:\Users\EVV RAMANA\Desktop\2.jpg"
    imgCapture.Picture = LoadPicture(Path1)
End Sub
