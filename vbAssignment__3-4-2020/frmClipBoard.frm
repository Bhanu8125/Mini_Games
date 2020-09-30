VERSION 5.00
Begin VB.Form frmClipBoard 
   Caption         =   "&Copy"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3180
      TabIndex        =   3
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtMessage 
      Height          =   2955
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   5835
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   915
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    Clipboard.Clear
End Sub

Private Sub cmdCopy_Click()
    If txtMessage.Text = Empty Then
        MsgBox "No Text To Copy", vbCritical + vbOKOnly, frmClipBoard.Caption
        Exit Sub
    End If
    If Len(txtMessage.SelText) > 0 Then
        Clipboard.SetText txtMessage.SelText
    Else
        Clipboard.SetText txtMessage.Text
    End If
End Sub

Private Sub cmdPaste_Click()
    Dim cliptext As String
        cliptext = Clipboard.GetText
        If cliptext = Empty Then
            MsgBox "No Text to Paste", vbCritical + vbOKOnly, frmClipBoard.Caption
            Exit Sub
        End If
        Dim ipos As Long
'        With txtMessage
'            If .SelLength = 0 Then
'                ipos = .SelStart
'            Else
'                ipos = .SelStart + .SelLength
'            End If
'        End With
        'txtMessage.SelStart = ipos
        txtMessage.SelText = Clipboard.GetText
End Sub

Private Sub txtMessage_Change()

End Sub
