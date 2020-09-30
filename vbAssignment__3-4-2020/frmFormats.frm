VERSION 5.00
Begin VB.Form frmFormats 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUserType 
      Caption         =   "UserType"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPredefined 
      Caption         =   "Predefined"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   780
      Width           =   480
   End
End
Attribute VB_Name = "frmFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPredefined_Click()
    lblDisplay.Caption = ""
    lblDisplay.Caption = Format(8972.234, "General Number") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(8972.2, "Fixed") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(6648972.265, "Standard") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(6648972.265, "currency") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(0.56324, "Percent") & vbNewLine
End Sub

Private Sub cmdUserType_Click()
    lblDisplay.Caption = ""
    lblDisplay.Caption = Format(8972.234, "0") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(8972.234, "0.0") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(6648972.26565, "0.00") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(6648972.265, "#,#0.00") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(6648972.265, "$#,#0.00") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(0.576, "0%") & vbNewLine
    lblDisplay.Caption = lblDisplay.Caption & Format(0.5766, "0.00%") & vbNewLine
End Sub
