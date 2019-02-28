VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7092
   ScaleWidth      =   16320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calc"
      Height          =   732
      Left            =   10800
      TabIndex        =   3
      Top             =   1560
      Width           =   1092
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   852
      Left            =   10860
      TabIndex        =   8
      Top             =   2700
      Width           =   1152
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   792
      Left            =   11040
      TabIndex        =   7
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox txtC 
      Height          =   732
      Left            =   5220
      TabIndex        =   2
      Top             =   420
      Width           =   1392
   End
   Begin VB.TextBox txtB 
      Height          =   792
      Left            =   2760
      TabIndex        =   1
      Top             =   420
      Width           =   1212
   End
   Begin VB.TextBox txtA 
      Height          =   732
      Left            =   540
      TabIndex        =   0
      Top             =   420
      Width           =   1152
   End
   Begin VB.Image imgAcute 
      Height          =   2112
      Left            =   3480
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.Label Label 
      Caption         =   "PERIMETER"
      Height          =   312
      Index           =   4
      Left            =   6360
      TabIndex        =   14
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label lblPerimeter 
      BorderStyle     =   1  'Fixed Single
      Height          =   732
      Left            =   6240
      TabIndex        =   13
      Top             =   2340
      Width           =   1752
   End
   Begin VB.Label Label 
      Caption         =   "AREA"
      Height          =   252
      Index           =   3
      Left            =   4080
      TabIndex        =   12
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label lblArea 
      BorderStyle     =   1  'Fixed Single
      Height          =   732
      Left            =   3960
      TabIndex        =   11
      Top             =   2340
      Width           =   1692
   End
   Begin VB.Label lblIsNot 
      Caption         =   "THIS IS NOT A TRIANGLE"
      Height          =   1032
      Left            =   780
      TabIndex        =   10
      Top             =   4140
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label lblIs 
      Caption         =   "THIS IS A TRIANGLE"
      Height          =   792
      Left            =   840
      TabIndex        =   9
      Top             =   2460
      Visible         =   0   'False
      Width           =   2112
   End
   Begin VB.Label Label 
      Caption         =   "side c"
      Height          =   552
      Index           =   2
      Left            =   5460
      TabIndex        =   6
      Top             =   1560
      Width           =   2232
   End
   Begin VB.Label Label 
      Caption         =   "side b"
      Height          =   492
      Index           =   1
      Left            =   3060
      TabIndex        =   5
      Top             =   1380
      Width           =   1452
   End
   Begin VB.Label Label 
      Caption         =   "side a"
      Height          =   552
      Index           =   0
      Left            =   660
      TabIndex        =   4
      Top             =   1440
      Width           =   1272
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Single
Dim b As Single
Dim c As Single




Private Sub cmdCalculate_Click()
Dim a As Single
Dim b As Single
Dim c As Single
Dim s As Single

a = Val(txtA)
b = Val(txtB)
c = Val(txtC)



If a + b > c And b + c > a And a + c > b Then

lblIs.Visible = True
 lblPerimeter = a + b + c
 s = (a + b + c) / 2
 lblArea = Sqr(s * (s - a) * (s - b) * (s - c))
        If a ^ 2 + b ^ 2 > c ^ 2 Then
        imgAcute.Visible = True
        End If

Else
lblIsNot.Visible = True

End If




cmdClear.SetFocus

End Sub

Private Sub cmdCalculate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdClear.SetFocus
End If
End Sub

Private Sub cmdClearKeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdExit.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
txtA = " "
txtB = " "
txtC = " "
lblIs.Visible = False
lblIsNot.Visible = False

End Sub

Private Sub cmdExit_Click()
End

End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtC.SetFocus
End If

End Sub

Private Sub txtC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdCalculate.SetFocus
End If

End Sub
