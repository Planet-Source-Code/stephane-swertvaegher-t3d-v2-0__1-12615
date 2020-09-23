VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "RaiseRaise (filled)"
      Height          =   330
      Left            =   1530
      TabIndex        =   11
      Top             =   1485
      Width           =   1410
   End
   Begin VB.CommandButton Command10 
      Caption         =   "RaiseInset (filled)"
      Height          =   330
      Left            =   1530
      TabIndex        =   10
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton Command9 
      Caption         =   "None (filled)"
      Height          =   330
      Left            =   1530
      TabIndex        =   9
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton Command8 
      Caption         =   "InsetRaise (filled)"
      Height          =   330
      Left            =   1530
      TabIndex        =   8
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton Command7 
      Caption         =   "InsetInset (filled)"
      Height          =   330
      Left            =   1530
      TabIndex        =   7
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Change Color (Random)"
      Height          =   555
      Left            =   3150
      TabIndex        =   6
      Top             =   180
      Width           =   1320
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RaiseRaise"
      Height          =   330
      Left            =   45
      TabIndex        =   5
      Top             =   1485
      Width           =   1410
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RaiseInset"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "None"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "InsetRaise"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InsetInset"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1410
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   420
      Left            =   135
      TabIndex        =   12
      Top             =   3465
      Width           =   4380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "This is the new T3D"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   2430
      Width           =   2130
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dInsetInset
T3D Form1, Label2, 5, T3dInsetInset
End Sub

Private Sub Command10_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dRaiseInset, T3dF1
T3D Form1, Label2, 5, T3dRaiseInset, T3dF1
End Sub

Private Sub Command11_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dRaiseRaise, T3dF1
T3D Form1, Label2, 5, T3dRaiseRaise, T3dF1
End Sub

Private Sub Command2_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dInsetRaise
T3D Form1, Label2, 5, T3dInsetRaise
End Sub

Private Sub Command3_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dNone
T3D Form1, Label2, 5, T3dNone
End Sub

Private Sub Command4_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dRaiseInset
T3D Form1, Label2, 5, T3dRaiseInset
End Sub

Private Sub Command5_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dRaiseRaise
T3D Form1, Label2, 5, T3dRaiseRaise
End Sub

Private Sub Command6_Click()
Form1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
'you can also make the labels BackStyle = Transparent
Label1.BackColor = Form1.BackColor
Label2.BackColor = Form1.BackColor
End Sub

Private Sub Command7_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dInsetInset, T3dF1
T3D Form1, Label2, 5, T3dInsetInset, T3dF1
End Sub

Private Sub Command8_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dInsetRaise, T3dF1
T3D Form1, Label2, 5, T3dInsetRaise, T3dF1
End Sub

Private Sub Command9_Click()
Form1.Cls
T3D Form1, Label1, 5, T3dNone, T3dF1
T3D Form1, Label2, 5, T3dNone, T3dF1
End Sub

Private Sub Form_Load()
Label2.Caption = "Syntax:" & vbCr & "T3D, Form, Control, Bevel, [Style], [Filled]"
End Sub
