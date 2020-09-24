VERSION 5.00
Begin VB.Form fDemo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "IAMC"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txOper2 
      Height          =   2670
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   8
      Text            =   "fDemo.frx":0000
      Top             =   345
      Width           =   3060
   End
   Begin VB.TextBox txOper1 
      Height          =   2670
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   7
      Text            =   "fDemo.frx":0276
      Top             =   345
      Width           =   3060
   End
   Begin VB.TextBox txResult 
      Height          =   2610
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      Top             =   3390
      Width           =   7800
   End
   Begin VB.CommandButton btRoot 
      Caption         =   "Square Root"
      Height          =   390
      Left            =   3630
      TabIndex        =   5
      Top             =   2655
      Width           =   1140
   End
   Begin VB.CommandButton btMod 
      Caption         =   "Mod"
      Height          =   390
      Left            =   3630
      TabIndex        =   4
      Top             =   2190
      Width           =   1140
   End
   Begin VB.CommandButton btDivide 
      Caption         =   "Divide"
      Height          =   390
      Left            =   3630
      TabIndex        =   3
      Top             =   1710
      Width           =   1140
   End
   Begin VB.CommandButton btMultiply 
      Caption         =   "Multiply"
      Height          =   390
      Left            =   3630
      TabIndex        =   2
      Top             =   1245
      Width           =   1140
   End
   Begin VB.CommandButton btSubtract 
      Caption         =   "Subtract"
      Height          =   390
      Left            =   3630
      TabIndex        =   1
      Top             =   780
      Width           =   1140
   End
   Begin VB.CommandButton btadd 
      Caption         =   "Add"
      Height          =   390
      Left            =   3630
      TabIndex        =   0
      Top             =   315
      Width           =   1140
   End
   Begin VB.Label lb 
      Caption         =   "Operand 2 (or Estimate) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   2085
   End
   Begin VB.Label lb 
      Caption         =   "Operand 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   10
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lb 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   9
      Top             =   3150
      Width           =   555
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Oper1()     As Byte 'operand1
Private Oper2()     As Byte 'operand2
Private Reslt()     As Byte 'result
Private Remai()     As Byte 'remainder
Private tmp()       As Byte
Private cAri        As cIAAM

Private Sub btadd_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.Add Oper1, Oper2, Reslt
    txResult = FromOper(Reslt)

End Sub

Private Sub btDivide_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.Divide Oper1, Oper2, Reslt, Remai
    txResult = FromOper(Reslt)

End Sub

Private Sub btMod_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.Divide Oper1, Oper2, Reslt, Remai
    txResult = FromOper(Remai)

End Sub

Private Sub btMultiply_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.Multiply Oper1, Oper2, Reslt
    txResult = FromOper(Reslt)

End Sub

Private Sub btRoot_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.SquareRoot Oper1, Oper2
    txResult = FromOper(Oper2)

End Sub

Private Sub btSubtract_Click()

    ToOper txOper1, Oper1
    ToOper txOper2, Oper2
    cAri.Subtract Oper1, Oper2, Reslt
    txResult = FromOper(Reslt)

End Sub

Private Sub Form_Load()

    Set cAri = New cIAAM

End Sub

Private Function FromOper(Arr() As Byte) As String

  Dim i As Long

    For i = UBound(Arr) To 0 Step -1
        FromOper = FromOper & CStr(Arr(i) And 15)
        If Arr(i) > 9 Then
            FromOper = FromOper & "-"
        End If
    Next i
    Enabled = True
    Screen.MousePointer = vbNormal

End Function

Private Sub ToOper(Text As String, Arr() As Byte)

  Dim i     As Long
  Dim j     As Long
  Dim Sign  As Byte

    Screen.MousePointer = vbHourglass

    Enabled = False
    If Len(Text) Then
        j = Len(Text)
        ReDim Arr(0 To j - 1)
        For i = j To 1 Step -1
            If Mid$(Text, i, 1) = "-" Then
                Sign = &HF0
              Else 'NOT MID$(TEXT,...
                Arr(j - i) = Val(Mid$(Text, i, 1))
            End If
        Next i
        Arr(0) = Arr(0) Or Sign
        If Sign Then
            ReDim Preserve Arr(0 To j - 2)
        End If
      Else 'LEN(TEXT) = FALSE/0
        ReDim Arr(0 To 0)
    End If

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Mai-03 00:13)  Decl: 7  Code: 107  Total: 114 Lines
':) CommentOnly: 0 (0%)  Commented: 6 (5,3%)  Empty: 30 (26,3%)  Max Logic Depth: 4
