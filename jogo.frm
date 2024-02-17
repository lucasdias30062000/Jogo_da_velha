VERSION 5.00
Begin VB.Form jogo 
   Caption         =   "Form2"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11460
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5670
      Index           =   0
      Left            =   1485
      Picture         =   "jogo.frx":0000
      ScaleHeight     =   5610
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.PictureBox Picture2 
         Height          =   3780
         Left            =   2025
         Picture         =   "jogo.frx":B169
         ScaleHeight     =   3720
         ScaleWidth      =   4425
         TabIndex        =   1
         Top             =   960
         Width           =   4485
         Begin VB.TextBox C3_L3 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Index           =   8
            Left            =   2985
            TabIndex        =   10
            Top             =   2565
            Width           =   1050
         End
         Begin VB.TextBox C2_L3 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   7
            Left            =   1755
            TabIndex        =   9
            Top             =   2550
            Width           =   1080
         End
         Begin VB.TextBox C1_L3 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   6
            Left            =   510
            TabIndex        =   8
            Top             =   2550
            Width           =   1095
         End
         Begin VB.TextBox C3_L2 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   5
            Left            =   2985
            TabIndex        =   7
            Top             =   1395
            Width           =   1050
         End
         Begin VB.TextBox C2_L2 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   4
            Left            =   1785
            TabIndex        =   6
            Top             =   1380
            Width           =   1050
         End
         Begin VB.TextBox C1_L2 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   3
            Left            =   480
            TabIndex        =   5
            Top             =   1380
            Width           =   1140
         End
         Begin VB.TextBox C3_L1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   2
            Left            =   2985
            TabIndex        =   4
            Top             =   285
            Width           =   1050
         End
         Begin VB.TextBox C2_L1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   1
            Left            =   1785
            TabIndex        =   3
            Top             =   270
            Width           =   1050
         End
         Begin VB.TextBox C1_L1 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Index           =   0
            Left            =   480
            TabIndex        =   2
            Top             =   270
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "jogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'x | x | x'
'
'
Private Sub C1_L1_change(Index As Integer)
 If ((C1_L1(0) <> "0") And (C1_L1(0) <> "X") And (C1_L1(0) <> "x") And (C1_L1(0) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C1_L1(0) <> "" And C1_L1(0) = C2_L1(1) And C2_L1(1) = C3_L1(2)) And ((C1_L1(0) = "0") Or (C1_L1(0) = "X") Or (C1_L1(0) = "x")) Then
    MsgBox "Jogador " & C1_L1(0) & " Ganhou"
    End If
'''''''''''''''''''''''''''''''''
    'x | |
    'x | |
    'x | |
  If (C1_L1(0) <> "" And C1_L1(0) = C1_L2(3) And C1_L1(0) = C1_L3(6)) And ((C1_L1(0) = "0") Or (C1_L1(0) = "X") Or (C1_L1(0) = "x")) Then
    MsgBox "O Ganhador é " & C1_L1(0)
    End If
''''''''''''''''''''''''''''''''''''
    'x |   |
    '  | x |
    '  |   | x
  If (C1_L1(0) <> "" And C1_L1(0) = C2_L2(4) And C1_L1(0) = C3_L3(8)) And ((C1_L1(0) = "0") Or (C1_L1(0) = "X") Or (C1_L1(0) = "x")) Then
    MsgBox "O Ganhador é " & C1_L1(0)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''''''''''''
'x | x | x'
'
'

Private Sub C2_L1_change(Index As Integer)
 If ((C2_L1(1) <> "0") And (C2_L1(1) <> "X") And (C2_L1(1) <> "x") And (C2_L1(1) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
    If (C2_L1(1) <> "" And C2_L1(1) = C1_L1(0) And C2_L1(1) = C3_L1(2)) And ((C2_L1(1) = "0") Or (C2_L1(1) = "X") Or (C2_L1(1) = "x")) Then
    MsgBox "Jogador " & C2_L1(1) & " Ganhou"
    End If
''''''''''''''''''''''''''
    ' | x |
    ' | x |
    ' | x |
  If (C2_L1(1) <> "" And C2_L1(1) = C2_L2(4) And C2_L1(1) = C2_L3(7)) And ((C2_L1(1) = "0") Or (C2_L1(1) = "X") Or (C2_L1(1) = "x")) Then
    MsgBox "O Ganhador é " & C2_L1(1)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''
'x | x | x'
'
'
Private Sub C3_L1_change(Index As Integer)
 If ((C3_L1(2) <> "0") And (C3_L1(2) <> "X") And (C3_L1(2) <> "x") And (C3_L1(2) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C3_L1(2) <> "" And C3_L1(2) = C2_L1(1) And C3_L1(2) = C1_L1(0)) And ((C3_L1(2) = "0") Or (C3_L1(2) = "X") Or (C3_L1(2) = "x")) Then
    MsgBox "O Ganhador é " & C3_L1(2)
    End If
''''''''''''''''''''''''''''''''''''''''
' |  | x
' |  | x
' |  | x
 If (C3_L1(2) <> "" And C3_L1(2) = C3_L2(5) And C3_L1(2) = C3_L3(8)) And ((C3_L1(2) = "0") Or (C3_L1(2) = "X") Or (C3_L1(2) = "x")) Then
    MsgBox "O Ganhador é " & C3_L1(2)
    Else
    Call verificar_empate
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''
    '
    'x | x | x'
    '
Private Sub C1_L2_change(Index As Integer)
 If ((C1_L2(3) <> "0") And (C1_L2(3) <> "X") And (C1_L2(3) <> "x") And (C1_L2(3) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C1_L2(3) <> "" And C1_L2(3) = C2_L2(4) And C1_L2(3) = C3_L2(5)) And ((C1_L2(3) = "0") Or (C1_L2(3) = "X") Or (C1_L2(3) = "x")) Then
    MsgBox "O Ganhador é " & C1_L2(3)
    End If
'''''''''''''''''''''''''''''''''
    'x | |
    'x | |
    'x | |
  If (C1_L2(3) <> "" And C1_L2(3) = C1_L1(0) And C1_L2(3) = C1_L3(6)) And ((C1_L2(3) = "0") Or (C1_L2(3) = "X") Or (C1_L2(3) = "x")) Then
    MsgBox "O Ganhador é " & C1_L2(3)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''
    '
    'x | x | x'
    '
Private Sub C2_L2_change(Index As Integer)
 If ((C2_L2(4) <> "0") And (C2_L2(4) <> "X") And (C2_L2(4) <> "x") And (C2_L2(4) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C2_L2(4) <> "" And C2_L2(4) = C1_L2(3) And C2_L2(4) = C3_L2(5)) And ((C2_L2(4) = "0") Or (C2_L2(4) = "X") Or (C2_L2(4) = "x")) Then
    MsgBox "O Ganhador é " & C2_L2(4)
    End If
'''''''''''''''''''''''''''''''''''
    '  | x |
    '  | x |
    '  | x |
  If (C2_L2(4) <> "" And C2_L2(4) = C2_L1(1) And C2_L2(4) = C2_L3(7)) And ((C2_L2(4) = "0") Or (C2_L2(4) = "X") Or (C2_L2(4) = "x")) Then
    MsgBox "O Ganhador é " & C2_L2(4)
    End If
''''''''''''''''''''''''''''''''''
    ' x |   |
    '   | x |
    '   |   | x
  If (C2_L2(4) <> "" And C2_L2(4) = C1_L1(0) And C2_L2(4) = C3_L3(8)) And ((C2_L2(4) = "0") Or (C2_L2(4) = "X") Or (C2_L2(4) = "x")) Then
    MsgBox "O Ganhador é " & C2_L2(4)
    End If
''''''''''''''''''''''''''''''''''
    '   |   | x
    '   | x |
    ' x |   |
  If (C2_L2(4) <> "" And C2_L2(4) = C3_L1(2) And C2_L2(4) = C1_L3(6)) And ((C2_L2(4) = "0") Or (C2_L2(4) = "X") Or (C2_L2(4) = "x")) Then
    MsgBox "O Ganhador é " & C2_L2(4)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''''''''''
    '
    'x | x | x'
    '
Private Sub C3_L2_change(Index As Integer)
 If ((C3_L2(5) <> "0") And (C3_L2(5) <> "X") And (C3_L2(5) <> "x") And (C3_L2(5) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C3_L2(5) <> "" And C3_L2(5) = C2_L2(4) And C3_L2(5) = C1_L2(3)) And ((C3_L2(5) = "0") Or (C3_L2(5) = "X") Or (C3_L2(5) = "x")) Then
    MsgBox "O Ganhador é " & C3_L2(5)
    End If
'''''''''''''''''''''''''''''''''
    '  | | x
    '  | | x
    '  | | x
  If (C3_L2(5) <> "" And C3_L2(5) = C3_L1(2) And C3_L2(5) = C3_L3(8)) And ((C3_L2(5) = "0") Or (C3_L2(5) = "X") Or (C3_L2(5) = "x")) Then
    MsgBox "O Ganhador é " & C3_L2(5)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''
    '  |   |
    '  |   |
    'x | x | x
Private Sub C1_L3_change(Index As Integer)
 If ((C1_L3(6) <> "0") And (C1_L3(6) <> "X") And (C1_L3(6) <> "x") And (C1_L3(6) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
 If (C1_L3(6) <> "" And C1_L3(6) = C2_L3(7) And C1_L3(6) = C3_L3(8)) And ((C1_L3(6) = "0") Or (C1_L3(6) = "X") Or (C1_L3(6) = "x")) Then
    MsgBox "O Ganhador é " & C1_L3(6)
    End If
'''''''''''''''''''''''''''''''''
    'x | |
    'x | |
    'x | |
  If (C1_L3(6) <> "" And C1_L3(6) = C1_L2(3) And C1_L3(6) = C1_L1(0)) And ((C1_L3(6) = "0") Or (C1_L3(6) = "X") Or (C1_L3(6) = "x")) Then
    MsgBox "O Ganhador é " & C1_L3(6)
    End If
''''''''''''''''''''''''''''''''''
    '   |   | x
    '   | x |
    ' x |   |
  If (C1_L3(6) <> "" And C1_L3(6) = C2_L2(4) And C1_L3(6) = C3_L1(2)) And ((C1_L3(6) = "0") Or (C1_L3(6) = "X") Or (C1_L3(6) = "x")) Then
    MsgBox "O Ganhador é " & C1_L3(6)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''''''''''''
    '
    '
    'x | x | x
Private Sub C2_L3_change(Index As Integer)
 If ((C2_L3(7) <> "0") And (C2_L3(7) <> "X") And (C2_L3(7) <> "x") And (C2_L3(7) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
  If (C2_L3(7) <> "" And C2_L3(7) = C1_L3(6) And C2_L3(7) = C3_L3(8)) And ((C2_L3(7) = "0") Or (C2_L3(7) = "X") Or (C2_L3(7) = "x")) Then
    MsgBox "O Ganhador é " & C3_L1(2)
    End If
''''''''''''''''''''''''''''''''''''''''''
    '  | x |
    '  | x |
    '  | x |
  If (C2_L3(7) <> "" And C2_L3(7) = C2_L2(4) And C2_L3(7) = C2_L1(1)) And ((C2_L3(7) = "0") Or (C2_L3(7) = "X") Or (C2_L3(7) = "x")) Then
    MsgBox "O Ganhador é " & C2_L3(7)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''''''''''
    '  |   |
    '  |   |
    'x | x | x
Private Sub C3_L3_change(Index As Integer)
 If ((C3_L3(8) <> "0") And (C3_L3(8) <> "X") And (C3_L3(8) <> "x") And (C3_L3(8) <> BS)) Then
 MsgBox "Digite X ou 0"
 End If
  If (C3_L3(8) <> "" And C3_L3(8) = C2_L3(7) And C3_L3(8) = C1_L3(6)) And ((C3_L3(8) = "0") Or (C3_L3(8) = "X") Or (C3_L3(8) = "x")) Then
    MsgBox "O Ganhador é " & C3_L3(8)
    End If
''''''''''''''''''''''''''''''''''''''''''
    '  |   | x
    '  |   | x
    '  |   | x
  If (C3_L3(8) <> "" And C3_L3(8) = C3_L2(5) And C3_L3(8) = C3_L1(2)) And ((C3_L3(8) = "0") Or (C3_L3(8) = "X") Or (C3_L3(8) = "x")) Then
    MsgBox "O Ganhador é " & C3_L3(8)
    End If
''''''''''''''''''''''''''''''''''''''''''
    ' x |   |
    '   | x |
    '   |   | x
    
  If (C3_L3(8) <> "" And C3_L3(8) = C2_L2(4) And C3_L3(8) = C1_L1(0)) And ((C3_L3(8) = "0") Or (C3_L3(8) = "X") Or (C3_L3(8) = "x")) Then
    MsgBox "O Ganhador é " & C3_L3(8)
    Else
    Call verificar_empate
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''

Private Sub verificar_empate()
 If C1_L1(0) <> "" And C2_L1(1) <> "" And C3_L1(2) <> "" And _
    C1_L2(3) <> "" And C2_L2(4) <> "" And C3_L2(5) <> "" And _
    C1_L3(6) <> "" And C2_L3(7) <> "" And C3_L3(8) <> "" Then
    MsgBox "Empate"
    End
 End If
End Sub

