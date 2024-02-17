VERSION 5.00
Begin VB.Form start 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6330
      Left            =   915
      Picture         =   "start.frx":0000
      ScaleHeight     =   6270
      ScaleWidth      =   9360
      TabIndex        =   0
      Top             =   180
      Width           =   9420
      Begin VB.CommandButton start 
         Caption         =   "START"
         Height          =   810
         Left            =   3270
         TabIndex        =   1
         Top             =   4935
         Width           =   2880
      End
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub start_Click()
 jogo.Show
End Sub
