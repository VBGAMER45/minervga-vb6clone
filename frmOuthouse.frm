VERSION 5.00
Begin VB.Form frmOuthouse 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outhouse"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOuthouse 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   314
      TabIndex        =   0
      Top             =   0
      Width           =   4770
   End
End
Attribute VB_Name = "frmOuthouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' MinerVGA VB6 Edition by vbgamer45
' https://github.com/VBGAMER45/minervga-vb6clone
' https://www.theprogrammingzone.com/
' ============================================================================
Option Explicit

' ============================================================================
' MinerVGA - Outhouse Form (Easter egg building)
' ============================================================================

Private Rumors(0 To 9) As String

Private Sub Form_Load()
    ' Initialize rumors/hints
    Rumors(0) = "The deeper you go, the better the ore!"
    Rumors(1) = "Platinum is worth a fortune at the bank."
    Rumors(2) = "Watch out for springs - they flood fast!"
    Rumors(3) = "The clover brings luck to those who find it."
    Rumors(4) = "Miss Mimi only wants wealthy suitors..."
    Rumors(5) = "Volcanic rock is tough but valuable."
    Rumors(6) = "The elevator can be upgraded at the store."
    Rumors(7) = "A torch is needed to light dynamite."
    Rumors(8) = "Diamonds are a girl's best friend."
    Rumors(9) = "Don't forget to visit the bank!"

    Call DrawOuthouse
End Sub

Private Sub DrawOuthouse()
    Dim RumorIdx As Integer

    picOuthouse.Cls
    picOuthouse.ForeColor = vbWhite
    picOuthouse.FontName = "Consolas"
    picOuthouse.FontSize = 10
    picOuthouse.FontBold = True

    ' Title
    picOuthouse.CurrentX = 100
    picOuthouse.CurrentY = 20
    picOuthouse.Print "OUTHOUSE"

    picOuthouse.FontBold = False
    picOuthouse.FontSize = 8

    ' Description
    picOuthouse.ForeColor = &H808080
    picOuthouse.CurrentX = 20
    picOuthouse.CurrentY = 60
    picOuthouse.Print "You step into the small wooden"
    picOuthouse.CurrentX = 20
    picOuthouse.CurrentY = 75
    picOuthouse.Print "structure. It smells... interesting."

    ' Random rumor from the wall
    picOuthouse.ForeColor = vbYellow
    picOuthouse.CurrentX = 20
    picOuthouse.CurrentY = 110
    picOuthouse.Print "Scratched on the wall:"

    RumorIdx = Int(Rnd * 10)
    picOuthouse.ForeColor = vbCyan
    picOuthouse.CurrentX = 20
    picOuthouse.CurrentY = 130
    picOuthouse.Print Chr(34) & Rumors(RumorIdx) & Chr(34)

    ' Instructions
    picOuthouse.ForeColor = vbGreen
    picOuthouse.CurrentX = 60
    picOuthouse.CurrentY = 180
    picOuthouse.Print "Press any key to leave"

    picOuthouse.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub picOuthouse_Click()
    Unload Me
End Sub
