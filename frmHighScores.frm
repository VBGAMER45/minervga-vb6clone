VERSION 5.00
Begin VB.Form frmHighScores
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MinerVGA Hall of Fame"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picScores
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   506
      TabIndex        =   0
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - High Scores Form
' ============================================================================

Private Const MAX_SCORES As Integer = 10

Private Sub Form_Load()
    Call DrawHighScores
End Sub

Private Sub DrawHighScores()
    Dim Y As Integer
    Dim i As Integer

    picScores.Cls
    picScores.FontName = "Consolas"
    picScores.FontSize = 12
    picScores.FontBold = True

    Y = 20

    ' Title
    picScores.ForeColor = vbYellow
    picScores.Line (120, Y)-(390, Y + 30), vbYellow, B
    picScores.CurrentX = 150
    picScores.CurrentY = Y + 5
    picScores.Print "MinerVGA Hall of Fame"
    Y = Y + 60

    ' Column headers
    picScores.ForeColor = vbCyan
    picScores.FontSize = 10
    picScores.CurrentX = 40
    picScores.CurrentY = Y
    picScores.Print "Rank"
    picScores.CurrentX = 100
    picScores.CurrentY = Y
    picScores.Print "Name"
    picScores.CurrentX = 280
    picScores.CurrentY = Y
    picScores.Print "Score"
    picScores.CurrentX = 380
    picScores.CurrentY = Y
    picScores.Print "Date"
    Y = Y + 25

    ' Draw line
    picScores.Line (30, Y)-(480, Y), vbCyan
    Y = Y + 10

    ' Draw scores
    picScores.FontBold = False
    For i = 1 To MAX_SCORES
        If i <= 3 Then
            ' Gold, Silver, Bronze
            Select Case i
                Case 1: picScores.ForeColor = &HD7FF&    ' Gold
                Case 2: picScores.ForeColor = &HC0C0C0   ' Silver
                Case 3: picScores.ForeColor = &H507FFF   ' Bronze
            End Select
        Else
            picScores.ForeColor = vbWhite
        End If

        picScores.CurrentX = 40
        picScores.CurrentY = Y
        picScores.Print Format(i, "00") & "."

        picScores.CurrentX = 100
        picScores.CurrentY = Y
        If HighScores(i).PlayerName = "" Then
            picScores.Print "---"
        Else
            picScores.Print Left(HighScores(i).PlayerName & "               ", 15)
        End If

        picScores.CurrentX = 280
        picScores.CurrentY = Y
        If HighScores(i).Score > 0 Then
            picScores.Print "$" & Format(HighScores(i).Score, "#,##0")
        Else
            picScores.Print "$0"
        End If

        picScores.CurrentX = 380
        picScores.CurrentY = Y
        If HighScores(i).ScoreDate <> "" Then
            picScores.Print HighScores(i).ScoreDate
        Else
            picScores.Print "---"
        End If

        Y = Y + 22
    Next i

    ' Show continue instruction
    Y = Y + 15
    picScores.ForeColor = vbYellow
    picScores.CurrentX = 100
    picScores.CurrentY = Y
    picScores.Print "Press any key to continue..."

    picScores.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub picScores_Click()
    Unload Me
End Sub
