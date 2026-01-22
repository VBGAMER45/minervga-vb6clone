VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MinerVGA"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAbout 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   394
      TabIndex        =   0
      Top             =   0
      Width           =   5970
   End
End
Attribute VB_Name = "frmAbout"
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
' MinerVGA - About Form
' ============================================================================

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1

' Link areas for click detection
Private Type LinkArea
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
    URL As String
End Type

Private Links(1 To 2) As LinkArea

Private Sub Form_Load()
    Call DrawAboutScreen
End Sub

Private Sub DrawAboutScreen()
    Dim Y As Integer

    picAbout.Cls
    picAbout.FontName = "Consolas"

    Y = 20

    ' Title with border
    picAbout.FillStyle = 0  ' Solid
    picAbout.FillColor = &H404040
    picAbout.Line (50, Y - 5)-(350, Y + 35), &H606060, B
    picAbout.ForeColor = vbYellow
    picAbout.FontSize = 14
    picAbout.FontBold = True
    picAbout.CurrentX = 110
    picAbout.CurrentY = Y + 5
    picAbout.Print "MinerVGA VB6"
    Y = Y + 60

    ' Author
    picAbout.FontSize = 10
    picAbout.FontBold = False
    picAbout.ForeColor = vbWhite
    picAbout.CurrentX = 80
    picAbout.CurrentY = Y
    picAbout.Print "by Jonathan Valentin"
    Y = Y + 40

    ' Original game credit
    picAbout.ForeColor = vbCyan
    picAbout.FontSize = 9
    picAbout.CurrentX = 40
    picAbout.CurrentY = Y
    picAbout.Print "Based on the original game by"
    Y = Y + 20
    picAbout.CurrentX = 50
    picAbout.CurrentY = Y
    picAbout.Print "Lordo Frodo (Harrell W. Stiles)"
    Y = Y + 50

    ' Separator line
    picAbout.Line (30, Y)-(370, Y), &H808080
    Y = Y + 20

    ' Links section
    picAbout.ForeColor = vbYellow
    picAbout.FontSize = 9
    picAbout.FontBold = True
    picAbout.CurrentX = 150
    picAbout.CurrentY = Y
    picAbout.Print "Links:"
    Y = Y + 30

    ' GitHub link
    picAbout.ForeColor = &HFF8080  ' Light blue (for link)
    picAbout.FontBold = False
    picAbout.FontUnderline = True
    picAbout.CurrentX = 30
    picAbout.CurrentY = Y
    picAbout.Print "GitHub Repository"
    Links(1).X1 = 30
    Links(1).Y1 = Y
    Links(1).X2 = 200
    Links(1).Y2 = Y + 18
    Links(1).URL = "https://github.com/VBGAMER45/minervga-vb6clone"
    Y = Y + 25

    ' Website link
    picAbout.CurrentX = 30
    picAbout.CurrentY = Y
    picAbout.Print "The Programming Zone"
    Links(2).X1 = 30
    Links(2).Y1 = Y
    Links(2).X2 = 200
    Links(2).Y2 = Y + 18
    Links(2).URL = "https://www.theprogrammingzone.com/"
    picAbout.FontUnderline = False
    Y = Y + 50

    ' Separator line
    picAbout.Line (30, Y)-(370, Y), &H808080
    Y = Y + 20

    picAbout.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub picAbout_Click()
    ' Check if clicked on a link
    Dim i As Integer
    Dim ClickX As Integer, ClickY As Integer

    ' Get click position relative to picAbout
    ClickX = (picAbout.MousePointer + 1)  ' This won't work, need different approach
End Sub

Private Sub picAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer

    ' Check if clicked on a link
    For i = 1 To 2
        If X >= Links(i).X1 And X <= Links(i).X2 And _
           Y >= Links(i).Y1 And Y <= Links(i).Y2 Then
            ' Open URL in browser
            Call ShellExecute(Me.hwnd, "open", Links(i).URL, vbNullString, vbNullString, SW_SHOWNORMAL)
            Exit For
        End If
    Next i
End Sub

Private Sub picAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim OnLink As Boolean

    OnLink = False

    ' Check if hovering over a link
    For i = 1 To 2
        If X >= Links(i).X1 And X <= Links(i).X2 And _
           Y >= Links(i).Y1 And Y <= Links(i).Y2 Then
            OnLink = True
            Exit For
        End If
    Next i

    ' Change cursor
    If OnLink Then
        picAbout.MousePointer = 99  ' Hand cursor (or use 10 for up arrow)
    Else
        picAbout.MousePointer = 0   ' Default
    End If
End Sub
