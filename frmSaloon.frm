VERSION 5.00
Begin VB.Form frmSaloon
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Miss Mimi's Saloon"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit
      Caption         =   "&Leave Saloon"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdMimi
      Caption         =   "Visit Miss &Mimi"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdNight
      Caption         =   "&Night to Remember ($50)"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdFood
      Caption         =   "Buy &Food ($10)"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdBeer
      Caption         =   "Buy &Beer ($5)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraStatus
      BackColor       =   &H00400040&
      Caption         =   "Your Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Label lblCash
         BackColor       =   &H00400040&
         Caption         =   "Cash: $0"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblRing
         BackColor       =   &H00400040&
         Caption         =   "Ring: No"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4935
      End
   End
   Begin VB.Frame fraProgress
      BackColor       =   &H00400040&
      Caption         =   "Progress to Win"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5175
      Begin VB.Label lblProgress
         BackColor       =   &H00400040&
         Caption         =   "Progress"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Label lblHint
      BackColor       =   &H00400040&
      Caption         =   ""
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   5175
   End
End
Attribute VB_Name = "frmSaloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - Saloon Form (Win Condition)
' ============================================================================

Private Hints(1 To 10) As String

Private Sub Form_Load()
    ' Initialize hints
    Call InitHints
    Call UpdateDisplay
End Sub

Private Sub InitHints()
    Hints(1) = "The deeper you dig, the richer the ore!"
    Hints(2) = "Platinum is worth a fortune - look deep!"
    Hints(3) = "A lantern helps you spot trouble before it's too late."
    Hints(4) = "Always have an escape route when using dynamite!"
    Hints(5) = "The elevator can be upgraded at the store."
    Hints(6) = "Cave-ins fill nearby tunnels - watch out!"
    Hints(7) = "Whirlpools flood everything nearby with water."
    Hints(8) = "A drill cuts through granite like butter."
    Hints(9) = "Miss Mimi requires $20,000 AND a ring!"
    Hints(10) = "The hospital can save your life - for a price."
End Sub

Private Sub UpdateDisplay()
    ' Update cash
    lblCash.Caption = "Cash: $" & Format(Player.Cash, "#,##0")

    ' Update ring status
    If HasRing Then
        lblRing.Caption = "Ring: Yes (Diamond Ring)"
        lblRing.ForeColor = &HFFFF00  ' Cyan
    Else
        lblRing.Caption = "Ring: No (Buy at Store - $100)"
        lblRing.ForeColor = vbWhite
    End If

    ' Update progress
    Dim ProgressText As String
    Dim CashProgress As Single
    CashProgress = (Player.Cash / WIN_MONEY) * 100
    If CashProgress > 100 Then CashProgress = 100

    ProgressText = "Money: $" & Format(Player.Cash, "#,##0") & " / $" & Format(WIN_MONEY, "#,##0")
    ProgressText = ProgressText & " (" & Format(CashProgress, "0") & "%)" & vbCrLf
    ProgressText = ProgressText & "Ring: "
    If HasRing Then
        ProgressText = ProgressText & "YES"
    Else
        ProgressText = ProgressText & "NO"
    End If

    lblProgress.Caption = ProgressText

    ' Update buttons
    cmdBeer.Enabled = (Player.Cash >= COST_BEER)
    cmdFood.Enabled = (Player.Cash >= COST_FOOD)
    cmdNight.Enabled = (Player.Cash >= COST_NIGHT)

    ' Mimi is available if you have some money
    cmdMimi.Enabled = (Player.Cash >= 1000)  ' Need at least $1000 to talk to her
End Sub

Private Sub cmdBeer_Click()
    If Player.Cash >= COST_BEER Then
        Player.Cash = Player.Cash - COST_BEER
        MsgBox "You enjoy a cold beer. Refreshing!", vbInformation, "Saloon"
        Call UpdateDisplay
    End If
End Sub

Private Sub cmdFood_Click()
    If Player.Cash >= COST_FOOD Then
        Player.Cash = Player.Cash - COST_FOOD
        ' Food heals a tiny bit
        If Player.Health < 100 Then
            Player.Health = Player.Health + 2
            If Player.Health > 100 Then Player.Health = 100
            MsgBox "You enjoy a hearty meal. You feel a bit better! (+2 health)", vbInformation, "Saloon"
        Else
            MsgBox "You enjoy a hearty meal.", vbInformation, "Saloon"
        End If
        Call UpdateDisplay
    End If
End Sub

Private Sub cmdNight_Click()
    If Player.Cash >= COST_NIGHT Then
        Player.Cash = Player.Cash - COST_NIGHT

        ' Show a random hint
        Dim HintNum As Integer
        HintNum = Int(Rnd * 10) + 1
        lblHint.Caption = "Tip: " & Hints(HintNum)

        MsgBox "You spend a memorable evening..." & vbCrLf & vbCrLf & _
               "One of the girls whispers:" & vbCrLf & _
               Chr(34) & Hints(HintNum) & Chr(34), vbInformation, "Saloon"

        Call UpdateDisplay
    End If
End Sub

Private Sub cmdMimi_Click()
    ' Check win condition
    If CheckWinCondition() Then
        ' Player wins!
        Dim WinMsg As String
        WinMsg = "Miss Mimi's eyes light up as she sees the diamond ring!" & vbCrLf & vbCrLf
        WinMsg = WinMsg & "With $" & Format(Player.Cash, "#,##0") & " and a beautiful ring," & vbCrLf
        WinMsg = WinMsg & "she agrees to marry you!" & vbCrLf & vbCrLf
        WinMsg = WinMsg & "CONGRATULATIONS! YOU WIN!"

        MsgBox WinMsg, vbInformation, "You Win!"

        GameState = STATE_WON
        Unload Me
    Else
        ' Not enough yet
        Dim Msg As String
        Msg = "Miss Mimi looks at you skeptically..." & vbCrLf & vbCrLf

        If Player.Cash < WIN_MONEY And Not HasRing Then
            Msg = Msg & Chr(34) & "Honey, you need $" & Format(WIN_MONEY, "#,##0") & " AND a diamond ring " & _
                  "before I'll even consider settling down!" & Chr(34)
        ElseIf Player.Cash < WIN_MONEY Then
            Dim Needed As Long
            Needed = WIN_MONEY - Player.Cash
            Msg = Msg & Chr(34) & "That's a lovely ring, sugar, but I need $" & Format(WIN_MONEY, "#,##0") & _
                  " for our retirement! You're $" & Format(Needed, "#,##0") & " short!" & Chr(34)
        Else
            Msg = Msg & Chr(34) & "You've got the money, but where's my ring?! " & _
                  "Buy one at the general store!" & Chr(34)
        End If

        MsgBox Msg, vbExclamation, "Miss Mimi"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
