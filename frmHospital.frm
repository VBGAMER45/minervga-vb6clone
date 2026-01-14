VERSION 5.00
Begin VB.Form frmHospital
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Town Hospital"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit
      Caption         =   "&Leave Hospital"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdHealFull
      Caption         =   "Full &Recovery"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdHeal25
      Caption         =   "Heal &25%"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdHeal10
      Caption         =   "Heal &10%"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame fraStatus
      BackColor       =   &H00800000&
      Caption         =   "Patient Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label lblHealth
         BackColor       =   &H00800000&
         Caption         =   "Health: 100%"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblCash
         BackColor       =   &H00800000&
         Caption         =   "Cash: $0"
         BeginProperty Font
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.Label lblCost
      BackColor       =   &H00800000&
      Caption         =   "Cost: $5 per health point restored"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label lblInfo
      BackColor       =   &H00800000&
      Caption         =   ""
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - Hospital Form
' ============================================================================

Private Sub Form_Load()
    Call UpdateDisplay
End Sub

Private Sub UpdateDisplay()
    ' Update health display
    lblHealth.Caption = "Health: " & Player.Health & "%"

    ' Color based on health level
    If Player.Health < 20 Then
        lblHealth.ForeColor = vbRed
    ElseIf Player.Health < 50 Then
        lblHealth.ForeColor = vbYellow
    Else
        lblHealth.ForeColor = vbGreen
    End If

    ' Update cash
    lblCash.Caption = "Cash: $" & Format(Player.Cash, "#,##0")

    ' Calculate costs
    Dim MissingHealth As Integer
    Dim Cost10 As Long, Cost25 As Long, CostFull As Long

    MissingHealth = 100 - Player.Health
    If MissingHealth < 0 Then MissingHealth = 0

    Cost10 = 10 * HEAL_COST_PER_POINT
    Cost25 = 25 * HEAL_COST_PER_POINT
    CostFull = MissingHealth * HEAL_COST_PER_POINT

    ' Update button captions
    cmdHeal10.Caption = "Heal 10% ($" & Cost10 & ")"
    cmdHeal25.Caption = "Heal 25% ($" & Cost25 & ")"
    cmdHealFull.Caption = "Full Recovery ($" & CostFull & ")"

    ' Enable/disable buttons based on affordability and need
    cmdHeal10.Enabled = (Player.Cash >= Cost10) And (Player.Health < 100)
    cmdHeal25.Enabled = (Player.Cash >= Cost25) And (Player.Health < 100)
    cmdHealFull.Enabled = (Player.Cash >= CostFull) And (Player.Health < 100) And (CostFull > 0)

    ' Update cost info
    lblCost.Caption = "Cost: $" & HEAL_COST_PER_POINT & " per health point restored" & vbCrLf & _
                      "Full recovery would cost: $" & CostFull

    ' Update info
    If Player.Health >= 100 Then
        lblInfo.Caption = "You are in perfect health!"
        lblInfo.ForeColor = vbGreen
    ElseIf Player.Health < 20 Then
        lblInfo.Caption = "CRITICAL: Seek immediate treatment!"
        lblInfo.ForeColor = vbRed
    Else
        lblInfo.Caption = ""
    End If
End Sub

Private Sub cmdHeal10_Click()
    Dim Cost As Long
    Cost = 10 * HEAL_COST_PER_POINT

    If Player.Cash >= Cost Then
        Player.Cash = Player.Cash - Cost
        Call HealPlayer(10)
        MsgBox "You feel 10% better!", vbInformation, "Hospital"
        Call UpdateDisplay
    End If
End Sub

Private Sub cmdHeal25_Click()
    Dim Cost As Long
    Cost = 25 * HEAL_COST_PER_POINT

    If Player.Cash >= Cost Then
        Player.Cash = Player.Cash - Cost
        Call HealPlayer(25)
        MsgBox "You feel much better!", vbInformation, "Hospital"
        Call UpdateDisplay
    End If
End Sub

Private Sub cmdHealFull_Click()
    Dim MissingHealth As Integer
    Dim Cost As Long

    MissingHealth = 100 - Player.Health
    If MissingHealth <= 0 Then Exit Sub

    Cost = MissingHealth * HEAL_COST_PER_POINT

    If Player.Cash >= Cost Then
        Player.Cash = Player.Cash - Cost
        Call HealPlayer(MissingHealth)
        MsgBox "You are fully recovered!", vbInformation, "Hospital"
        Call UpdateDisplay
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
