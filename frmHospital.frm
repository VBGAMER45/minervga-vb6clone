VERSION 5.00
Begin VB.Form frmHospital
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "St.Woody Memorial Hospital"
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
   Begin VB.PictureBox picHospital
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   506
      TabIndex        =   0
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - Hospital Form (Graphical Interface matching JS version)
' ============================================================================

Private Const COST_PER_DAY As Long = 100
Private Const SURGERY_COST As Long = 300
Private Const HEALTH_PER_DAY As Integer = 10

Private Sub Form_Load()
    Call DrawHospitalInterface
End Sub

Private Sub DrawHospitalInterface()
    Dim Y As Integer
    Dim MissingHealth As Integer
    Dim DaysNeeded As Single
    Dim FullCost As Long

    picHospital.Cls
    picHospital.FontName = "Consolas"
    picHospital.FontSize = 10
    picHospital.FontBold = True

    Y = 20

    ' Draw title box
    picHospital.ForeColor = vbCyan
    picHospital.Line (100, Y)-(410, Y + 25), vbCyan, B
    picHospital.CurrentX = 115
    picHospital.CurrentY = Y + 5
    picHospital.Print "St.Woody Memorial Hospital"
    Y = Y + 50

    ' Welcome message
    picHospital.ForeColor = vbCyan
    picHospital.FontSize = 9
    picHospital.FontBold = False
    picHospital.CurrentX = 20
    picHospital.CurrentY = Y
    picHospital.Print "Welcome to St. Woody's. We are pleased to take care of you."
    Y = Y + 15
    picHospital.CurrentX = 20
    picHospital.CurrentY = Y
    picHospital.Print "In God we trust, all others pay cash (No insurance accepted)"
    Y = Y + 15
    picHospital.CurrentX = 80
    picHospital.CurrentY = Y
    picHospital.Print "What type of service do you want?"
    Y = Y + 30

    ' Press X to leave
    picHospital.ForeColor = vbYellow
    picHospital.FontBold = True
    picHospital.CurrentX = 100
    picHospital.CurrentY = Y
    picHospital.Print "Press X to Leave the Hospital"
    Y = Y + 35

    ' Calculate bedrest needed
    MissingHealth = 100 - Player.Health
    If MissingHealth < 0 Then MissingHealth = 0

    DaysNeeded = MissingHealth / HEALTH_PER_DAY
    If DaysNeeded < 0.1 Then DaysNeeded = 0.1
    FullCost = CLng(DaysNeeded * COST_PER_DAY)

    ' Surgery info
    picHospital.ForeColor = vbWhite
    picHospital.FontBold = False
    picHospital.CurrentX = 20
    picHospital.CurrentY = Y
    If Player.Health < 100 Then
        picHospital.Print "You may be in need of Surgery."
    Else
        picHospital.Print "You appear to be in good health."
    End If
    Y = Y + 18

    picHospital.CurrentX = 20
    picHospital.CurrentY = Y
    picHospital.Print "Your bedrest will probably be about " & Format(DaysNeeded, "0.0") & " day(s)."
    Y = Y + 18

    picHospital.CurrentX = 20
    picHospital.CurrentY = Y
    picHospital.Print "Our fees are quite reasonable, $" & COST_PER_DAY & " per day."
    Y = Y + 40

    ' Options
    picHospital.ForeColor = vbGreen
    picHospital.CurrentX = 60
    picHospital.CurrentY = Y
    picHospital.Print "press A to stay until mostly healed."
    Y = Y + 25

    picHospital.CurrentX = 60
    picHospital.CurrentY = Y
    picHospital.Print "press D to stay one day and night."
    Y = Y + 25

    picHospital.CurrentX = 60
    picHospital.CurrentY = Y
    picHospital.Print "press S for Surgical procedures ($" & SURGERY_COST & ")."

    picHospital.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim MissingHealth As Integer
    Dim DaysNeeded As Single
    Dim FullCost As Long
    Dim HealAmount As Integer

    MissingHealth = 100 - Player.Health
    If MissingHealth < 0 Then MissingHealth = 0
    DaysNeeded = MissingHealth / HEALTH_PER_DAY
    FullCost = CLng(DaysNeeded * COST_PER_DAY)

    Select Case KeyCode
        Case vbKeyX, vbKeyEscape
            Unload Me

        Case vbKeyA  ' Stay until mostly healed
            If Player.Health >= 100 Then
                Call AddMessage("Already healthy!")
            ElseIf Player.Cash < FullCost Then
                Call AddMessage("Need $" & FullCost)
            Else
                Player.Cash = Player.Cash - FullCost
                Call HealPlayer(MissingHealth)
                Call AddMessage("Fully healed!")
                Call PlayPurchaseSound
            End If
            Call DrawHospitalInterface

        Case vbKeyD  ' Stay one day
            If Player.Health >= 100 Then
                Call AddMessage("Already healthy!")
            ElseIf Player.Cash < COST_PER_DAY Then
                Call AddMessage("Need $" & COST_PER_DAY)
            Else
                Player.Cash = Player.Cash - COST_PER_DAY
                HealAmount = HEALTH_PER_DAY
                If Player.Health + HealAmount > 100 Then
                    HealAmount = 100 - Player.Health
                End If
                Call HealPlayer(HealAmount)
                Call AddMessage("Rested 1 day")
                Call PlayPurchaseSound
            End If
            Call DrawHospitalInterface

        Case vbKeyS  ' Surgery
            If Player.Health >= 100 Then
                Call AddMessage("Already healthy!")
            ElseIf Player.Cash < SURGERY_COST Then
                Call AddMessage("Need $" & SURGERY_COST)
            Else
                Player.Cash = Player.Cash - SURGERY_COST
                ' Surgery heals 50% of missing health
                HealAmount = MissingHealth \ 2
                If HealAmount < 10 Then HealAmount = 10
                If Player.Health + HealAmount > 100 Then
                    HealAmount = 100 - Player.Health
                End If
                Call HealPlayer(HealAmount)
                Call AddMessage("Surgery done!")
                Call PlayPurchaseSound
            End If
            Call DrawHospitalInterface
    End Select
End Sub

Private Sub picHospital_Click()
    ' Allow clicking to dismiss (optional)
End Sub
