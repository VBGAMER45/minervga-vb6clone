VERSION 5.00
Begin VB.Form frmSaloon 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sweet Miss Mimi's Place"
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
   Begin VB.PictureBox picSaloon 
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
Attribute VB_Name = "frmSaloon"
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
' MinerVGA - Saloon Form (Miss Mimi's Hospitality House)
' ============================================================================

Private Const COST_NIGHT As Long = 50      ' Night to remember
Private Const COST_MEAL As Long = 10       ' Square meal
Private Const COST_BREW As Long = 2        ' Beer at the bar
Private Const COST_CONDOM_SALOON As Long = 5  ' Condom purchase

Private Const HEAL_MEAL As Integer = 5     ' Health from meal
Private Const HEAL_NIGHT As Integer = 15   ' Health from night's rest

Private picTileset As StdPicture
Private Rumors(1 To 10) As String
Private CurrentMessage As String
Private MessageColor As Long

Private Sub Form_Load()
    ' Load tileset for icons
    On Error Resume Next
    Set picTileset = LoadPicture(App.Path & "\tileset.bmp")
    On Error GoTo 0

    ' Initialize rumors and message
    Call InitRumors
    CurrentMessage = ""
    MessageColor = vbWhite

    Call DrawSaloonInterface
End Sub

Private Sub InitRumors()
    Rumors(1) = "Nurse Ratchett will put out for a diamond."
    Rumors(2) = "The deeper you dig, the richer the ore!"
    Rumors(3) = "Platinum is worth a fortune - look deep!"
    Rumors(4) = "Miss Mimi requires $20,000 AND a ring!"
    Rumors(5) = "The elevator man takes cash for upgrades."
    Rumors(6) = "Cave-ins fill nearby tunnels with rubble."
    Rumors(7) = "A drill cuts through granite like butter."
    Rumors(8) = "Whirlpools flood everything nearby!"
    Rumors(9) = "The bank prices change - time it right!"
    Rumors(10) = "A four-leaf clover brings good luck."
End Sub

Private Sub DrawSaloonInterface()
    Dim Y As Integer

    picSaloon.Cls
    picSaloon.FontName = "Consolas"
    picSaloon.FontSize = 10
    picSaloon.FontBold = True

    Y = 20

    ' Draw decorative icons on left (drinks)
    Call DrawSpriteIcon(60, Y, SPR_DYNAMITE)  ' Red bottle-like
    Call DrawSpriteIcon(85, Y, SPR_BUCKET)    ' White container

    ' Draw title box
    picSaloon.ForeColor = &H8080FF  ' Light red/pink
    picSaloon.Line (150, Y)-(380, Y + 25), &H8080FF, B
    picSaloon.CurrentX = 165
    picSaloon.CurrentY = Y + 5
    picSaloon.Print "Sweet Miss Mimi's Place"

    ' Draw decorative icons on right (lamp and ring)
    Call DrawSpriteIcon(400, Y, SPR_LAMP)
    Call DrawSpriteIcon(430, Y, SPR_RING)

    Y = Y + 50

    ' Welcome message
    picSaloon.ForeColor = vbWhite
    picSaloon.FontSize = 9
    picSaloon.FontBold = False
    picSaloon.CurrentX = 20
    picSaloon.CurrentY = Y
    picSaloon.Print "Welcome to Miss Mimi's Hospitality House.  The beer is cold,"
    Y = Y + 15
    picSaloon.CurrentX = 20
    picSaloon.CurrentY = Y
    picSaloon.Print "the food is warm, and the girls are hot!  Stay for the show!"
    Y = Y + 15
    picSaloon.CurrentX = 100
    picSaloon.CurrentY = Y
    picSaloon.Print "What can we do for you, cutie?"
    Y = Y + 35

    ' Press X to leave
    picSaloon.ForeColor = vbWhite
    picSaloon.FontBold = False
    picSaloon.CurrentX = 100
    picSaloon.CurrentY = Y
    picSaloon.Print "Press X to Leave the Hospitality House."
    Y = Y + 35

    ' Status message based on health
    picSaloon.ForeColor = vbWhite
    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    If Player.Health < 50 Then
        picSaloon.Print "You are in immediate need of Hospitality!"
    ElseIf Player.Health < 80 Then
        picSaloon.Print "You look like you could use some Hospitality."
    Else
        picSaloon.Print "You look healthy, but everyone needs Hospitality!"
    End If
    Y = Y + 15

    ' Prices
    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "Our fees are $" & COST_NIGHT & " per night, $" & COST_MEAL & " per meal,"
    Y = Y + 15
    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "$" & COST_BREW & " per brew, and $" & COST_CONDOM_SALOON & " for a condom."
    Y = Y + 35

    ' Options
    picSaloon.ForeColor = vbWhite
    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "press A for an audience with Mimi."
    Y = Y + 25

    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "press B to have a brew at the bar."
    Y = Y + 25

    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "press S for a night to remember."
    Y = Y + 25

    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "press M for a square meal."
    Y = Y + 25

    picSaloon.CurrentX = 60
    picSaloon.CurrentY = Y
    picSaloon.Print "press C to purchase a condom."
    Y = Y + 30

    ' Message display area (right after menu options)
    If CurrentMessage <> "" Then
        picSaloon.ForeColor = MessageColor
        picSaloon.FontBold = True
        picSaloon.CurrentX = 60
        picSaloon.CurrentY = Y
        picSaloon.Print ">> " & CurrentMessage
        picSaloon.FontBold = False
    End If
    Y = Y + 25

    ' Cash display
    picSaloon.ForeColor = vbGreen
    picSaloon.FontBold = True
    picSaloon.CurrentX = 300
    picSaloon.CurrentY = Y
    picSaloon.Print "Your Cash: $" & Format(Player.Cash, "#,##0")
    Y = Y + 20

    ' Health display
    If Player.Health < 50 Then
        picSaloon.ForeColor = vbRed
    ElseIf Player.Health < 80 Then
        picSaloon.ForeColor = vbYellow
    Else
        picSaloon.ForeColor = vbGreen
    End If
    picSaloon.CurrentX = 300
    picSaloon.CurrentY = Y
    picSaloon.Print "Health: " & Player.Health & "%"
    Y = Y + 30

    ' Rumor at bottom
    Dim RumorNum As Integer
    RumorNum = Int(Rnd * 10) + 1
    picSaloon.ForeColor = vbYellow
    picSaloon.FontBold = False
    picSaloon.CurrentX = 20
    picSaloon.CurrentY = Y
    picSaloon.Print "Rumor:" & Rumors(RumorNum)

    picSaloon.Refresh
End Sub

Private Sub DrawSpriteIcon(ByVal X As Integer, ByVal Y As Integer, ByVal SpriteIdx As Integer)
    Dim SrcX As Integer, SrcY As Integer
    Dim Col As Integer, Row As Integer

    If picTileset Is Nothing Then Exit Sub

    Col = SpriteIdx Mod 8
    Row = SpriteIdx \ 8
    SrcX = Col * CELL_WIDTH
    SrcY = Row * CELL_HEIGHT

    picSaloon.PaintPicture picTileset, X, Y, CELL_WIDTH, CELL_HEIGHT, _
                          SrcX, SrcY, CELL_WIDTH, CELL_HEIGHT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyX, vbKeyEscape
            Unload Me

        Case vbKeyA  ' Audience with Mimi (win condition check)
            Call VisitMimi

        Case vbKeyB  ' Brew at the bar
            Call HaveBrew

        Case vbKeyS  ' Night to remember
            Call SpendNight

        Case vbKeyM  ' Square meal
            Call HaveMeal

        Case vbKeyC  ' Purchase condom
            Call BuyCondom
    End Select
End Sub

Private Sub VisitMimi()
    ' Check win condition
    If CheckWinCondition() Then
        ' Player wins!
        Call AddMessage("YOU WIN!")
        Call PlayPurchaseSound
        GameState = STATE_WON
        Unload Me
    Else
        ' Not enough yet
        If Player.Cash < WIN_MONEY And Not HasRing Then
            Call SetLocalMessage("Mimi says: Need $" & Format(WIN_MONEY, "#,##0") & " AND a ring!", vbYellow)
            Call AddMessage("Need $" & Format(WIN_MONEY, "#,##0") & " + ring!")
        ElseIf Player.Cash < WIN_MONEY Then
            Dim Needed As Long
            Needed = WIN_MONEY - Player.Cash
            Call SetLocalMessage("Mimi says: Need $" & Format(Needed, "#,##0") & " more, sweetie!", vbYellow)
            Call AddMessage("Need $" & Format(Needed, "#,##0") & " more!")
        Else
            Call SetLocalMessage("Mimi says: Where's my diamond ring?!", vbYellow)
            Call AddMessage("Need a ring!")
        End If
        Call DrawSaloonInterface
    End If
End Sub

Private Sub HaveBrew()
    If Player.Cash < COST_BREW Then
        Call SetLocalMessage("Not enough cash for a brew!", vbRed)
        Call AddMessage("Need $" & COST_BREW)
        Call DrawSaloonInterface
        Exit Sub
    End If

    Player.Cash = Player.Cash - COST_BREW
    Call SetLocalMessage("*BURP!* That hit the spot!", vbGreen)
    Call AddMessage("Burp!")
    Call PlayPurchaseSound
    Call DrawSaloonInterface
End Sub

Private Sub SpendNight()
    If Player.Cash < COST_NIGHT Then
        Call SetLocalMessage("Not enough cash for a night!", vbRed)
        Call AddMessage("Need $" & COST_NIGHT)
        Call DrawSaloonInterface
        Exit Sub
    End If

    Player.Cash = Player.Cash - COST_NIGHT

    ' Heal player
    If Player.Health < 100 Then
        Call HealPlayer(HEAL_NIGHT)
        Call SetLocalMessage("What a night! You feel rested. (+" & HEAL_NIGHT & " HP)", vbGreen)
        Call AddMessage("Rested +" & HEAL_NIGHT & " HP")
    Else
        Call SetLocalMessage("What a night to remember!", vbGreen)
        Call AddMessage("What a night!")
    End If

    Call PlayPurchaseSound
    Call DrawSaloonInterface
End Sub

Private Sub HaveMeal()
    If Player.Cash < COST_MEAL Then
        Call SetLocalMessage("Not enough cash for a meal!", vbRed)
        Call AddMessage("Need $" & COST_MEAL)
        Call DrawSaloonInterface
        Exit Sub
    End If

    Player.Cash = Player.Cash - COST_MEAL

    ' Heal player
    If Player.Health < 100 Then
        Call HealPlayer(HEAL_MEAL)
        Call SetLocalMessage("Delicious! That was filling. (+" & HEAL_MEAL & " HP)", vbGreen)
        Call AddMessage("Ate +" & HEAL_MEAL & " HP")
    Else
        Call SetLocalMessage("Delicious meal! You're stuffed!", vbGreen)
        Call AddMessage("Delicious!")
    End If

    Call PlayPurchaseSound
    Call DrawSaloonInterface
End Sub

Private Sub BuyCondom()
    If HasCondom Then
        Call SetLocalMessage("You already have a condom!", vbYellow)
        Call AddMessage("Already have one!")
        Call DrawSaloonInterface
        Exit Sub
    End If

    If Player.Cash < COST_CONDOM_SALOON Then
        Call SetLocalMessage("Not enough cash for a condom!", vbRed)
        Call AddMessage("Need $" & COST_CONDOM_SALOON)
        Call DrawSaloonInterface
        Exit Sub
    End If

    Player.Cash = Player.Cash - COST_CONDOM_SALOON
    Call GiveItem(ITEM_CONDOM)
    Call SetLocalMessage("Condom purchased! Stay safe out there!", vbGreen)
    Call AddMessage("Purchased!")
    Call PlayPurchaseSound
    Call DrawSaloonInterface
End Sub

Private Sub picSaloon_Click()
    Unload Me
End Sub

Private Sub SetLocalMessage(ByVal Msg As String, ByVal MsgColor As Long)
    CurrentMessage = Msg
    MessageColor = MsgColor
End Sub
