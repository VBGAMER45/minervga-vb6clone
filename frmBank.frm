VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Community Bank"
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
   Begin VB.PictureBox picBank 
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
Attribute VB_Name = "frmBank"
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
' MinerVGA - Bank Form
' ============================================================================

Private picTileset As StdPicture

Private Sub Form_Load()
    ' Load tileset for mineral icons
    On Error Resume Next
    Set picTileset = LoadPicture(App.Path & "\tileset.bmp")
    On Error GoTo 0

    ' Update mineral prices (changes every 30 seconds)
    Call UpdateMineralPrices

    Call DrawBankInterface
End Sub

Private Sub DrawBankInterface()
    Dim Y As Integer
    Dim SilverValue As Long, GoldValue As Long, PlatValue As Long, DiamondValue As Long
    Dim SilverTotal As Long, GoldTotal As Long, PlatTotal As Long

    picBank.Cls
    picBank.FontName = "Consolas"
    picBank.FontSize = 10
    picBank.FontBold = True

    Y = 20

    ' Draw mineral icons at top left
    Call DrawMineralIcon(20, Y, SPR_PLATINUM)
    Call DrawMineralIcon(40, Y, SPR_GOLD)
    Call DrawMineralIcon(60, Y, SPR_SILVER)

    ' Draw title box
    picBank.ForeColor = vbCyan
    picBank.Line (150, Y)-(360, Y + 25), vbCyan, B
    picBank.CurrentX = 160
    picBank.CurrentY = Y + 5
    picBank.Print "The Community Bank"

    ' Draw $ symbols on sides of title
    picBank.ForeColor = vbGreen
    picBank.CurrentX = 375
    picBank.CurrentY = Y + 5
    picBank.Print "$"
    Call DrawMineralIcon(390, Y, SPR_GOLD)
    picBank.CurrentX = 415
    picBank.CurrentY = Y + 5
    picBank.Print "$"

    Y = Y + 50

    ' Welcome message
    picBank.ForeColor = vbCyan
    picBank.FontSize = 9
    picBank.FontBold = False
    picBank.CurrentX = 20
    picBank.CurrentY = Y
    picBank.Print "Welcome to The Community Bank. We are pleased to assay"
    Y = Y + 15
    picBank.CurrentX = 20
    picBank.CurrentY = Y
    picBank.Print "for you. Our rates are not only competetive..."
    Y = Y + 15
    picBank.CurrentX = 80
    picBank.CurrentY = Y
    picBank.Print "they are the only ones in town!"
    Y = Y + 30

    ' Press X to leave
    picBank.ForeColor = vbYellow
    picBank.FontBold = True
    picBank.CurrentX = 120
    picBank.CurrentY = Y
    picBank.Print "Press X to Leave the Bank"
    Y = Y + 30

    ' Calculate values using current market prices
    PlatTotal = CLng(Player.Platinum * CurrentPlatinumPrice)
    GoldTotal = CLng(Player.Gold * CurrentGoldPrice)
    SilverTotal = CLng(Player.Silver * CurrentSilverPrice)

    ' Grid layout for quotes and minerals
    Dim ColLabel As Integer, ColPlat As Integer, ColGold As Integer, ColSilver As Integer
    ColLabel = 20
    ColPlat = 180
    ColGold = 280
    ColSilver = 380

    ' Today's Quotes label
    picBank.ForeColor = vbWhite
    picBank.FontBold = True
    picBank.CurrentX = ColLabel
    picBank.CurrentY = Y
    picBank.Print "Today's Quotes:"

    ' Column headers with mineral icons
    Call DrawMineralIcon(ColPlat, Y - 2, SPR_PLATINUM)
    picBank.ForeColor = COLOR_PLATINUM
    picBank.CurrentX = ColPlat + 20
    picBank.CurrentY = Y
    picBank.Print "Platinum"

    Call DrawMineralIcon(ColGold, Y - 2, SPR_GOLD)
    picBank.ForeColor = COLOR_GOLD
    picBank.CurrentX = ColGold + 20
    picBank.CurrentY = Y
    picBank.Print "Gold"

    Call DrawMineralIcon(ColSilver, Y - 2, SPR_SILVER)
    picBank.ForeColor = COLOR_SILVER
    picBank.CurrentX = ColSilver + 20
    picBank.CurrentY = Y
    picBank.Print "Silver"
    Y = Y + 22

    ' Row 1: Price per oz (current market prices) with trend indicator
    picBank.ForeColor = vbWhite
    picBank.FontBold = False
    picBank.CurrentX = ColLabel
    picBank.CurrentY = Y
    picBank.Print "Price/oz:"

    ' Platinum price with trend
    Call DrawPriceWithTrend(ColPlat + 20, Y, CurrentPlatinumPrice, BASE_PLATINUM_VALUE, "#,##0")

    ' Gold price with trend
    Call DrawPriceWithTrend(ColGold + 20, Y, CurrentGoldPrice, BASE_GOLD_VALUE, "0")

    ' Silver price with trend
    Call DrawPriceWithTrend(ColSilver + 20, Y, CurrentSilverPrice, BASE_SILVER_VALUE, "0.0")
    Y = Y + 18

    ' Row 2: Your oz
    picBank.ForeColor = vbWhite
    picBank.CurrentX = ColLabel
    picBank.CurrentY = Y
    picBank.Print "Your oz:"
    picBank.ForeColor = COLOR_PLATINUM
    picBank.CurrentX = ColPlat + 20
    picBank.CurrentY = Y
    picBank.Print Format(Player.Platinum, "0")
    picBank.ForeColor = COLOR_GOLD
    picBank.CurrentX = ColGold + 20
    picBank.CurrentY = Y
    picBank.Print Format(Player.Gold, "0")
    picBank.ForeColor = COLOR_SILVER
    picBank.CurrentX = ColSilver + 20
    picBank.CurrentY = Y
    picBank.Print Format(Player.Silver, "0")
    Y = Y + 18

    ' Row 3: Your value $
    picBank.ForeColor = vbWhite
    picBank.CurrentX = ColLabel
    picBank.CurrentY = Y
    picBank.Print "Value $:"
    picBank.ForeColor = COLOR_PLATINUM
    picBank.CurrentX = ColPlat + 20
    picBank.CurrentY = Y
    picBank.Print "$" & Format(PlatTotal, "0")
    picBank.ForeColor = COLOR_GOLD
    picBank.CurrentX = ColGold + 20
    picBank.CurrentY = Y
    picBank.Print "$" & Format(GoldTotal, "0")
    picBank.ForeColor = COLOR_SILVER
    picBank.CurrentX = ColSilver + 20
    picBank.CurrentY = Y
    picBank.Print "$" & Format(SilverTotal, "0")
    Y = Y + 30

    ' Total value
    picBank.ForeColor = vbGreen
    picBank.FontBold = True
    picBank.CurrentX = ColLabel
    picBank.CurrentY = Y
    picBank.Print "Total Mineral Value: $" & Format(PlatTotal + GoldTotal + SilverTotal, "#,##0")
    Y = Y + 30

    ' Options
    picBank.ForeColor = vbGreen
    picBank.CurrentX = 60
    picBank.CurrentY = Y
    picBank.Print "press A to cash in all of your Minerals"
    Y = Y + 25

    picBank.CurrentX = 60
    picBank.CurrentY = Y
    picBank.Print "press P to cash in all of your Platinum"
    Y = Y + 25

    picBank.CurrentX = 60
    picBank.CurrentY = Y
    picBank.Print "press G to cash in all of your Gold"
    Y = Y + 25

    picBank.CurrentX = 60
    picBank.CurrentY = Y
    picBank.Print "press S to cash in all of your Silver"
    Y = Y + 25

    picBank.CurrentX = 60
    picBank.CurrentY = Y
    picBank.Print "press D to cash in all of your Gemstones"

    picBank.Refresh
End Sub

Private Sub DrawMineralIcon(ByVal X As Integer, ByVal Y As Integer, ByVal SpriteIdx As Integer)
    Dim SrcX As Integer, SrcY As Integer
    Dim Col As Integer, Row As Integer

    If picTileset Is Nothing Then Exit Sub

    Col = SpriteIdx Mod 8
    Row = SpriteIdx \ 8
    SrcX = Col * CELL_WIDTH
    SrcY = Row * CELL_HEIGHT

    picBank.PaintPicture picTileset, X, Y, CELL_WIDTH, CELL_HEIGHT, _
                         SrcX, SrcY, CELL_WIDTH, CELL_HEIGHT
End Sub

Private Sub DrawPriceWithTrend(ByVal X As Integer, ByVal Y As Integer, _
                               ByVal CurrentPrice As Single, ByVal BasePrice As Long, _
                               ByVal FormatStr As String)
    Dim PctChange As Single
    Dim TrendSymbol As String
    Dim TrendColor As Long

    ' Calculate percentage change from base
    PctChange = ((CurrentPrice - BasePrice) / BasePrice) * 100

    ' Determine trend symbol and color
    If PctChange >= 5 Then
        TrendSymbol = "+"
        TrendColor = vbGreen  ' Green for up
    ElseIf PctChange <= -5 Then
        TrendSymbol = "-"
        TrendColor = vbRed    ' Red for down
    Else
        TrendSymbol = "~"
        TrendColor = vbYellow ' Yellow for stable
    End If

    ' Draw price
    picBank.ForeColor = TrendColor
    picBank.CurrentX = X
    picBank.CurrentY = Y
    picBank.Print "$" & Format(CurrentPrice, FormatStr) & TrendSymbol
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Amount As Long

    Select Case KeyCode
        Case vbKeyX, vbKeyEscape
            Unload Me

        Case vbKeyA  ' Sell all minerals
            Amount = SellMinerals()
            If Amount > 0 Then
                Call AddMessage("Sold all for $" & Amount)
                Call PlayPurchaseSound
            End If
            Call DrawBankInterface

        Case vbKeyP  ' Sell platinum
            Amount = SellPlatinum()
            If Amount > 0 Then
                Call AddMessage("Sold Pt for $" & Amount)
                Call PlayPurchaseSound
            End If
            Call DrawBankInterface

        Case vbKeyG  ' Sell gold
            Amount = SellGold()
            If Amount > 0 Then
                Call AddMessage("Sold Au for $" & Amount)
                Call PlayPurchaseSound
            End If
            Call DrawBankInterface

        Case vbKeyS  ' Sell silver
            Amount = SellSilver()
            If Amount > 0 Then
                Call AddMessage("Sold Ag for $" & Amount)
                Call PlayPurchaseSound
            End If
            Call DrawBankInterface

        Case vbKeyD  ' Sell diamond/gemstones
            Amount = SellDiamond()
            If Amount > 0 Then
                Call AddMessage("Sold gem for $" & Amount)
                Call PlayPurchaseSound
            End If
            Call DrawBankInterface
    End Select
End Sub

Private Sub picBank_Click()
    Unload Me
End Sub

' Color constants for minerals
Private Property Get COLOR_PLATINUM() As Long
    COLOR_PLATINUM = &HFFFFFF
End Property

Private Property Get COLOR_GOLD() As Long
    COLOR_GOLD = &HD7FF&
End Property

Private Property Get COLOR_SILVER() As Long
    COLOR_SILVER = &HC0C0C0
End Property
