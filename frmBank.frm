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
Option Explicit

' ============================================================================
' MinerVGA - Bank Form (Graphical Interface matching JS version)
' ============================================================================

Private picTileset As StdPicture

Private Sub Form_Load()
    ' Load tileset for mineral icons
    On Error Resume Next
    Set picTileset = LoadPicture(App.Path & "\javascript\tileset.bmp")
    On Error GoTo 0

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

    ' Today's Quotes header and values
    picBank.ForeColor = vbWhite
    picBank.FontBold = False
    picBank.CurrentX = 20
    picBank.CurrentY = Y
    picBank.Print "Today's Quotes: "
    picBank.ForeColor = COLOR_PLATINUM
    picBank.Print "Platinum "
    picBank.ForeColor = vbWhite
    picBank.Print Format(PLATINUM_VALUE, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_GOLD
    picBank.Print "Gold "
    picBank.ForeColor = vbWhite
    picBank.Print Format(GOLD_VALUE, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_SILVER
    picBank.Print "Silver "
    picBank.ForeColor = vbWhite
    picBank.Print Format(SILVER_VALUE, "0.0")
    Y = Y + 18

    ' Your Mineral Oz
    picBank.ForeColor = vbWhite
    picBank.CurrentX = 20
    picBank.CurrentY = Y
    picBank.Print "Your Mineral Oz "
    picBank.ForeColor = COLOR_PLATINUM
    picBank.Print "Platinum "
    picBank.ForeColor = vbWhite
    picBank.Print Format(Player.Platinum, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_GOLD
    picBank.Print "Gold "
    picBank.ForeColor = vbWhite
    picBank.Print Format(Player.Gold, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_SILVER
    picBank.Print "Silver "
    picBank.ForeColor = vbWhite
    picBank.Print Format(Player.Silver, "0")
    Y = Y + 18

    ' Calculate values
    PlatTotal = Player.Platinum * PLATINUM_VALUE
    GoldTotal = Player.Gold * GOLD_VALUE
    SilverTotal = Player.Silver * SILVER_VALUE

    ' Your Minerals $
    picBank.ForeColor = vbWhite
    picBank.CurrentX = 20
    picBank.CurrentY = Y
    picBank.Print "Your Minerals $ "
    picBank.ForeColor = COLOR_PLATINUM
    picBank.Print "Platinum "
    picBank.ForeColor = vbWhite
    picBank.Print Format(PlatTotal, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_GOLD
    picBank.Print "Gold "
    picBank.ForeColor = vbWhite
    picBank.Print Format(GoldTotal, "0")
    picBank.ForeColor = vbWhite
    picBank.Print " : "
    picBank.ForeColor = COLOR_SILVER
    picBank.Print "Silver "
    picBank.ForeColor = vbWhite
    picBank.Print Format(SilverTotal, "0.0")
    Y = Y + 40

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
    ' Allow clicking to dismiss (optional)
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
