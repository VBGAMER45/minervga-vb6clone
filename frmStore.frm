VERSION 5.00
Begin VB.Form frmStore 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Store"
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
   Begin VB.PictureBox picStore 
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
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - General Store Form (Updated with item icons)
' ============================================================================

' --- Tileset for drawing item icons ---
Private picTileset As StdPicture

Private Sub Form_Load()
    ' Load tileset for icons
    On Error Resume Next
    Set picTileset = LoadPicture(App.Path & "\tileset.bmp")
    On Error GoTo 0

    ' Draw the store interface
    Call DrawStoreInterface
End Sub

Private Sub DrawStoreInterface()
    Dim Y As Integer
    Dim CurrentDepth As Integer, MaxDepth As Integer

    picStore.Cls
    picStore.ForeColor = vbWhite
    picStore.FontName = "Consolas"
    picStore.FontBold = True

    Y = 20

    ' Title bar
    picStore.FillStyle = 0  ' Solid
    picStore.FillColor = &H404040
    picStore.Line (100, Y - 5)-(400, Y + 20), &H404040, BF
    picStore.ForeColor = vbWhite
    picStore.FontSize = 12
    picStore.CurrentX = 170
    picStore.CurrentY = Y
    picStore.Print "General Store"
    Y = Y + 40

    ' Welcome message
    picStore.FontSize = 8
    picStore.FontBold = False
    picStore.ForeColor = vbWhite
    picStore.CurrentX = 20
    picStore.CurrentY = Y
    picStore.Print "Howdy stranger! Welcome to Emus' new and used Mining Equip-"
    Y = Y + 15
    picStore.CurrentX = 20
    picStore.CurrentY = Y
    picStore.Print "ment. We don't take returns, so... What would you like to buy?"
    Y = Y + 30

    ' Exit instruction
    picStore.ForeColor = vbYellow
    picStore.CurrentX = 150
    picStore.CurrentY = Y
    picStore.Print "Press X to Leave the store"
    Y = Y + 30

    ' Elevator info
    CurrentDepth = (MaxElevatorDepth - 4) * 6
    MaxDepth = (MAX_MINE_DEPTH - 4) * 6
    picStore.ForeColor = vbCyan
    picStore.CurrentX = 20
    picStore.CurrentY = Y
    picStore.Print "Your Elevator depth is " & CurrentDepth & " ft. Press V to buy 60 ft. for $" & COST_ELEVATOR_UPGRADE
    Y = Y + 35

    ' Draw item list with icons
    Call DrawItemList(Y)

    ' Your Cash - middle right of screen
    picStore.ForeColor = vbGreen
    picStore.FontSize = 10
    picStore.FontBold = True
    picStore.CurrentX = 320
    picStore.CurrentY = 200
    picStore.Print "Your Cash:"
    picStore.CurrentX = 320
    picStore.CurrentY = 220
    picStore.Print "$" & Format(Player.Cash, "#,##0")
    picStore.FontSize = 8
    picStore.FontBold = False

    picStore.Refresh
End Sub

Private Sub DrawItemList(ByVal StartY As Integer)
    Dim Y As Integer
    Dim ItemSprite As Integer
    Dim ItemKey As String
    Dim ItemCost As Long
    Dim ItemOwned As Boolean

    Y = StartY

    ' Item A - Shovel ($100)
    Call DrawStoreItem(Y, "A", SPR_SHOVEL, COST_SHOVEL, HasShovel)
    Y = Y + 30

    ' Item B - Pickaxe ($150)
    Call DrawStoreItem(Y, "B", SPR_PICKAXE, COST_PICKAXE, HasPickaxe)
    Y = Y + 30

    ' Item C - Drill ($250)
    Call DrawStoreItem(Y, "C", SPR_DRILL, COST_DRILL, HasDrill)
    Y = Y + 30

    ' Item D - Lantern ($150 in JS, $100 in VB6)
    Call DrawStoreItem(Y, "D", SPR_LAMP, COST_LANTERN, HasLantern)
    Y = Y + 30

    ' Item E - Bucket ($200)
    Call DrawStoreItem(Y, "E", SPR_BUCKET, COST_BUCKET, HasBucket)
    Y = Y + 30

    ' Item F - Torch ($100)
    Call DrawStoreItem(Y, "F", SPR_TORCH, COST_TORCH, HasTorch)
    Y = Y + 30

    ' Item G - Dynamite ($300)
    Call DrawStoreItem(Y, "G", SPR_DYNAMITE, COST_DYNAMITE, HasDynamite)
    Y = Y + 30

    ' Item R - Ring ($100) - Requires gemstone to purchase
    Call DrawStoreItem(Y, "R", SPR_RING, COST_RING, HasRing, True)
    Y = Y + 30
End Sub

Private Sub DrawStoreItem(ByVal Y As Integer, ByVal Key As String, ByVal SpriteIdx As Integer, ByVal Cost As Long, ByVal Owned As Boolean, Optional ByVal RequiresGem As Boolean = False)
    Dim X As Integer
    Dim SrcX As Integer, SrcY As Integer

    X = 60

    ' Draw sprite icon from tileset
    If Not picTileset Is Nothing Then
        SrcX = (SpriteIdx Mod 8) * CELL_WIDTH
        SrcY = (SpriteIdx \ 8) * CELL_HEIGHT
        picStore.PaintPicture picTileset, X, Y, CELL_WIDTH, CELL_HEIGHT, _
                              SrcX, SrcY, CELL_WIDTH, CELL_HEIGHT
    End If

    ' Draw text
    X = X + CELL_WIDTH + 15

    If Owned Then
        picStore.ForeColor = &H808080  ' Gray for owned
        picStore.CurrentX = X
        picStore.CurrentY = Y + 4
        picStore.Print "Press " & Key & " - OWNED"
    ElseIf RequiresGem And Not HasCollectedGemstone Then
        picStore.ForeColor = vbYellow  ' Yellow for locked
        picStore.CurrentX = X
        picStore.CurrentY = Y + 4
        picStore.Print "Press " & Key & " - $" & Cost & " (Requires Gemstone)"
    Else
        picStore.ForeColor = vbWhite
        picStore.CurrentX = X
        picStore.CurrentY = Y + 4
        picStore.Print "Press " & Key & " to buy this for $" & Cost
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Purchased As Boolean
    Purchased = False

    Select Case KeyCode
        Case vbKeyX, vbKeyEscape
            Unload Me
            Exit Sub

        Case vbKeyA  ' Shovel
            If TryBuyItem(ITEM_SHOVEL, COST_SHOVEL) Then Purchased = True

        Case vbKeyB  ' Pickaxe
            If TryBuyItem(ITEM_PICKAXE, COST_PICKAXE) Then Purchased = True

        Case vbKeyC  ' Drill
            If TryBuyItem(ITEM_DRILL, COST_DRILL) Then Purchased = True

        Case vbKeyD  ' Lantern
            If TryBuyItem(ITEM_LANTERN, COST_LANTERN) Then Purchased = True

        Case vbKeyE  ' Bucket
            If TryBuyItem(ITEM_BUCKET, COST_BUCKET) Then Purchased = True

        Case vbKeyF  ' Torch
            If TryBuyItem(ITEM_TORCH, COST_TORCH) Then Purchased = True

        Case vbKeyG  ' Dynamite
            If TryBuyItem(ITEM_DYNAMITE, COST_DYNAMITE) Then Purchased = True

        Case vbKeyR  ' Ring
            If TryBuyItem(ITEM_RING, COST_RING) Then Purchased = True

        Case vbKeyV  ' Elevator upgrade
            Call UpgradeElevatorFromStore

    End Select

    If Purchased Then
        Call PlayPurchaseSound
        Call DrawStoreInterface
    End If
End Sub

Private Function TryBuyItem(ByVal ItemID As Integer, ByVal Cost As Long) As Boolean
    TryBuyItem = False

    ' Check if already owned
    If HasItem(ItemID) Then
        Call AddMessage("Already owned!")
        Exit Function
    End If

    ' Special check for ring - requires having collected a gemstone first
    If ItemID = ITEM_RING Then
        If Not HasCollectedGemstone Then
            Call AddMessage("Need gemstone first!")
            MsgBox "The jeweler says: 'I need a gemstone to craft this ring!" & vbCrLf & _
                   "Find a gemstone while mining and bring it to me.'", vbInformation, "General Store"
            Exit Function
        End If
    End If

    ' Check if can afford
    If Player.Cash < Cost Then
        Call AddMessage("Not enough $!")
        Exit Function
    End If

    ' Purchase
    Player.Cash = Player.Cash - Cost
    Call GiveItem(ItemID)
    Call AddMessage("Purchased!")
    TryBuyItem = True
End Function

Private Sub UpgradeElevatorFromStore()
    If Player.Cash < COST_ELEVATOR_UPGRADE Then
        Call AddMessage("Need $" & COST_ELEVATOR_UPGRADE)
        Exit Sub
    End If

    If MaxElevatorDepth >= MAX_MINE_DEPTH Then
        Call AddMessage("Max depth!")
        Exit Sub
    End If

    Player.Cash = Player.Cash - COST_ELEVATOR_UPGRADE
    Call UpgradeElevator
    Call PlayPurchaseSound
    Call AddMessage("Elevator upgraded!")
    Call DrawStoreInterface
End Sub
