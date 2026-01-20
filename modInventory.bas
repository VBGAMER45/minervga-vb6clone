Attribute VB_Name = "modInventory"
Option Explicit

' ============================================================================
' MinerVGA - Inventory and Save/Load Module
' ============================================================================

' --- Item IDs ---
Public Const ITEM_SHOVEL As Integer = 1
Public Const ITEM_PICKAXE As Integer = 2
Public Const ITEM_DRILL As Integer = 3
Public Const ITEM_LANTERN As Integer = 4
Public Const ITEM_BUCKET As Integer = 5
Public Const ITEM_TORCH As Integer = 6
Public Const ITEM_DYNAMITE As Integer = 7
Public Const ITEM_RING As Integer = 8
Public Const ITEM_CONDOM As Integer = 9
Public Const ITEM_PUMP As Integer = 10
Public Const ITEM_CLOVER As Integer = 11
Public Const ITEM_DIAMOND As Integer = 12

' ============================================================================
' Item Purchase
' ============================================================================
Public Function BuyItem(ByVal ItemID As Integer) As Boolean
    Dim Cost As Long
    Dim ItemName As String

    ' Check if already owned
    If HasItem(ItemID) Then
        MsgBox "You already have this item!", vbExclamation, "Store"
        BuyItem = False
        Exit Function
    End If

    ' Get cost
    Cost = GetItemCost(ItemID)
    ItemName = GetItemName(ItemID)

    ' Check if can afford
    If Player.Cash < Cost Then
        MsgBox "You don't have enough money! You need $" & Cost, vbExclamation, "Store"
        BuyItem = False
        Exit Function
    End If

    ' Purchase
    Player.Cash = Player.Cash - Cost
    Call GiveItem(ItemID)

    MsgBox "You bought a " & ItemName & " for $" & Cost, vbInformation, "Store"
    BuyItem = True
End Function

' ============================================================================
' Item Queries
' ============================================================================
Public Function HasItem(ByVal ItemID As Integer) As Boolean
    Select Case ItemID
        Case ITEM_SHOVEL: HasItem = HasShovel
        Case ITEM_PICKAXE: HasItem = HasPickaxe
        Case ITEM_DRILL: HasItem = HasDrill
        Case ITEM_LANTERN: HasItem = HasLantern
        Case ITEM_BUCKET: HasItem = HasBucket
        Case ITEM_TORCH: HasItem = HasTorch
        Case ITEM_DYNAMITE: HasItem = HasDynamite
        Case ITEM_RING: HasItem = HasRing
        Case ITEM_CONDOM: HasItem = HasCondom
        Case ITEM_PUMP: HasItem = HasPump
        Case ITEM_CLOVER: HasItem = HasClover
        Case ITEM_DIAMOND: HasItem = HasDiamond
        Case Else: HasItem = False
    End Select
End Function

Public Function GetItemCost(ByVal ItemID As Integer) As Long
    Select Case ItemID
        Case ITEM_SHOVEL: GetItemCost = COST_SHOVEL
        Case ITEM_PICKAXE: GetItemCost = COST_PICKAXE
        Case ITEM_DRILL: GetItemCost = COST_DRILL
        Case ITEM_LANTERN: GetItemCost = COST_LANTERN
        Case ITEM_BUCKET: GetItemCost = COST_BUCKET
        Case ITEM_TORCH: GetItemCost = COST_TORCH
        Case ITEM_DYNAMITE: GetItemCost = COST_DYNAMITE
        Case ITEM_RING: GetItemCost = COST_RING
        Case ITEM_CONDOM: GetItemCost = COST_CONDOM
        Case ITEM_PUMP: GetItemCost = 0  ' Must be found
        Case ITEM_CLOVER: GetItemCost = 0  ' Must be found
        Case ITEM_DIAMOND: GetItemCost = 0  ' Must be found
        Case Else: GetItemCost = 0
    End Select
End Function

Public Function GetItemName(ByVal ItemID As Integer) As String
    Select Case ItemID
        Case ITEM_SHOVEL: GetItemName = "Shovel"
        Case ITEM_PICKAXE: GetItemName = "Pickaxe"
        Case ITEM_DRILL: GetItemName = "Drill"
        Case ITEM_LANTERN: GetItemName = "Lantern"
        Case ITEM_BUCKET: GetItemName = "Bucket"
        Case ITEM_TORCH: GetItemName = "Torch"
        Case ITEM_DYNAMITE: GetItemName = "Dynamite"
        Case ITEM_RING: GetItemName = "Diamond Ring"
        Case ITEM_CONDOM: GetItemName = "Condom"
        Case ITEM_PUMP: GetItemName = "Pump"
        Case ITEM_CLOVER: GetItemName = "Four-Leaf Clover"
        Case ITEM_DIAMOND: GetItemName = "Diamond"
        Case Else: GetItemName = "Unknown"
    End Select
End Function

Public Function GetItemDescription(ByVal ItemID As Integer) As String
    Select Case ItemID
        Case ITEM_SHOVEL: GetItemDescription = "Reduces digging cost"
        Case ITEM_PICKAXE: GetItemDescription = "Reduces digging cost"
        Case ITEM_DRILL: GetItemDescription = "Drills through granite"
        Case ITEM_LANTERN: GetItemDescription = "Light source (lasts longer)"
        Case ITEM_BUCKET: GetItemDescription = "Required to pump water"
        Case ITEM_TORCH: GetItemDescription = "Light source, lights dynamite"
        Case ITEM_DYNAMITE: GetItemDescription = "Blasts through obstacles"
        Case ITEM_RING: GetItemDescription = "Needed to win the game"
        Case ITEM_CONDOM: GetItemDescription = "Reduces digging cost"
        Case ITEM_PUMP: GetItemDescription = "Reduces water pumping cost"
        Case ITEM_CLOVER: GetItemDescription = "Improves your luck"
        Case ITEM_DIAMOND: GetItemDescription = "Worth $1000 at the bank"
        Case Else: GetItemDescription = ""
    End Select
End Function

' ============================================================================
' Item Management
' ============================================================================
Public Sub GiveItem(ByVal ItemID As Integer)
    Select Case ItemID
        Case ITEM_SHOVEL: HasShovel = True
        Case ITEM_PICKAXE: HasPickaxe = True
        Case ITEM_DRILL
            HasDrill = True
            DrillUses = MAX_DRILL_USES  ' 5 uses
        Case ITEM_LANTERN
            HasLantern = True
            LanternFuel = LANTERN_MAX_FUEL
        Case ITEM_BUCKET
            HasBucket = True
            BucketUses = MAX_BUCKET_USES  ' 20 uses
        Case ITEM_TORCH
            HasTorch = True
            TorchFuel = TORCH_MAX_FUEL
        Case ITEM_DYNAMITE: HasDynamite = True
        Case ITEM_RING: HasRing = True
        Case ITEM_CONDOM: HasCondom = True
        Case ITEM_PUMP: HasPump = True
        Case ITEM_CLOVER: HasClover = True
        Case ITEM_DIAMOND: HasDiamond = True
    End Select
End Sub

Public Sub RemoveItem(ByVal ItemID As Integer)
    Select Case ItemID
        Case ITEM_SHOVEL: HasShovel = False
        Case ITEM_PICKAXE: HasPickaxe = False
        Case ITEM_DRILL
            HasDrill = False
            DrillUses = 0
        Case ITEM_LANTERN
            HasLantern = False
            LanternFuel = 0
        Case ITEM_BUCKET
            HasBucket = False
            BucketUses = 0
        Case ITEM_TORCH
            HasTorch = False
            TorchFuel = 0
        Case ITEM_DYNAMITE: HasDynamite = False
        Case ITEM_RING: HasRing = False
        Case ITEM_CONDOM: HasCondom = False
        Case ITEM_PUMP: HasPump = False
        Case ITEM_CLOVER: HasClover = False
        Case ITEM_DIAMOND: HasDiamond = False
    End Select
End Sub

' ============================================================================
' Inventory Count
' ============================================================================
Public Function CountOwnedItems() As Integer
    Dim Count As Integer
    Count = 0

    If HasShovel Then Count = Count + 1
    If HasPickaxe Then Count = Count + 1
    If HasDrill Then Count = Count + 1
    If HasLantern Then Count = Count + 1
    If HasBucket Then Count = Count + 1
    If HasTorch Then Count = Count + 1
    If HasDynamite Then Count = Count + 1
    If HasRing Then Count = Count + 1
    If HasCondom Then Count = Count + 1
    If HasPump Then Count = Count + 1
    If HasClover Then Count = Count + 1
    If HasDiamond Then Count = Count + 1

    CountOwnedItems = Count
End Function

Public Function GetInventoryString() As String
    Dim Items As String
    Items = ""

    If HasShovel Then Items = Items & "Shovel, "
    If HasPickaxe Then Items = Items & "Pickaxe, "
    If HasDrill Then Items = Items & "Drill, "
    If HasLantern Then Items = Items & "Lantern (" & LanternFuel & "), "
    If HasBucket Then Items = Items & "Bucket, "
    If HasTorch Then Items = Items & "Torch (" & TorchFuel & "), "
    If HasDynamite Then Items = Items & "Dynamite, "
    If HasRing Then Items = Items & "Ring, "
    If HasCondom Then Items = Items & "Condom, "
    If HasPump Then Items = Items & "Pump, "
    If HasClover Then Items = Items & "Clover, "
    If HasDiamond Then Items = Items & "Diamond, "

    If Len(Items) > 2 Then
        Items = Left(Items, Len(Items) - 2)  ' Remove trailing ", "
    Else
        Items = "None"
    End If

    GetInventoryString = Items
End Function

' ============================================================================
' Get Random Owned Item (for item loss on whirlpool)
' ============================================================================
Public Function GetRandomOwnedItemID() As Integer
    Dim OwnedItems(1 To 12) As Integer
    Dim Count As Integer
    Dim i As Integer

    Count = 0

    If HasShovel Then Count = Count + 1: OwnedItems(Count) = ITEM_SHOVEL
    If HasPickaxe Then Count = Count + 1: OwnedItems(Count) = ITEM_PICKAXE
    If HasDrill Then Count = Count + 1: OwnedItems(Count) = ITEM_DRILL
    If HasLantern Then Count = Count + 1: OwnedItems(Count) = ITEM_LANTERN
    If HasBucket Then Count = Count + 1: OwnedItems(Count) = ITEM_BUCKET
    If HasTorch Then Count = Count + 1: OwnedItems(Count) = ITEM_TORCH
    If HasDynamite Then Count = Count + 1: OwnedItems(Count) = ITEM_DYNAMITE
    If HasCondom Then Count = Count + 1: OwnedItems(Count) = ITEM_CONDOM
    If HasPump Then Count = Count + 1: OwnedItems(Count) = ITEM_PUMP
    If HasClover Then Count = Count + 1: OwnedItems(Count) = ITEM_CLOVER
    ' Note: Ring and Diamond are not included - too valuable to lose randomly

    If Count = 0 Then
        GetRandomOwnedItemID = 0
    Else
        GetRandomOwnedItemID = OwnedItems(Int(Rnd * Count) + 1)
    End If
End Function

' ============================================================================
' Bank Operations (uses current market prices)
' ============================================================================
Public Function SellMinerals() As Long
    Dim Total As Long
    Total = 0

    ' Use current market prices
    Total = Total + CLng(Player.Silver * CurrentSilverPrice)
    Total = Total + CLng(Player.Gold * CurrentGoldPrice)
    Total = Total + CLng(Player.Platinum * CurrentPlatinumPrice)

    Player.Cash = Player.Cash + Total
    Player.Silver = 0
    Player.Gold = 0
    Player.Platinum = 0

    SellMinerals = Total
End Function

Public Function SellPlatinum() As Long
    Dim Total As Long
    Total = CLng(Player.Platinum * CurrentPlatinumPrice)
    Player.Cash = Player.Cash + Total
    Player.Platinum = 0
    SellPlatinum = Total
End Function

Public Function SellGold() As Long
    Dim Total As Long
    Total = CLng(Player.Gold * CurrentGoldPrice)
    Player.Cash = Player.Cash + Total
    Player.Gold = 0
    SellGold = Total
End Function

Public Function SellSilver() As Long
    Dim Total As Long
    Total = CLng(Player.Silver * CurrentSilverPrice)
    Player.Cash = Player.Cash + Total
    Player.Silver = 0
    SellSilver = Total
End Function

Public Function SellDiamond() As Long
    ' Sell diamond for $1000 (fixed price)
    If HasDiamond Then
        HasDiamond = False
        Player.Cash = Player.Cash + DIAMOND_VALUE
        SellDiamond = DIAMOND_VALUE
    Else
        SellDiamond = 0
    End If
End Function

Public Function GetMineralValue() As Long
    Dim Total As Long
    Total = 0

    ' Use current market prices
    Total = Total + CLng(Player.Silver * CurrentSilverPrice)
    Total = Total + CLng(Player.Gold * CurrentGoldPrice)
    Total = Total + CLng(Player.Platinum * CurrentPlatinumPrice)

    GetMineralValue = Total
End Function

' ============================================================================
' Save/Load Game (with file path parameter for Common Dialog support)
' ============================================================================
Public Sub SaveGame(Optional ByVal FilePath As String = "")
    Dim FileNum As Integer
    Dim GridPath As String

    ' Use default path if not specified
    If FilePath = "" Then
        FilePath = App.Path & "\MINERVGA.SAV"
    End If

    ' Derive grid path from save path
    GridPath = Left(FilePath, Len(FilePath) - 4) & ".GRD"

    FileNum = FreeFile

    On Error GoTo SaveError

    Open FilePath For Output As #FileNum

    ' File version for future compatibility
    Print #FileNum, "MINERVGA_SAVE_V2"

    ' Player state
    Print #FileNum, Player.X
    Print #FileNum, Player.Y
    Print #FileNum, Player.Health
    Print #FileNum, Player.Cash
    Print #FileNum, Player.Silver
    Print #FileNum, Player.Gold
    Print #FileNum, Player.Platinum
    Print #FileNum, Player.Facing

    ' Inventory (save as integers for reliable loading)
    Print #FileNum, CInt(HasShovel)
    Print #FileNum, CInt(HasPickaxe)
    Print #FileNum, CInt(HasDrill)
    Print #FileNum, CInt(HasLantern)
    Print #FileNum, CInt(HasBucket)
    Print #FileNum, CInt(HasTorch)
    Print #FileNum, CInt(HasDynamite)
    Print #FileNum, CInt(HasRing)
    Print #FileNum, CInt(HasCondom)
    Print #FileNum, CInt(HasPump)
    Print #FileNum, CInt(HasClover)
    Print #FileNum, CInt(HasDiamond)
    Print #FileNum, CInt(HasCollectedGemstone)

    ' Fuel and durability
    Print #FileNum, LanternFuel
    Print #FileNum, TorchFuel
    Print #FileNum, BucketUses
    Print #FileNum, DrillUses

    ' Elevator
    Print #FileNum, ElevatorY
    Print #FileNum, MaxElevatorDepth

    ' Luck
    Print #FileNum, PlayerLuck

    Close #FileNum

    ' Save grid separately
    Call SaveGrid(GridPath)

    Call AddMessage("Game saved!")
    Exit Sub

SaveError:
    MsgBox "Error saving game: " & Err.Description, vbCritical, "MinerVGA"
    Close #FileNum
End Sub

Public Sub LoadGame(Optional ByVal FilePath As String = "")
    Dim FileNum As Integer
    Dim GridPath As String
    Dim FileVersion As String

    ' Use default path if not specified
    If FilePath = "" Then
        FilePath = App.Path & "\MINERVGA.SAV"
    End If

    If Dir(FilePath) = "" Then
        MsgBox "Save file not found!", vbExclamation, "MinerVGA"
        Exit Sub
    End If

    ' Derive grid path from save path
    GridPath = Left(FilePath, Len(FilePath) - 4) & ".GRD"

    On Error GoTo LoadError

    FileNum = FreeFile
    Open FilePath For Input As #FileNum

    ' Check file version
    Input #FileNum, FileVersion

    If Left(FileVersion, 13) = "MINERVGA_SAVE" Then
        ' New format with version header
        Call LoadGameV2(FileNum, FileVersion)
    Else
        ' Old format - first value is Player.X
        Player.X = CInt(FileVersion)
        Call LoadGameV1(FileNum)
    End If

    Close #FileNum

    ' Load grid
    Call LoadGrid(GridPath)

    ' Make sure game is in playing state
    GameState = STATE_PLAYING

    Call AddMessage("Game loaded!")
    Exit Sub

LoadError:
    MsgBox "Error loading game: " & Err.Description, vbCritical, "MinerVGA"
    Close #FileNum
End Sub

Private Sub LoadGameV1(ByVal FileNum As Integer)
    Dim TempVar As Variant

    ' Old save format (Player.X already read)
    Input #FileNum, Player.Y
    Input #FileNum, Player.Health
    Input #FileNum, Player.Cash
    Input #FileNum, Player.Silver
    Input #FileNum, Player.Gold
    Input #FileNum, Player.Platinum
    Input #FileNum, Player.Facing

    ' Inventory (read as variant to handle both True/False strings and integers)
    Input #FileNum, TempVar: HasShovel = CBool(TempVar)
    Input #FileNum, TempVar: HasPickaxe = CBool(TempVar)
    Input #FileNum, TempVar: HasDrill = CBool(TempVar)
    Input #FileNum, TempVar: HasLantern = CBool(TempVar)
    Input #FileNum, TempVar: HasBucket = CBool(TempVar)
    Input #FileNum, TempVar: HasTorch = CBool(TempVar)
    Input #FileNum, TempVar: HasDynamite = CBool(TempVar)
    Input #FileNum, TempVar: HasRing = CBool(TempVar)
    Input #FileNum, TempVar: HasCondom = CBool(TempVar)
    Input #FileNum, TempVar: HasPump = CBool(TempVar)
    Input #FileNum, TempVar: HasClover = CBool(TempVar)
    Input #FileNum, TempVar: HasDiamond = CBool(TempVar)

    ' Fuel
    Input #FileNum, LanternFuel
    Input #FileNum, TorchFuel

    ' Elevator
    Input #FileNum, ElevatorY
    Input #FileNum, MaxElevatorDepth

    ' Set defaults for new fields
    BucketUses = MAX_BUCKET_USES
    DrillUses = MAX_DRILL_USES
    PlayerLuck = 0
End Sub

Private Sub LoadGameV2(ByVal FileNum As Integer, ByVal Version As String)
    Dim TempVar As Variant

    ' New save format (V2)
    Input #FileNum, Player.X
    Input #FileNum, Player.Y
    Input #FileNum, Player.Health
    Input #FileNum, Player.Cash
    Input #FileNum, Player.Silver
    Input #FileNum, Player.Gold
    Input #FileNum, Player.Platinum
    Input #FileNum, Player.Facing

    ' Inventory (read as variant to handle both True/False strings and integers)
    Input #FileNum, TempVar: HasShovel = CBool(TempVar)
    Input #FileNum, TempVar: HasPickaxe = CBool(TempVar)
    Input #FileNum, TempVar: HasDrill = CBool(TempVar)
    Input #FileNum, TempVar: HasLantern = CBool(TempVar)
    Input #FileNum, TempVar: HasBucket = CBool(TempVar)
    Input #FileNum, TempVar: HasTorch = CBool(TempVar)
    Input #FileNum, TempVar: HasDynamite = CBool(TempVar)
    Input #FileNum, TempVar: HasRing = CBool(TempVar)
    Input #FileNum, TempVar: HasCondom = CBool(TempVar)
    Input #FileNum, TempVar: HasPump = CBool(TempVar)
    Input #FileNum, TempVar: HasClover = CBool(TempVar)
    Input #FileNum, TempVar: HasDiamond = CBool(TempVar)

    ' HasCollectedGemstone (new field - check for EOF for old saves)
    If Not EOF(FileNum) Then
        Input #FileNum, TempVar: HasCollectedGemstone = CBool(TempVar)
    Else
        ' Old save file - default based on HasDiamond
        HasCollectedGemstone = HasDiamond
    End If

    ' Fuel and durability
    If Not EOF(FileNum) Then Input #FileNum, LanternFuel
    If Not EOF(FileNum) Then Input #FileNum, TorchFuel
    If Not EOF(FileNum) Then Input #FileNum, BucketUses
    If Not EOF(FileNum) Then Input #FileNum, DrillUses

    ' Elevator
    If Not EOF(FileNum) Then Input #FileNum, ElevatorY
    If Not EOF(FileNum) Then Input #FileNum, MaxElevatorDepth

    ' Luck
    If Not EOF(FileNum) Then Input #FileNum, PlayerLuck
End Sub
