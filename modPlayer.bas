Attribute VB_Name = "modPlayer"
Option Explicit

' ============================================================================
' MinerVGA - Player State and Actions Module
' ============================================================================

' --- Player State Type ---
Public Type PlayerState
    X As Integer           ' Grid X position (column)
    Y As Integer           ' Grid Y position (row)
    Health As Integer      ' 0-100 (can exceed 100 briefly)
    Cash As Long           ' Money in bank
    Silver As Integer      ' Silver ore count (not sold yet)
    Gold As Integer        ' Gold ore count (not sold yet)
    Platinum As Integer    ' Platinum ore count (not sold yet)
    Facing As Integer      ' FACING_LEFT or FACING_RIGHT
End Type

' --- Global Player Instance ---
Public Player As PlayerState

' --- Inventory Flags ---
Public HasShovel As Boolean
Public HasPickaxe As Boolean
Public HasDrill As Boolean
Public HasLantern As Boolean
Public HasBucket As Boolean
Public HasTorch As Boolean
Public HasDynamite As Boolean
Public HasRing As Boolean
Public HasCondom As Boolean
Public HasPump As Boolean
Public HasClover As Boolean
Public HasDiamond As Boolean
Public HasCollectedGemstone As Boolean  ' True if player has ever found a diamond (required to buy ring)

' --- Fuel Tracking ---
Public LanternFuel As Integer
Public TorchFuel As Integer

' --- Item Durability ---
Public BucketUses As Integer   ' Bucket lasts 20 uses
Public DrillUses As Integer    ' Drill lasts 5 uses
Public Const MAX_BUCKET_USES As Integer = 20
Public Const MAX_DRILL_USES As Integer = 5

' --- Luck System ---
Public PlayerLuck As Integer  ' Base luck value (clover adds +20)

' --- Elevator State ---
Public ElevatorY As Integer        ' Current Y position of elevator car
Public MaxElevatorDepth As Integer ' Maximum row elevator can reach

' ============================================================================
' Player Initialization
' ============================================================================
Public Sub InitPlayer()
    ' Set starting position (town area, near center)
    Player.X = 10
    Player.Y = 3  ' Town row (above road)

    ' Starting stats
    Player.Health = STARTING_HEALTH
    Player.Cash = STARTING_CASH
    Player.Silver = 0
    Player.Gold = 0
    Player.Platinum = 0
    Player.Facing = FACING_RIGHT

    ' Clear inventory
    HasShovel = False
    HasPickaxe = False
    HasDrill = False
    HasLantern = False
    HasBucket = False
    HasTorch = False
    HasDynamite = False
    HasRing = False
    HasCondom = False
    HasPump = False
    HasClover = False
    HasDiamond = False
    HasCollectedGemstone = False

    ' Reset fuel, durability, and luck
    LanternFuel = 0
    TorchFuel = 0
    BucketUses = 0
    DrillUses = 0
    PlayerLuck = 0

    ' Elevator starts at top with limited depth
    ElevatorY = 3  ' Top position (town level)
    MaxElevatorDepth = MIN_ELEVATOR_DEPTH
End Sub

' ============================================================================
' Luck Calculation
' ============================================================================
Public Function GetPlayerLuck() As Integer
    Dim Luck As Integer
    Luck = PlayerLuck

    If HasClover Then
        Luck = Luck + 20
    End If

    GetPlayerLuck = Luck
End Function

' ============================================================================
' Player Movement
' ============================================================================
Public Function MovePlayer(ByVal Direction As Integer) As Boolean
    Dim NewX As Integer
    Dim NewY As Integer

    NewX = Player.X
    NewY = Player.Y

    Select Case Direction
        Case KEY_LEFT
            NewX = Player.X - 1
            Player.Facing = FACING_LEFT
        Case KEY_RIGHT
            NewX = Player.X + 1
            Player.Facing = FACING_RIGHT
        Case KEY_UP
            NewY = Player.Y - 1
        Case KEY_DOWN
            NewY = Player.Y + 1
    End Select

    ' Check bounds
    If NewX < 0 Or NewX >= GRID_COLS Then
        MovePlayer = False
        Exit Function
    End If
    If NewY < 0 Or NewY >= GRID_ROWS Then
        MovePlayer = False
        Exit Function
    End If

    ' Check if can enter cell
    If CanEnterCell(NewX, NewY) Then
        ' Handle entering the cell (may trigger events)
        Call EnterCell(NewX, NewY)

        ' Only move player if cell didn't become a blocking hazard
        ' (e.g., digging into water/granite/cavein/whirlpool/spring modifier)
        If CanEnterCell(NewX, NewY) Then
            Player.X = NewX
            Player.Y = NewY
            MovePlayer = True
        Else
            ' Cell became blocking after digging (hazard revealed)
            MovePlayer = False
        End If
    Else
        ' Check for hazards that deal damage on bump
        Call HandleBumpDamage(NewX, NewY)
        MovePlayer = False
    End If
End Function

' ============================================================================
' Bump Damage (when blocked by hazards)
' ============================================================================
Private Sub HandleBumpDamage(ByVal X As Integer, ByVal Y As Integer)
    Dim CellType As Integer

    ' Check bounds
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        Exit Sub
    End If

    CellType = Grid(X, Y).CellType

    Select Case CellType
        Case CELL_WATER
            Call PlayWaterSplash
            Call InjurePlayer(DAMAGE_WATER)
            Call AddMessage("Water! -" & DAMAGE_WATER & " HP")

        Case CELL_WHIRLPOOL
            Call PlaySpringSound
            Call InjurePlayer(DAMAGE_WHIRLPOOL)
            Call FloodNearby(X, Y)
            Call LoseRandomItem
            Call AddMessage("Whirlpool! -" & DAMAGE_WHIRLPOOL & " HP")

        Case CELL_CAVE
            Call PlayCaveInSound
            Call InjurePlayer(DAMAGE_CAVEIN)
            Call LoseRandomItem
            Call AddMessage("Cave-in! -" & DAMAGE_CAVEIN & " HP")

        Case CELL_GRANITE
            ' No damage, just blocked
            Call AddMessage("Need drill!")

        Case CELL_SPRING
            Call PlaySpringSound
            Call InjurePlayer(DAMAGE_SPRING)
            Call FloodNearby(X, Y)
            Call LoseRandomItem
            Call AddMessage("Spring! -" & DAMAGE_SPRING & " HP")
    End Select
End Sub

' ============================================================================
' Cell Entry Logic
' ============================================================================
Public Function CanEnterCell(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Dim CellType As Integer
    CellType = Grid(X, Y).CellType

    Select Case CellType
        Case CELL_AIR
            ' Can only enter air from same row (walking)
            If Player.Y = Y Then
                CanEnterCell = True
            Else
                CanEnterCell = False
            End If

        Case CELL_DOOR
            ' Can walk on doors
            CanEnterCell = True

        Case CELL_DIRT
            ' Can dig into dirt
            CanEnterCell = True

        Case CELL_DUG
            ' Can walk through dug areas
            CanEnterCell = True

        Case CELL_ELEVATOR, CELL_ELEVATOR_CAR
            ' Can enter elevator shaft only where car is
            If Y = ElevatorY Then
                CanEnterCell = True
            Else
                CanEnterCell = False
            End If

        Case CELL_ROAD
            ' Cannot walk on road (barrier)
            CanEnterCell = False

        Case CELL_WATER
            ' Cannot enter water - must pump with bucket using P key
            CanEnterCell = False

        Case CELL_GRANITE
            ' Cannot enter granite - must drill with D key
            CanEnterCell = False

        Case CELL_CAVE
            ' Cannot enter cave-in - must use dynamite
            CanEnterCell = False

        Case CELL_WHIRLPOOL
            ' Cannot enter whirlpool - must use dynamite
            CanEnterCell = False

        Case CELL_SPRING
            ' Cannot enter spring - must use dynamite
            CanEnterCell = False

        Case Else
            CanEnterCell = False
    End Select
End Function

Public Sub EnterCell(ByVal X As Integer, ByVal Y As Integer)
    Dim CellType As Integer
    Dim Modifier As Integer
    Dim DigCost As Long

    CellType = Grid(X, Y).CellType
    Modifier = Grid(X, Y).Modifier

    Select Case CellType
        Case CELL_DIRT
            ' Digging dirt - costs money
            DigCost = CalculateDigCost()
            Player.Cash = Player.Cash - DigCost

            ' Play appropriate digging sound based on material
            Call PlayMiningSound(Modifier)

            ' Check for modifiers (minerals/hazards)
            Call HandleModifier(X, Y, Modifier)

            ' Mark as dug ONLY if HandleModifier didn't convert to a blocking cell
            ' (hazards like water, granite, cave-in, whirlpool, spring stay as blocking cells)
            If Grid(X, Y).CellType = CELL_DIRT Then
                Grid(X, Y).CellType = CELL_DUG
                Grid(X, Y).Dug = True
            End If

            ' Use up light source fuel
            Call UseLightFuel

        Case CELL_WATER
            ' Take water damage
            Call PlayWaterSplash
            Call InjurePlayer(DAMAGE_WATER)

        Case CELL_WHIRLPOOL
            ' Take whirlpool damage and flood nearby
            Call PlaySpringSound
            Call InjurePlayer(DAMAGE_WHIRLPOOL)
            Call FloodNearby(X, Y)

        Case CELL_CAVE
            ' Take cave-in damage
            Call PlayCaveInSound
            Call InjurePlayer(DAMAGE_CAVEIN)
            Call TriggerCaveIn(X, Y)
    End Select
End Sub

' ============================================================================
' Mining Sound Based on Material
' ============================================================================
Private Sub PlayMiningSound(ByVal Modifier As Integer)
    Select Case Modifier
        Case MOD_SILVER, MOD_GOLD, MOD_PLATINUM, MOD_DIAMOND
            ' Found valuable mineral!
            Call PlayMineralSound

        Case MOD_SANDSTONE
            ' Soft sandstone
            Call PlaySandstoneSound

        Case MOD_VOLCANIC
            ' Hard volcanic rock
            Call PlayVolcanicSound

        Case MOD_GRANITE
            ' Hard granite
            Call PlayGraniteSound

        Case MOD_WATER, MOD_WHIRLPOOL, MOD_SPRING
            ' Water-related
            Call PlayWaterSplash

        Case MOD_CAVEIN
            ' Cave-in rumble
            Call PlayCaveInSound

        Case MOD_CLOVER, MOD_PUMP
            ' Found an item
            Call PlayItemSound

        Case Else
            ' Regular digging
            Call PlayDigSound
    End Select
End Sub

' ============================================================================
' Modifier Handling (Minerals and Hazards)
' ============================================================================
Private Sub HandleModifier(ByVal X As Integer, ByVal Y As Integer, ByVal Modifier As Integer)
    Select Case Modifier
        Case MOD_NONE
            ' Check for luck bonus if player has clover
            Call CheckLuckyFind(X, Y)

        Case MOD_SILVER
            Player.Silver = Player.Silver + 1
            Call AddMessage("Found Silver!")

        Case MOD_GOLD
            Player.Gold = Player.Gold + 1
            Call AddMessage("Found Gold!")

        Case MOD_PLATINUM
            Player.Platinum = Player.Platinum + 1
            Call AddMessage("Found Platinum!")

        Case MOD_DIAMOND
            ' Found a gemstone! Give diamond and mark as collected
            HasDiamond = True
            HasCollectedGemstone = True
            Call AddMessage("Found Gemstone!")
            Call PlayItemSound

        Case MOD_CLOVER
            ' Found lucky clover!
            HasClover = True
            PlayerLuck = PlayerLuck + 20
            Call AddMessage("Found Clover!")
            Call PlayItemSound

        Case MOD_PUMP
            ' Found water pump!
            HasPump = True
            Call AddMessage("Found Pump!")
            Call PlayItemSound

        Case MOD_CAVEIN
            Call InjurePlayer(DAMAGE_CAVEIN)
            Call TriggerCaveIn(X, Y)
            Call LoseRandomItem  ' Chance to lose item on cave-in
            Grid(X, Y).CellType = CELL_CAVE

        Case MOD_WATER
            Grid(X, Y).CellType = CELL_WATER
            Call InjurePlayer(DAMAGE_WATER)

        Case MOD_WHIRLPOOL
            Grid(X, Y).CellType = CELL_WHIRLPOOL
            Call InjurePlayer(DAMAGE_WHIRLPOOL)
            Call FloodNearby(X, Y)
            Call LoseRandomItem  ' Chance to lose item on whirlpool

        Case MOD_GRANITE
            ' Granite blocks entry - convert to blocking cell
            Grid(X, Y).CellType = CELL_GRANITE

        Case MOD_SPRING
            ' Spring floods area and damages player
            Grid(X, Y).CellType = CELL_SPRING
            Call InjurePlayer(DAMAGE_SPRING)
            Call FloodNearby(X, Y)
            Call LoseRandomItem  ' Chance to lose item on spring
    End Select
End Sub

' ============================================================================
' Lucky Find Check (when clover owned, chance to find bonus items)
' ============================================================================
Private Sub CheckLuckyFind(ByVal X As Integer, ByVal Y As Integer)
    Dim LuckRoll As Integer
    Dim LuckBonus As Integer

    ' Only check if player has clover
    If Not HasClover Then Exit Sub

    ' Roll for lucky find (base 5% chance, +1% per luck point)
    LuckBonus = GetPlayerLuck()
    LuckRoll = Int(Rnd * 100) + 1

    If LuckRoll <= 5 + (LuckBonus \ 5) Then
        ' Lucky find! Determine what was found
        Dim FindRoll As Integer
        FindRoll = Int(Rnd * 100) + 1

        If FindRoll <= 5 And Not HasDiamond Then
            ' Found a gemstone! (5% of lucky finds)
            HasDiamond = True
            HasCollectedGemstone = True
            Call AddMessage("Lucky Gemstone!")
            Call PlayItemSound
        ElseIf FindRoll <= 15 And Not HasPump Then
            ' Found a pump! (10% of lucky finds)
            HasPump = True
            Call AddMessage("Lucky Pump!")
            Call PlayItemSound
        ElseIf FindRoll <= 50 Then
            ' Found gold! (35% of lucky finds)
            Player.Gold = Player.Gold + 1
            Call AddMessage("Lucky Gold!")
            Call PlayMineralSound
        Else
            ' Found silver (50% of lucky finds)
            Player.Silver = Player.Silver + 1
            Call AddMessage("Lucky Silver!")
            Call PlayMineralSound
        End If
    End If
End Sub

' ============================================================================
' Random Item Loss (for hazards like whirlpool and cave-in)
' ============================================================================
Public Function LoseRandomItem() As Boolean
    ' 50% chance to lose a random item from inventory
    Dim ItemLost As Integer
    Dim ItemName As String

    If Rnd > 0.5 Then
        LoseRandomItem = False
        Exit Function
    End If

    ' Get a random owned item (uses function from modInventory)
    ItemLost = GetRandomOwnedItemID()

    If ItemLost = 0 Then
        ' No items to lose
        LoseRandomItem = False
        Exit Function
    End If

    ' Get the item name before removing
    ItemName = GetItemName(ItemLost)

    ' Remove the item
    Call RemoveItem(ItemLost)

    ' Notify player
    Call AddMessage("Lost " & ItemName & "!")

    LoseRandomItem = True
End Function

' ============================================================================
' Digging Cost Calculation
' ============================================================================
Public Function CalculateDigCost() As Long
    Dim Cost As Long
    Cost = DIG_COST_BASE

    If HasShovel Then Cost = Cost - DIG_COST_SHOVEL_REDUCTION
    If HasPickaxe Then Cost = Cost - DIG_COST_PICKAXE_REDUCTION

    If Cost < 1 Then Cost = 1

    CalculateDigCost = Cost
End Function

' ============================================================================
' Player Health
' ============================================================================
Public Sub InjurePlayer(ByVal Amount As Integer)
    Player.Health = Player.Health - Amount

    If Player.Health <= 0 Then
        Call PlayerDeath
    End If
End Sub

Public Sub HealPlayer(ByVal Amount As Integer)
    Player.Health = Player.Health + Amount
    If Player.Health > 100 Then Player.Health = 100
End Sub

Private Sub PlayerDeath()
    GameState = STATE_DEAD
    MsgBox "You have died! Game Over.", vbCritical, "MinerVGA"
End Sub

' ============================================================================
' Light Source Fuel
' ============================================================================
Private Sub UseLightFuel()
    ' Randomly use fuel when digging
    If HasLantern And LanternFuel > 0 Then
        If Rnd < 0.1 Then  ' 10% chance per dig
            LanternFuel = LanternFuel - 1
            If LanternFuel <= 0 Then
                HasLantern = False
                LanternFuel = 0
            End If
        End If
    End If

    If HasTorch And TorchFuel > 0 Then
        If Rnd < 0.2 Then  ' 20% chance per dig (burns faster)
            TorchFuel = TorchFuel - 1
            If TorchFuel <= 0 Then
                HasTorch = False
                TorchFuel = 0
            End If
        End If
    End If
End Sub

' ============================================================================
' Hazard Effects
' ============================================================================
Private Sub TriggerCaveIn(ByVal CenterX As Integer, ByVal CenterY As Integer)
    ' Cave-in fills nearby dug cells with dirt again
    Dim X As Integer, Y As Integer

    For Y = CenterY - 2 To CenterY + 2
        For X = CenterX - 2 To CenterX + 2
            If X >= 0 And X < GRID_COLS And Y >= 0 And Y < GRID_ROWS Then
                If Grid(X, Y).CellType = CELL_DUG Then
                    ' Refill with dirt (not player's cell)
                    If Not (X = Player.X And Y = Player.Y) Then
                        Grid(X, Y).CellType = CELL_DIRT
                        Grid(X, Y).Dug = False
                        Grid(X, Y).Modifier = MOD_NONE
                    End If
                End If
            End If
        Next X
    Next Y
End Sub

Private Sub FloodNearby(ByVal CenterX As Integer, ByVal CenterY As Integer)
    ' Whirlpool floods 3-15 random nearby dug cells with water
    Dim X As Integer, Y As Integer
    Dim DugCells(0 To 99, 0 To 1) As Integer  ' Store coordinates of DUG cells
    Dim DugCount As Integer
    Dim FloodCount As Integer
    Dim TargetFlood As Integer
    Dim i As Integer, j As Integer
    Dim TempX As Integer, TempY As Integer

    ' First, find all CELL_DUG within range (5x5 area around whirlpool)
    DugCount = 0
    For Y = CenterY - 2 To CenterY + 4
        For X = CenterX - 3 To CenterX + 3
            If X >= 0 And X < GRID_COLS And Y >= 0 And Y < GRID_ROWS Then
                If Grid(X, Y).CellType = CELL_DUG Then
                    If Not (X = Player.X And Y = Player.Y) Then
                        If DugCount < 100 Then
                            DugCells(DugCount, 0) = X
                            DugCells(DugCount, 1) = Y
                            DugCount = DugCount + 1
                        End If
                    End If
                End If
            End If
        Next X
    Next Y

    ' Determine how many to flood (3-15, but not more than available)
    TargetFlood = Int(Rnd * 13) + 3  ' Random 3-15
    If TargetFlood > DugCount Then TargetFlood = DugCount

    ' Shuffle the array and flood the first TargetFlood cells
    For i = DugCount - 1 To 1 Step -1
        j = Int(Rnd * (i + 1))
        ' Swap
        TempX = DugCells(i, 0)
        TempY = DugCells(i, 1)
        DugCells(i, 0) = DugCells(j, 0)
        DugCells(i, 1) = DugCells(j, 1)
        DugCells(j, 0) = TempX
        DugCells(j, 1) = TempY
    Next i

    ' Flood the selected cells
    For i = 0 To TargetFlood - 1
        Grid(DugCells(i, 0), DugCells(i, 1)).CellType = CELL_WATER
    Next i
End Sub

' ============================================================================
' Elevator Control
' ============================================================================
Public Function IsInElevator() As Boolean
    ' Check if player is in the elevator car
    IsInElevator = (Player.X = GRID_COLS - 1 And Player.Y = ElevatorY)
End Function

Public Sub ElevatorToTop()
    ' Only works if player is in elevator
    If IsInElevator() Then
        ElevatorY = 3  ' Town level
        Player.Y = ElevatorY
        Call PlayElevatorSound
    End If
End Sub

Public Sub ElevatorToBottom()
    ' Only works if player is in elevator
    If IsInElevator() Then
        ElevatorY = MaxElevatorDepth
        Player.Y = ElevatorY
        Call PlayElevatorSound
    End If
End Sub

Public Function ElevatorUp() As Boolean
    ' Move elevator up one row (if player is in elevator)
    ElevatorUp = False
    If Not IsInElevator() Then Exit Function

    If ElevatorY > 3 Then  ' Can't go above town level
        ElevatorY = ElevatorY - 1
        Player.Y = ElevatorY
        Call PlayElevatorSound
        ElevatorUp = True
    End If
End Function

Public Function ElevatorDown() As Boolean
    ' Move elevator down one row (if player is in elevator)
    ElevatorDown = False
    If Not IsInElevator() Then Exit Function

    If ElevatorY < MaxElevatorDepth Then  ' Can't go below max depth
        ElevatorY = ElevatorY + 1
        Player.Y = ElevatorY
        Call PlayElevatorSound
        ElevatorDown = True
    End If
End Function

Public Sub UpgradeElevator()
    If MaxElevatorDepth < MAX_MINE_DEPTH Then
        MaxElevatorDepth = MaxElevatorDepth + ELEVATOR_UPGRADE_ROWS
        If MaxElevatorDepth > MAX_MINE_DEPTH Then
            MaxElevatorDepth = MAX_MINE_DEPTH
        End If
    End If
End Sub

' ============================================================================
' Special Actions
' ============================================================================
Public Function DrillGranite(ByVal Direction As Integer) As Boolean
    Dim TargetX As Integer, TargetY As Integer

    If Not HasDrill Then
        Call AddMessage("Need drill!")
        DrillGranite = False
        Exit Function
    End If

    ' Determine target cell based on facing direction
    TargetX = Player.X
    TargetY = Player.Y

    Select Case Direction
        Case KEY_LEFT: TargetX = Player.X - 1
        Case KEY_RIGHT: TargetX = Player.X + 1
        Case KEY_UP: TargetY = Player.Y - 1
        Case KEY_DOWN: TargetY = Player.Y + 1
    End Select

    ' Check bounds
    If TargetX < 0 Or TargetX >= GRID_COLS Or TargetY < 0 Or TargetY >= GRID_ROWS Then
        DrillGranite = False
        Exit Function
    End If

    ' Check if target is granite
    If Grid(TargetX, TargetY).CellType = CELL_GRANITE Then
        Player.Cash = Player.Cash - DRILL_COST
        Grid(TargetX, TargetY).CellType = CELL_DUG
        Grid(TargetX, TargetY).Dug = True

        ' Use drill durability
        DrillUses = DrillUses - 1
        If DrillUses <= 0 Then
            HasDrill = False
            DrillUses = 0
            Call AddMessage("Drill broke!")
        End If

        DrillGranite = True
    Else
        DrillGranite = False
    End If
End Function

Public Function PumpWater(ByVal Direction As Integer) As Boolean
    Dim TargetX As Integer, TargetY As Integer
    Dim Cost As Long

    If Not HasBucket Then
        Call AddMessage("Need bucket!")
        PumpWater = False
        Exit Function
    End If

    ' Determine target cell
    TargetX = Player.X
    TargetY = Player.Y

    Select Case Direction
        Case KEY_LEFT: TargetX = Player.X - 1
        Case KEY_RIGHT: TargetX = Player.X + 1
        Case KEY_UP: TargetY = Player.Y - 1
        Case KEY_DOWN: TargetY = Player.Y + 1
    End Select

    ' Check bounds
    If TargetX < 0 Or TargetX >= GRID_COLS Or TargetY < 0 Or TargetY >= GRID_ROWS Then
        PumpWater = False
        Exit Function
    End If

    ' Check if target is water (only regular water can be pumped, not whirlpool/spring)
    If Grid(TargetX, TargetY).CellType = CELL_WATER Then
        If HasPump Then
            Cost = PUMP_COST_WITH_PUMP
        Else
            Cost = PUMP_COST_BASE
        End If

        Player.Cash = Player.Cash - Cost
        Grid(TargetX, TargetY).CellType = CELL_DUG

        ' Use bucket durability
        BucketUses = BucketUses - 1
        If BucketUses <= 0 Then
            HasBucket = False
            BucketUses = 0
            Call AddMessage("Bucket broke!")
        End If

        PumpWater = True
    Else
        PumpWater = False
    End If
End Function

Public Function UseDynamiteOnTarget(ByVal Direction As Integer) As Boolean
    ' Use dynamite to clear a cave-in or whirlpool in the specified direction
    Dim TargetX As Integer, TargetY As Integer
    Dim CellType As Integer

    If Not HasDynamite Then
        Call AddMessage("Need dynamite!")
        UseDynamiteOnTarget = False
        Exit Function
    End If

    If Not HasTorch Then
        Call AddMessage("Need torch!")
        UseDynamiteOnTarget = False
        Exit Function
    End If

    ' Determine target cell
    TargetX = Player.X
    TargetY = Player.Y

    Select Case Direction
        Case KEY_LEFT: TargetX = Player.X - 1
        Case KEY_RIGHT: TargetX = Player.X + 1
        Case KEY_UP: TargetY = Player.Y - 1
        Case KEY_DOWN: TargetY = Player.Y + 1
    End Select

    ' Check bounds
    If TargetX < 0 Or TargetX >= GRID_COLS Or TargetY < 0 Or TargetY >= GRID_ROWS Then
        UseDynamiteOnTarget = False
        Exit Function
    End If

    CellType = Grid(TargetX, TargetY).CellType

    ' Check if target is cave-in, whirlpool, or spring
    If CellType = CELL_CAVE Or CellType = CELL_WHIRLPOOL Or CellType = CELL_SPRING Then
        ' Clear the hazard
        Grid(TargetX, TargetY).CellType = CELL_DUG
        Grid(TargetX, TargetY).Dug = True
        Grid(TargetX, TargetY).Modifier = MOD_NONE

        ' Use up dynamite (one-time use)
        HasDynamite = False
        Call AddMessage("Boom!")

        ' Small damage from explosion
        Call InjurePlayer(5)

        UseDynamiteOnTarget = True
    Else
        Call AddMessage("No target!")
        UseDynamiteOnTarget = False
    End If
End Function

Public Function UseDynamite() As Boolean
    If Not HasDynamite Then
        Call AddMessage("Need dynamite!")
        UseDynamite = False
        Exit Function
    End If

    If Not HasTorch Then
        Call AddMessage("Need torch!")
        UseDynamite = False
        Exit Function
    End If

    ' Check escape route to the right
    If Player.X >= GRID_COLS - 1 Then
        Call AddMessage("No escape!")
        UseDynamite = False
        Exit Function
    End If

    If Grid(Player.X + 1, Player.Y).CellType <> CELL_DUG And _
       Grid(Player.X + 1, Player.Y).CellType <> CELL_AIR Then
        Call AddMessage("No escape!")
        UseDynamite = False
        Exit Function
    End If

    ' Blast 3x3 area to the left of player
    Dim X As Integer, Y As Integer
    For Y = Player.Y - 1 To Player.Y + 1
        For X = Player.X - 3 To Player.X - 1
            If X >= 0 And X < GRID_COLS And Y >= 0 And Y < GRID_ROWS Then
                If Grid(X, Y).CellType <> CELL_ELEVATOR And _
                   Grid(X, Y).CellType <> CELL_ELEVATOR_CAR And _
                   Grid(X, Y).CellType <> CELL_ROAD Then
                    Grid(X, Y).CellType = CELL_DUG
                    Grid(X, Y).Dug = True
                End If
            End If
        Next X
    Next Y

    ' Player takes some damage and moves right
    Call InjurePlayer(DAMAGE_DYNAMITE)
    Player.X = Player.X + 1

    ' Use up dynamite
    HasDynamite = False

    UseDynamite = True
End Function

' ============================================================================
' Win Condition Check
' ============================================================================
Public Function CheckWinCondition() As Boolean
    If Player.Cash >= WIN_MONEY And HasRing Then
        CheckWinCondition = True
    Else
        CheckWinCondition = False
    End If
End Function
