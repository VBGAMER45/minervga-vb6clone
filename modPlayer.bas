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

' --- Fuel Tracking ---
Public LanternFuel As Integer
Public TorchFuel As Integer

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

    ' Reset fuel
    LanternFuel = 0
    TorchFuel = 0

    ' Elevator starts at top with limited depth
    ElevatorY = 3  ' Top position (town level)
    MaxElevatorDepth = MIN_ELEVATOR_DEPTH
End Sub

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
        Player.X = NewX
        Player.Y = NewY
        MovePlayer = True
    Else
        MovePlayer = False
    End If
End Function

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
            ' Can enter water (takes damage)
            CanEnterCell = True

        Case CELL_GRANITE
            ' Cannot enter granite without drilling
            CanEnterCell = False

        Case CELL_CAVE
            ' Can enter cave-in area
            CanEnterCell = True

        Case CELL_WHIRLPOOL
            ' Can enter whirlpool (takes damage)
            CanEnterCell = True

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

            ' Check for modifiers (minerals/hazards)
            Call HandleModifier(X, Y, Modifier)

            ' Mark as dug
            Grid(X, Y).CellType = CELL_DUG
            Grid(X, Y).Dug = True

            ' Use up light source fuel
            Call UseLightFuel

        Case CELL_WATER
            ' Take water damage
            Call InjurePlayer(DAMAGE_WATER)

        Case CELL_WHIRLPOOL
            ' Take whirlpool damage and flood nearby
            Call InjurePlayer(DAMAGE_WHIRLPOOL)
            Call FloodNearby(X, Y)

        Case CELL_CAVE
            ' Take cave-in damage
            Call InjurePlayer(DAMAGE_CAVEIN)
            Call TriggerCaveIn(X, Y)
    End Select
End Sub

' ============================================================================
' Modifier Handling (Minerals and Hazards)
' ============================================================================
Private Sub HandleModifier(ByVal X As Integer, ByVal Y As Integer, ByVal Modifier As Integer)
    Select Case Modifier
        Case MOD_SILVER
            Player.Silver = Player.Silver + 1

        Case MOD_GOLD
            Player.Gold = Player.Gold + 1

        Case MOD_PLATINUM
            Player.Platinum = Player.Platinum + 1

        Case MOD_CAVEIN
            Call InjurePlayer(DAMAGE_CAVEIN)
            Call TriggerCaveIn(X, Y)
            Grid(X, Y).CellType = CELL_CAVE

        Case MOD_WATER
            Grid(X, Y).CellType = CELL_WATER
            Call InjurePlayer(DAMAGE_WATER)

        Case MOD_WHIRLPOOL
            Grid(X, Y).CellType = CELL_WHIRLPOOL
            Call InjurePlayer(DAMAGE_WHIRLPOOL)
            Call FloodNearby(X, Y)

        Case MOD_GRANITE
            ' Granite blocks entry - this shouldn't happen
            ' as CanEnterCell should have blocked it
            Grid(X, Y).CellType = CELL_GRANITE
    End Select
End Sub

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
    ' Whirlpool floods nearby dug cells with water
    Dim X As Integer, Y As Integer

    For Y = CenterY To CenterY + 4
        For X = CenterX - 2 To CenterX + 2
            If X >= 0 And X < GRID_COLS And Y >= 0 And Y < GRID_ROWS Then
                If Grid(X, Y).CellType = CELL_DUG Then
                    If Not (X = Player.X And Y = Player.Y) Then
                        Grid(X, Y).CellType = CELL_WATER
                    End If
                End If
            End If
        Next X
    Next Y
End Sub

' ============================================================================
' Elevator Control
' ============================================================================
Public Sub ElevatorToTop()
    ' Only works if player is in elevator
    If Player.X = GRID_COLS - 1 And Player.Y = ElevatorY Then
        ElevatorY = 3  ' Town level
        Player.Y = ElevatorY
    End If
End Sub

Public Sub ElevatorToBottom()
    ' Only works if player is in elevator
    If Player.X = GRID_COLS - 1 And Player.Y = ElevatorY Then
        ElevatorY = MaxElevatorDepth
        Player.Y = ElevatorY
    End If
End Sub

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
        MsgBox "You need a drill to break through granite!", vbExclamation, "MinerVGA"
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
        DrillGranite = True
    Else
        DrillGranite = False
    End If
End Function

Public Function PumpWater(ByVal Direction As Integer) As Boolean
    Dim TargetX As Integer, TargetY As Integer
    Dim Cost As Long

    If Not HasBucket Then
        MsgBox "You need a bucket to pump water!", vbExclamation, "MinerVGA"
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

    ' Check if target is water
    If Grid(TargetX, TargetY).CellType = CELL_WATER Or Grid(TargetX, TargetY).CellType = CELL_WHIRLPOOL Then
        If HasPump Then
            Cost = PUMP_COST_WITH_PUMP
        Else
            Cost = PUMP_COST_BASE
        End If

        Player.Cash = Player.Cash - Cost
        Grid(TargetX, TargetY).CellType = CELL_DUG
        PumpWater = True
    Else
        PumpWater = False
    End If
End Function

Public Function UseDynamite() As Boolean
    If Not HasDynamite Then
        MsgBox "You need dynamite!", vbExclamation, "MinerVGA"
        UseDynamite = False
        Exit Function
    End If

    If Not HasTorch Then
        MsgBox "You need a torch to light the dynamite!", vbExclamation, "MinerVGA"
        UseDynamite = False
        Exit Function
    End If

    ' Check escape route to the right
    If Player.X >= GRID_COLS - 1 Then
        MsgBox "No escape route! You need a clear path to your right!", vbExclamation, "MinerVGA"
        UseDynamite = False
        Exit Function
    End If

    If Grid(Player.X + 1, Player.Y).CellType <> CELL_DUG And _
       Grid(Player.X + 1, Player.Y).CellType <> CELL_AIR Then
        MsgBox "No escape route! You need a clear path to your right!", vbExclamation, "MinerVGA"
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
