Attribute VB_Name = "modGrid"
Option Explicit

' ============================================================================
' MinerVGA - Grid/Map Management Module
' ============================================================================

' --- Cell Data Type ---
Public Type CellData
    CellType As Integer    ' CELL_* constant
    Modifier As Integer    ' MOD_* constant (for dirt cells)
    Dug As Boolean         ' Has been mined
    DoorTarget As Integer  ' Which building (for door cells)
End Type

' --- Global Grid Array ---
Public Grid(GRID_COLS - 1, GRID_ROWS - 1) As CellData

' --- Viewport Position ---
Public ViewportX As Integer  ' Left column of viewport
Public ViewportY As Integer  ' Top row of viewport

' ============================================================================
' Helper function to check if column is a building wall
' ============================================================================
Private Function IsBuildingWall(ByVal X As Integer) As Boolean
    ' Hospital walls (3, 4, 6) - door is 5
    ' Bank walls (10, 12, 13) - door is 11
    ' Saloon walls (17, 18, 20, 21) - door is 19
    ' Store walls (24, 26, 27) - door is 25
    Select Case X
        Case 3, 4, 6, 10, 12, 13, 17, 18, 20, 21, 24, 26, 27
            IsBuildingWall = True
        Case Else
            IsBuildingWall = False
    End Select
End Function

' ============================================================================
' Grid Initialization
' ============================================================================
Public Sub InitializeGrid()
    Dim X As Integer, Y As Integer
    Dim RandVal As Integer

    Randomize Timer

    For Y = 0 To GRID_ROWS - 1
        For X = 0 To GRID_COLS - 1
            Grid(X, Y).Dug = False
            Grid(X, Y).Modifier = MOD_NONE
            Grid(X, Y).DoorTarget = 0

            ' Row 0-2: Sky (Air)
            If Y <= 2 Then
                Grid(X, Y).CellType = CELL_AIR

            ' Row 3: Town level with buildings, doors and elevator
            ElseIf Y = 3 Then
                If X = GRID_COLS - 1 Then
                    ' Elevator car starts here
                    Grid(X, Y).CellType = CELL_ELEVATOR_CAR
                ElseIf X = DOOR_OUTHOUSE Then
                    ' Outhouse door
                    Grid(X, Y).CellType = CELL_DOOR
                    Grid(X, Y).DoorTarget = BUILDING_OUTHOUSE
                ElseIf X = DOOR_HOSPITAL Then
                    ' Hospital door
                    Grid(X, Y).CellType = CELL_DOOR
                    Grid(X, Y).DoorTarget = BUILDING_HOSPITAL
                ElseIf X = DOOR_BANK Then
                    ' Bank door
                    Grid(X, Y).CellType = CELL_DOOR
                    Grid(X, Y).DoorTarget = BUILDING_BANK
                ElseIf X = DOOR_SALOON Then
                    ' Saloon door
                    Grid(X, Y).CellType = CELL_DOOR
                    Grid(X, Y).DoorTarget = BUILDING_SALOON
                ElseIf X = DOOR_STORE Then
                    ' Store door
                    Grid(X, Y).CellType = CELL_DOOR
                    Grid(X, Y).DoorTarget = BUILDING_STORE
                ElseIf IsBuildingWall(X) Then
                    ' Building walls (not doors) - still CELL_AIR but rendered differently
                    Grid(X, Y).CellType = CELL_AIR
                Else
                    ' Sky/open area
                    Grid(X, Y).CellType = CELL_AIR
                End If

            ' Row 4: Road (barrier between town and mine)
            ElseIf Y = 4 Then
                If X = GRID_COLS - 1 Then
                    Grid(X, Y).CellType = CELL_ELEVATOR
                Else
                    Grid(X, Y).CellType = CELL_ROAD
                End If

            ' Row 5+: Mine area
            Else
                ' Elevator shaft on right edge
                If X = GRID_COLS - 1 Then
                    Grid(X, Y).CellType = CELL_ELEVATOR
                Else
                    ' Generate dirt with random modifiers
                    Grid(X, Y).CellType = CELL_DIRT
                    Grid(X, Y).Modifier = GenerateModifier()
                End If
            End If
        Next X
    Next Y

    ' Initialize viewport to show town
    ViewportX = 0
    ViewportY = 0
End Sub

' ============================================================================
' Random Modifier Generation (Updated with new modifiers)
' ============================================================================
Private Function GenerateModifier() As Integer
    Dim RandVal As Integer

    ' Only some cells have modifiers (about 25% chance)
    If Rnd * 100 > 25 Then
        GenerateModifier = MOD_NONE
        Exit Function
    End If

    ' Generate random value 1-1040 (matching JS version)
    RandVal = Int(Rnd * 1040) + 1

    If RandVal <= CHANCE_PLATINUM Then
        GenerateModifier = MOD_PLATINUM
    ElseIf RandVal <= CHANCE_GOLD Then
        GenerateModifier = MOD_GOLD
    ElseIf RandVal <= CHANCE_SILVER Then
        GenerateModifier = MOD_SILVER
    ElseIf RandVal <= CHANCE_SPRING Then
        GenerateModifier = MOD_SPRING
    ElseIf RandVal <= CHANCE_CAVEIN Then
        GenerateModifier = MOD_CAVEIN
    ElseIf RandVal <= CHANCE_GRANITE Then
        GenerateModifier = MOD_GRANITE      ' Now in main spawn chain
    ElseIf RandVal <= CHANCE_VOLCANIC Then
        GenerateModifier = MOD_VOLCANIC
    Else
        ' Rare items and other materials
        Dim RareRoll As Integer
        RareRoll = Int(Rnd * 1000) + 1

        If RareRoll <= 2 Then
            GenerateModifier = MOD_DIAMOND   ' Very rare
        ElseIf RareRoll <= 5 Then
            GenerateModifier = MOD_PUMP      ' Very rare
        ElseIf RareRoll <= 8 Then
            GenerateModifier = MOD_CLOVER    ' Very rare
        ElseIf RareRoll <= 150 Then
            GenerateModifier = MOD_SANDSTONE ' Uncommon
        ElseIf RareRoll <= 200 Then
            GenerateModifier = MOD_WATER     ' Uncommon
        ElseIf RareRoll <= 230 Then
            GenerateModifier = MOD_WHIRLPOOL ' Rare
        Else
            GenerateModifier = MOD_NONE
        End If
    End If
End Function

' ============================================================================
' Viewport Management
' ============================================================================
Public Sub UpdateViewport()
    ' Center viewport on player, but clamp to grid bounds
    ViewportX = Player.X - (VIEWPORT_COLS \ 2)
    ViewportY = Player.Y - (VIEWPORT_ROWS \ 2)

    ' Clamp X
    If ViewportX < 0 Then ViewportX = 0
    If ViewportX > GRID_COLS - VIEWPORT_COLS Then
        ViewportX = GRID_COLS - VIEWPORT_COLS
    End If

    ' Clamp Y
    If ViewportY < 0 Then ViewportY = 0
    If ViewportY > GRID_ROWS - VIEWPORT_ROWS Then
        ViewportY = GRID_ROWS - VIEWPORT_ROWS
    End If
End Sub

' ============================================================================
' Grid Queries
' ============================================================================
Public Function GetCellType(ByVal X As Integer, ByVal Y As Integer) As Integer
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        GetCellType = -1
        Exit Function
    End If
    GetCellType = Grid(X, Y).CellType
End Function

Public Function GetCellModifier(ByVal X As Integer, ByVal Y As Integer) As Integer
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        GetCellModifier = -1
        Exit Function
    End If
    GetCellModifier = Grid(X, Y).Modifier
End Function

Public Function IsCellDug(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        IsCellDug = False
        Exit Function
    End If
    IsCellDug = Grid(X, Y).Dug
End Function

Public Function GetDoorTarget(ByVal X As Integer, ByVal Y As Integer) As Integer
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        GetDoorTarget = 0
        Exit Function
    End If
    GetDoorTarget = Grid(X, Y).DoorTarget
End Function

' ============================================================================
' Grid Modification
' ============================================================================
Public Sub SetCellType(ByVal X As Integer, ByVal Y As Integer, ByVal NewType As Integer)
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        Exit Sub
    End If
    Grid(X, Y).CellType = NewType
End Sub

Public Sub SetCellDug(ByVal X As Integer, ByVal Y As Integer, ByVal IsDug As Boolean)
    If X < 0 Or X >= GRID_COLS Or Y < 0 Or Y >= GRID_ROWS Then
        Exit Sub
    End If
    Grid(X, Y).Dug = IsDug
End Sub

' ============================================================================
' Depth Calculation
' ============================================================================
Public Function GetDepthInFeet(ByVal Row As Integer) As Integer
    ' Row 5 = surface (0 feet)
    ' Each row = 20 feet (cell height is 24 pixels, but we'll say 20 feet for simplicity)
    If Row <= 4 Then
        GetDepthInFeet = 0
    Else
        GetDepthInFeet = (Row - 4) * 20
    End If
End Function

' ============================================================================
' Light Detection (for finding minerals before digging)
' ============================================================================
Public Function CanSeeModifier(ByVal X As Integer, ByVal Y As Integer) As Boolean
    ' Player can see modifiers in adjacent cells if they have a light source
    Dim LightLevel As Integer

    LightLevel = 0
    If HasLantern And LanternFuel > 0 Then LightLevel = LightLevel + 2
    If HasTorch And TorchFuel > 0 Then LightLevel = LightLevel + 1
    If HasClover Then LightLevel = LightLevel + 1

    ' Check if cell is within light range
    Dim Distance As Integer
    Distance = Abs(Player.X - X) + Abs(Player.Y - Y)

    If Distance <= LightLevel Then
        CanSeeModifier = True
    Else
        CanSeeModifier = False
    End If
End Function

' ============================================================================
' Save/Load Grid State
' ============================================================================
Public Sub SaveGrid(ByVal FilePath As String)
    Dim FileNum As Integer
    Dim X As Integer, Y As Integer

    FileNum = FreeFile
    Open FilePath For Binary As #FileNum

    For Y = 0 To GRID_ROWS - 1
        For X = 0 To GRID_COLS - 1
            Put #FileNum, , Grid(X, Y).CellType
            Put #FileNum, , Grid(X, Y).Modifier
            Put #FileNum, , Grid(X, Y).Dug
        Next X
    Next Y

    Close #FileNum
End Sub

Public Sub LoadGrid(ByVal FilePath As String)
    Dim FileNum As Integer
    Dim X As Integer, Y As Integer

    If Dir(FilePath) = "" Then Exit Sub

    FileNum = FreeFile
    Open FilePath For Binary As #FileNum

    For Y = 0 To GRID_ROWS - 1
        For X = 0 To GRID_COLS - 1
            Get #FileNum, , Grid(X, Y).CellType
            Get #FileNum, , Grid(X, Y).Modifier
            Get #FileNum, , Grid(X, Y).Dug
        Next X
    Next Y

    Close #FileNum
End Sub
