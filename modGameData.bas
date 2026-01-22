Attribute VB_Name = "modGameData"
' ============================================================================
' MinerVGA VB6 Edition by vbgamer45
' https://github.com/VBGAMER45/minervga-vb6clone
' https://www.theprogrammingzone.com/
' ============================================================================
Option Explicit

' ============================================================================
' MinerVGA - Game Data Constants Module
' ============================================================================

' --- Cell Types ---
Public Const CELL_AIR As Integer = 0
Public Const CELL_DOOR As Integer = 1
Public Const CELL_DIRT As Integer = 2
Public Const CELL_DUG As Integer = 3
Public Const CELL_ELEVATOR As Integer = 4
Public Const CELL_ELEVATOR_CAR As Integer = 5
Public Const CELL_ROAD As Integer = 6
Public Const CELL_WATER As Integer = 7
Public Const CELL_GRANITE As Integer = 8
Public Const CELL_CAVE As Integer = 9
Public Const CELL_WHIRLPOOL As Integer = 10
Public Const CELL_SPRING As Integer = 11
Public Const CELL_SANDSTONE As Integer = 12
Public Const CELL_VOLCANIC As Integer = 13
Public Const CELL_WALL As Integer = 14
Public Const CELL_ROOF As Integer = 15

' --- Dirt Modifiers (hidden until dug) ---
Public Const MOD_NONE As Integer = 0
Public Const MOD_SILVER As Integer = 1
Public Const MOD_GOLD As Integer = 2
Public Const MOD_PLATINUM As Integer = 3
Public Const MOD_CAVEIN As Integer = 4
Public Const MOD_WATER As Integer = 5
Public Const MOD_WHIRLPOOL As Integer = 6
Public Const MOD_GRANITE As Integer = 7
Public Const MOD_DIAMOND As Integer = 8
Public Const MOD_PUMP As Integer = 9
Public Const MOD_CLOVER As Integer = 10
Public Const MOD_SPRING As Integer = 11
Public Const MOD_VOLCANIC As Integer = 12
Public Const MOD_SANDSTONE As Integer = 13

' --- Base Mineral Values (at Bank) ---
Public Const BASE_SILVER_VALUE As Long = 16
Public Const BASE_GOLD_VALUE As Long = 60
Public Const BASE_PLATINUM_VALUE As Long = 2000
Public Const DIAMOND_VALUE As Long = 1000  ' Diamond price is fixed

' --- Current Mineral Prices (fluctuate 1-50% up or down) ---
Public CurrentSilverPrice As Single
Public CurrentGoldPrice As Single
Public CurrentPlatinumPrice As Single
Public LastPriceUpdate As Double  ' Timer value of last price update
Public Const PRICE_UPDATE_INTERVAL As Long = 30  ' Seconds between price changes

' --- Item Costs (at Store) ---
Public Const COST_SHOVEL As Long = 100
Public Const COST_PICKAXE As Long = 150
Public Const COST_DRILL As Long = 250
Public Const COST_LANTERN As Long = 100
Public Const COST_BUCKET As Long = 200
Public Const COST_TORCH As Long = 100
Public Const COST_DYNAMITE As Long = 300
Public Const COST_RING As Long = 100
Public Const COST_CONDOM As Long = 100
Public Const COST_ELEVATOR_UPGRADE As Long = 500

' --- Fuel Capacities ---
Public Const LANTERN_MAX_FUEL As Integer = 300
Public Const TORCH_MAX_FUEL As Integer = 100

' --- Grid Dimensions ---
Public Const GRID_COLS As Integer = 40
Public Const GRID_ROWS As Integer = 118          ' Expanded
Public Const CELL_WIDTH As Integer = 16
Public Const CELL_HEIGHT As Integer = 24

' --- Screen/Viewport ---
Public Const VIEWPORT_COLS As Integer = 32       ' Game area (512px wide)
Public Const VIEWPORT_ROWS As Integer = 17
Public Const GAME_WIDTH As Integer = 512         ' 32 tiles * 16px
Public Const GAME_HEIGHT As Integer = 408        ' 17 tiles * 24px
Public Const SIDEBAR_WIDTH As Integer = 136      ' Right sidebar
Public Const SCREEN_WIDTH As Integer = 648       ' 512 + 136
Public Const SCREEN_HEIGHT As Integer = 432      ' Game area + margin

' --- Game Settings ---
Public Const STARTING_CASH As Long = 1500
Public Const STARTING_HEALTH As Integer = 100
Public Const WIN_MONEY As Long = 20000
Public Const MIN_ELEVATOR_DEPTH As Integer = 20  ' Starting elevator depth
Public Const MAX_MINE_DEPTH As Integer = 117     ' Bottom row (expanded)
Public Const ELEVATOR_UPGRADE_ROWS As Integer = 10 ' 60 feet = 10 rows per upgrade

' --- Damage Values ---
Public Const DAMAGE_CAVEIN As Integer = 20
Public Const DAMAGE_WATER As Integer = 4
Public Const DAMAGE_WHIRLPOOL As Integer = 20
Public Const DAMAGE_DYNAMITE As Integer = 30
Public Const DAMAGE_SPRING As Integer = 15
Public Const DAMAGE_VOLCANIC As Integer = 5

' --- Digging Costs ---
Public Const DIG_COST_BASE As Long = 10
Public Const DIG_COST_SHOVEL_REDUCTION As Long = 3
Public Const DIG_COST_PICKAXE_REDUCTION As Long = 4
Public Const DIG_COST_CONDOM_REDUCTION As Long = 1  ' Condom now functional
Public Const PUMP_COST_BASE As Long = 50
Public Const PUMP_COST_WITH_PUMP As Long = 20
Public Const DRILL_COST As Long = 25
Public Const VOLCANIC_COST_MULTIPLIER As Single = 1.5  ' 50% more to dig
Public Const SANDSTONE_COST_MULTIPLIER As Single = 0.5 ' 50% less to dig

' --- Hospital Costs ---
Public Const HEAL_COST_PER_POINT As Long = 5

' --- Saloon Costs ---
Public Const COST_BEER As Long = 5
Public Const COST_FOOD As Long = 10
Public Const COST_NIGHT As Long = 50

' --- Building Door Positions (Column) ---
Public Const DOOR_OUTHOUSE As Integer = 2
Public Const DOOR_HOSPITAL As Integer = 5
Public Const DOOR_BANK As Integer = 11
Public Const DOOR_SALOON As Integer = 19
Public Const DOOR_STORE As Integer = 25

' --- Building Door Target IDs ---
Public Const BUILDING_OUTHOUSE As Integer = 0
Public Const BUILDING_BANK As Integer = 1
Public Const BUILDING_STORE As Integer = 2
Public Const BUILDING_HOSPITAL As Integer = 3
Public Const BUILDING_SALOON As Integer = 4

' --- Player Facing Direction ---
Public Const FACING_LEFT As Integer = 0
Public Const FACING_RIGHT As Integer = 1

' --- Game States ---
Public Const STATE_TITLE As Integer = 0
Public Const STATE_PLAYING As Integer = 1
Public Const STATE_DEAD As Integer = 2
Public Const STATE_WON As Integer = 3
Public Const STATE_BANKRUPT As Integer = 4

' --- Bankruptcy Threshold ---
Public Const BANKRUPTCY_LIMIT As Long = -100

' --- Key Codes ---
Public Const KEY_LEFT As Integer = 37
Public Const KEY_UP As Integer = 38
Public Const KEY_RIGHT As Integer = 39
Public Const KEY_DOWN As Integer = 40
Public Const KEY_H As Integer = 72   ' Help
Public Const KEY_X As Integer = 88   ' Exit
Public Const KEY_S As Integer = 83   ' Save
Public Const KEY_R As Integer = 82   ' Restore
Public Const KEY_E As Integer = 69   ' Enter building
Public Const KEY_T As Integer = 84   ' Elevator Top
Public Const KEY_B As Integer = 66   ' Elevator Bottom
Public Const KEY_D As Integer = 68   ' Drill
Public Const KEY_P As Integer = 80   ' Pump
Public Const KEY_Y As Integer = 89   ' dYnamite
Public Const KEY_Q As Integer = 81   ' Quiet (sound)

' --- Modifier Spawn Chances (per 1040 - cumulative thresholds) ---
Public Const CHANCE_PLATINUM As Integer = 20       ' 2/1040 - Very rare
Public Const CHANCE_GOLD As Integer = 70         ' 30/1040 - Rare
Public Const CHANCE_SILVER As Integer = 125       ' 70/1040 - Uncommon
Public Const CHANCE_SPRING As Integer = 342       ' 240/1040 - Common
Public Const CHANCE_CAVEIN As Integer = 642       ' 300/1040 - Common
Public Const CHANCE_GRANITE As Integer = 892      ' 250/1040 - Common (slightly less than cavein)
Public Const CHANCE_VOLCANIC As Integer = 1042    ' 150/1040 - Less common
' Sandstone and rare items handled in secondary roll

' --- Sprite Indices (tileset.bmp layout) ---
' Row 0 (0-7)
Public Const SPR_BLACK As Integer = 0           ' Empty/black tile
Public Const SPR_CLEARED As Integer = 1         ' Dug/cleared area
Public Const SPR_CLOVER As Integer = 2          ' Four-leaf clover item
Public Const SPR_SHOVEL As Integer = 3          ' Shovel item
Public Const SPR_PICKAXE As Integer = 4         ' Pickaxe item
Public Const SPR_DRILL As Integer = 5           ' Drill item
Public Const SPR_LAMP As Integer = 6            ' Lamp/lantern item
Public Const SPR_BUCKET As Integer = 7          ' Bucket item
' Row 1 (8-15)
Public Const SPR_TORCH As Integer = 8           ' Torch item
Public Const SPR_DYNAMITE As Integer = 9        ' Dynamite item
Public Const SPR_DIRT As Integer = 10           ' Dirt tile 1
Public Const SPR_DIRT2 As Integer = 11          ' Dirt tile 2
Public Const SPR_HIDDEN1 As Integer = 12        ' Hidden mineral stage 1
Public Const SPR_HIDDEN2 As Integer = 13        ' Hidden mineral stage 2
Public Const SPR_GRANITE As Integer = 14        ' Granite/rock
Public Const SPR_GROUNDWATER As Integer = 15    ' Groundwater barrier
' Row 2 (16-23)
Public Const SPR_SHAFT As Integer = 16          ' Elevator shaft
Public Const SPR_ELEVATOR_MIDDLE As Integer = 17 ' Elevator car
Public Const SPR_BIRD As Integer = 18           ' Bird decoration
Public Const SPR_CACTUS As Integer = 19         ' Cactus decoration
Public Const SPR_SKY As Integer = 20            ' Sky background
Public Const SPR_CLOUD_LEFT As Integer = 21     ' Cloud left
Public Const SPR_CLOUD_RIGHT As Integer = 22    ' Cloud right
Public Const SPR_WALL As Integer = 23           ' Building wall
' Row 3 (24-31)
Public Const SPR_ROOF_LEFT As Integer = 24      ' Roof left
Public Const SPR_ROOF_MIDDLE As Integer = 25    ' Roof middle
Public Const SPR_ROOF_RIGHT As Integer = 26     ' Roof right
Public Const SPR_DOOR As Integer = 27           ' Generic door
Public Const SPR_WALL_WHEEL As Integer = 28     ' Store wall (wheel)
Public Const SPR_WALL_HOSPITAL As Integer = 29  ' Hospital wall (+)
Public Const SPR_WALL_BANK As Integer = 30      ' Bank wall ($)
Public Const SPR_DOOR_SALOON As Integer = 31    ' Saloon door
' Row 4 (32-39)
Public Const SPR_WALL_DRINK As Integer = 32     ' Saloon wall (drink)
Public Const SPR_HITCHINGPOST As Integer = 33   ' Hitching post
Public Const SPR_CLOUD_MIDDLE As Integer = 34   ' Cloud middle
Public Const SPR_OUTHOUSE As Integer = 35       ' Outhouse building
Public Const SPR_WALL_BROTHEL As Integer = 36   ' Brothel symbol
Public Const SPR_BORDER As Integer = 37         ' Road/border
Public Const SPR_ELEVATOR_BOTTOM As Integer = 38 ' Elevator bottom
Public Const SPR_ELEVATOR_TOP_UNDER As Integer = 39 ' Elevator top (underground)
' Row 5 (40-47)
Public Const SPR_ELEVATOR_TOP_ABOVE As Integer = 40 ' Elevator top (above ground)
Public Const SPR_SPRING As Integer = 41         ' Water spring
Public Const SPR_WATER As Integer = 42          ' Water tile
Public Const SPR_SILVER As Integer = 43         ' Silver ore
Public Const SPR_GOLD As Integer = 44           ' Gold ore
Public Const SPR_PLATINUM As Integer = 45       ' Platinum ore
Public Const SPR_CAVEIN As Integer = 46         ' Cave-in
Public Const SPR_SANDSTONE As Integer = 47      ' Sandstone
' Row 6 (48-55)
Public Const SPR_VOLCANIC As Integer = 48       ' Volcanic rock
Public Const SPR_PUMP As Integer = 49           ' Pump item
Public Const SPR_CONDOM As Integer = 50         ' Condom item
Public Const SPR_DIAMOND As Integer = 51        ' Diamond item
Public Const SPR_RING As Integer = 52           ' Ring item
Public Const SPR_HWS As Integer = 53            ' HWS memorial
Public Const SPR_PLAYER_LEFT As Integer = 54    ' Player facing left
Public Const SPR_PLAYER_RIGHT As Integer = 55   ' Player facing right
Public Const SPR_COUNT As Integer = 56

' --- Global Game State Variable ---
Public GameState As Integer
Public SoundEnabled As Boolean

' --- Message System ---
Public Const MAX_MESSAGES As Integer = 8
Public Messages(0 To 7) As String
Public MessageCount As Integer

' ============================================================================
' Message System Procedures
' ============================================================================
Public Sub AddMessage(ByVal Msg As String)
    Dim i As Integer

    ' Shift existing messages up
    For i = MAX_MESSAGES - 1 To 1 Step -1
        Messages(i) = Messages(i - 1)
    Next i

    ' Add new message at top
    Messages(0) = Msg

    If MessageCount < MAX_MESSAGES Then
        MessageCount = MessageCount + 1
    End If
End Sub

Public Sub ClearMessages()
    Dim i As Integer
    For i = 0 To MAX_MESSAGES - 1
        Messages(i) = ""
    Next i
    MessageCount = 0
End Sub

' ============================================================================
' Mineral Price System (prices fluctuate every 30 seconds)
' ============================================================================
Public Sub InitializePrices()
    ' Set initial prices to base values
    CurrentSilverPrice = BASE_SILVER_VALUE
    CurrentGoldPrice = BASE_GOLD_VALUE
    CurrentPlatinumPrice = BASE_PLATINUM_VALUE
    LastPriceUpdate = Timer

    ' Randomize prices immediately
    Call UpdateMineralPrices
End Sub

Public Sub UpdateMineralPrices()
    ' Called when entering bank - updates prices if 30 seconds have passed
    Dim CurrentTime As Double
    Dim ElapsedTime As Double

    CurrentTime = Timer

    ' Handle midnight rollover
    If CurrentTime < LastPriceUpdate Then
        ElapsedTime = (86400 - LastPriceUpdate) + CurrentTime
    Else
        ElapsedTime = CurrentTime - LastPriceUpdate
    End If

    ' Only update if enough time has passed
    If ElapsedTime >= PRICE_UPDATE_INTERVAL Then
        ' Randomize each price: 50% to 150% of base value (1-50% up or down)
        CurrentSilverPrice = RandomizePrice(BASE_SILVER_VALUE)
        CurrentGoldPrice = RandomizePrice(BASE_GOLD_VALUE)
        CurrentPlatinumPrice = RandomizePrice(BASE_PLATINUM_VALUE)

        LastPriceUpdate = CurrentTime
    End If
End Sub

Private Function RandomizePrice(ByVal BasePrice As Long) As Single
    ' Returns a price between 50% and 150% of base (1-50% change up or down)
    Dim Multiplier As Single

    ' Generate random multiplier between 0.50 and 1.50
    Multiplier = 0.5 + (Rnd * 1#)

    RandomizePrice = BasePrice * Multiplier
End Function

Public Function GetSilverPrice() As Single
    GetSilverPrice = CurrentSilverPrice
End Function

Public Function GetGoldPrice() As Single
    GetGoldPrice = CurrentGoldPrice
End Function

Public Function GetPlatinumPrice() As Single
    GetPlatinumPrice = CurrentPlatinumPrice
End Function
