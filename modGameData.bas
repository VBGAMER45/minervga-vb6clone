Attribute VB_Name = "modGameData"
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

' --- Dirt Modifiers (hidden until dug) ---
Public Const MOD_NONE As Integer = 0
Public Const MOD_SILVER As Integer = 1
Public Const MOD_GOLD As Integer = 2
Public Const MOD_PLATINUM As Integer = 3
Public Const MOD_CAVEIN As Integer = 4
Public Const MOD_WATER As Integer = 5
Public Const MOD_WHIRLPOOL As Integer = 6
Public Const MOD_GRANITE As Integer = 7

' --- Mineral Values (at Bank) ---
Public Const SILVER_VALUE As Long = 16
Public Const GOLD_VALUE As Long = 60
Public Const PLATINUM_VALUE As Long = 2000

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
Public Const GRID_ROWS As Integer = 45
Public Const CELL_WIDTH As Integer = 16
Public Const CELL_HEIGHT As Integer = 24

' --- Screen/Viewport ---
Public Const VIEWPORT_COLS As Integer = 40
Public Const VIEWPORT_ROWS As Integer = 17
Public Const SCREEN_WIDTH As Integer = 768
Public Const SCREEN_HEIGHT As Integer = 480

' --- Game Settings ---
Public Const STARTING_CASH As Long = 1500
Public Const STARTING_HEALTH As Integer = 100
Public Const WIN_MONEY As Long = 20000
Public Const MIN_ELEVATOR_DEPTH As Integer = 10  ' Row 10 initially
Public Const MAX_MINE_DEPTH As Integer = 44      ' Bottom row
Public Const ELEVATOR_UPGRADE_ROWS As Integer = 3 ' 60 feet = 3 rows

' --- Damage Values ---
Public Const DAMAGE_CAVEIN As Integer = 20
Public Const DAMAGE_WATER As Integer = 4
Public Const DAMAGE_WHIRLPOOL As Integer = 20
Public Const DAMAGE_DYNAMITE As Integer = 30

' --- Digging Costs ---
Public Const DIG_COST_BASE As Long = 10
Public Const DIG_COST_SHOVEL_REDUCTION As Long = 3
Public Const DIG_COST_PICKAXE_REDUCTION As Long = 4
Public Const PUMP_COST_BASE As Long = 50
Public Const PUMP_COST_WITH_PUMP As Long = 20
Public Const DRILL_COST As Long = 25

' --- Hospital Costs ---
Public Const HEAL_COST_PER_POINT As Long = 5

' --- Saloon Costs ---
Public Const COST_BEER As Long = 5
Public Const COST_FOOD As Long = 10
Public Const COST_NIGHT As Long = 50

' --- Building Door Positions (Column) ---
Public Const DOOR_BANK As Integer = 2
Public Const DOOR_STORE As Integer = 8
Public Const DOOR_HOSPITAL As Integer = 16
Public Const DOOR_SALOON As Integer = 22

' --- Player Facing Direction ---
Public Const FACING_LEFT As Integer = 0
Public Const FACING_RIGHT As Integer = 1

' --- Game States ---
Public Const STATE_TITLE As Integer = 0
Public Const STATE_PLAYING As Integer = 1
Public Const STATE_DEAD As Integer = 2
Public Const STATE_WON As Integer = 3

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

' --- Modifier Spawn Chances (per 1000) ---
Public Const CHANCE_PLATINUM As Integer = 3      ' 0.3%
Public Const CHANCE_GOLD As Integer = 13         ' 1.0% (cumulative 1.3%)
Public Const CHANCE_SILVER As Integer = 45       ' 3.2% (cumulative 4.5%)
Public Const CHANCE_CAVEIN As Integer = 60       ' 1.5% (cumulative 6.0%)
Public Const CHANCE_WATER As Integer = 65        ' 0.5% (cumulative 6.5%)
Public Const CHANCE_WHIRLPOOL As Integer = 70    ' 0.5% (cumulative 7.0%)
' Remaining 30% chance = Granite or nothing

' --- Global Game State Variable ---
Public GameState As Integer
Public SoundEnabled As Boolean
