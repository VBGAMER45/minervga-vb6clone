VERSION 5.00
Begin VB.Form frmMain
   BackColor       =   &H00000000&
   Caption         =   "MinerVGA"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGame
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   7320
   End
   Begin VB.PictureBox picGame
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   477
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   765
      TabIndex        =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label lblStatus
      BackColor       =   &H00000000&
      Caption         =   "Health: 100%  |  Cash: $1500  |  Depth: 0 ft"
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Width           =   11535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' MinerVGA - Main Game Form
' ============================================================================

' --- Image Storage ---
Private picPlayerLeft As StdPicture
Private picPlayerRight As StdPicture
Private picDirt1 As StdPicture
Private picDirt2 As StdPicture
Private picDug As StdPicture
Private picRock As StdPicture
Private picWater As StdPicture
Private picWhirlpool As StdPicture
Private picCave As StdPicture
Private picElevator As StdPicture
Private picRoad As StdPicture
Private picSky As StdPicture
Private picDoor As StdPicture

' --- Colors for drawing (fallback if no images) ---
Private Const COLOR_SKY As Long = &HFFFF00     ' Cyan
Private Const COLOR_DIRT As Long = &H4080&      ' Brown
Private Const COLOR_DUG As Long = &H404040      ' Dark gray
Private Const COLOR_ROCK As Long = &H808080     ' Gray
Private Const COLOR_WATER As Long = &HFF0000    ' Blue
Private Const COLOR_ROAD As Long = &H606060     ' Medium gray
Private Const COLOR_ELEVATOR As Long = &H8080&  ' Dark yellow
Private Const COLOR_PLAYER As Long = &HFFFF&    ' Yellow
Private Const COLOR_SILVER As Long = &HC0C0C0   ' Silver
Private Const COLOR_GOLD As Long = &H00D7FF    ' Gold
Private Const COLOR_PLATINUM As Long = &HFFFFFF ' White/Platinum

' --- Rendering State ---
Private UseImages As Boolean
Private LastDirection As Integer

' ============================================================================
' Form Events
' ============================================================================
Private Sub Form_Load()
    ' Initialize random seed
    Randomize Timer

    ' Try to load images
    Call LoadImages

    ' Initialize game
    Call InitPlayer
    Call InitializeGrid

    ' Set game state
    GameState = STATE_TITLE
    SoundEnabled = True

    ' Show title screen
    Call ShowTitleScreen
End Sub

Private Sub Form_Resize()
    ' Resize picture box to fit form
    picGame.Width = Me.ScaleWidth
    picGame.Height = Me.ScaleHeight - 25
    lblStatus.Top = picGame.Height + 5
    lblStatus.Width = Me.ScaleWidth
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Handle input based on game state
    If GameState = STATE_TITLE Then
        ' Any key starts the game
        GameState = STATE_PLAYING
        tmrGame.Enabled = True
        Call RenderGame
        Exit Sub
    End If

    If GameState = STATE_DEAD Or GameState = STATE_WON Then
        ' Any key returns to title
        Call InitPlayer
        Call InitializeGrid
        GameState = STATE_TITLE
        Call ShowTitleScreen
        Exit Sub
    End If

    ' Game is playing - handle input
    Select Case KeyCode
        Case KEY_LEFT, KEY_RIGHT, KEY_UP, KEY_DOWN
            Call MovePlayer(KeyCode)
            LastDirection = KeyCode

        Case KEY_H
            Call ShowHelp

        Case KEY_X
            Call ExitGame

        Case KEY_S
            Call SaveGame

        Case KEY_R
            Call LoadGame

        Case KEY_E
            Call EnterBuilding

        Case KEY_T
            Call ElevatorToTop

        Case KEY_B
            Call ElevatorToBottom

        Case KEY_D
            Call DrillGranite(LastDirection)

        Case KEY_P
            Call PumpWater(LastDirection)

        Case KEY_Y
            Call UseDynamite

        Case KEY_Q
            SoundEnabled = Not SoundEnabled
    End Select

    ' Check for death
    If Player.Health <= 0 Then
        GameState = STATE_DEAD
        tmrGame.Enabled = False
        Call ShowDeathScreen
    End If
End Sub

' ============================================================================
' Game Timer
' ============================================================================
Private Sub tmrGame_Timer()
    Call RenderGame
End Sub

' ============================================================================
' Image Loading
' ============================================================================
Private Sub LoadImages()
    On Error GoTo NoImages

    Dim ImagePath As String
    ImagePath = App.Path & "\images\"

    ' Try to load images (will be BMP files)
    Set picPlayerLeft = LoadPicture(ImagePath & "player_left.bmp")
    Set picPlayerRight = LoadPicture(ImagePath & "player_right.bmp")
    Set picDirt1 = LoadPicture(ImagePath & "dirt1.bmp")
    Set picDirt2 = LoadPicture(ImagePath & "dirt2.bmp")
    Set picDug = LoadPicture(ImagePath & "dug.bmp")
    Set picRock = LoadPicture(ImagePath & "rock.bmp")
    Set picWater = LoadPicture(ImagePath & "water.bmp")
    Set picWhirlpool = LoadPicture(ImagePath & "wp.bmp")
    Set picCave = LoadPicture(ImagePath & "cave.bmp")
    Set picElevator = LoadPicture(ImagePath & "elevator.bmp")
    Set picRoad = LoadPicture(ImagePath & "divider.bmp")
    Set picSky = LoadPicture(ImagePath & "sky.bmp")
    Set picDoor = LoadPicture(ImagePath & "door.bmp")

    UseImages = True
    Exit Sub

NoImages:
    ' Fall back to colored rectangles
    UseImages = False
End Sub

' ============================================================================
' Rendering
' ============================================================================
Private Sub RenderGame()
    Dim X As Integer, Y As Integer
    Dim ScreenX As Integer, ScreenY As Integer
    Dim CellType As Integer, Modifier As Integer

    ' Update viewport
    Call UpdateViewport

    ' Clear screen
    picGame.Cls

    ' Draw grid
    For Y = ViewportY To ViewportY + VIEWPORT_ROWS
        For X = ViewportX To ViewportX + VIEWPORT_COLS
            If X >= 0 And X < GRID_COLS And Y >= 0 And Y < GRID_ROWS Then
                ScreenX = (X - ViewportX) * CELL_WIDTH
                ScreenY = (Y - ViewportY) * CELL_HEIGHT

                CellType = Grid(X, Y).CellType
                Modifier = Grid(X, Y).Modifier

                Call DrawCell(ScreenX, ScreenY, CellType, Modifier, X, Y)
            End If
        Next X
    Next Y

    ' Draw player
    Call DrawPlayer

    ' Update status bar
    Call UpdateStatus

    ' Refresh display
    picGame.Refresh
End Sub

Private Sub DrawCell(ByVal ScreenX As Integer, ByVal ScreenY As Integer, _
                     ByVal CellType As Integer, ByVal Modifier As Integer, _
                     ByVal GridX As Integer, ByVal GridY As Integer)

    Dim DrawColor As Long

    If UseImages Then
        ' Draw using images
        Select Case CellType
            Case CELL_AIR
                If Not picSky Is Nothing Then
                    picGame.PaintPicture picSky, ScreenX, ScreenY
                End If

            Case CELL_DOOR
                If Not picDoor Is Nothing Then
                    picGame.PaintPicture picDoor, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), &H8000&, BF
                End If

            Case CELL_DIRT
                ' Alternate dirt textures based on position
                If (GridX + GridY) Mod 2 = 0 Then
                    If Not picDirt1 Is Nothing Then
                        picGame.PaintPicture picDirt1, ScreenX, ScreenY
                    End If
                Else
                    If Not picDirt2 Is Nothing Then
                        picGame.PaintPicture picDirt2, ScreenX, ScreenY
                    End If
                End If

                ' Show modifier hint if player has light
                If CanSeeModifier(GridX, GridY) And Modifier <> MOD_NONE Then
                    Call DrawModifierHint(ScreenX, ScreenY, Modifier)
                End If

            Case CELL_DUG
                If Not picDug Is Nothing Then
                    picGame.PaintPicture picDug, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), COLOR_DUG, BF
                End If

            Case CELL_ELEVATOR, CELL_ELEVATOR_CAR
                If Not picElevator Is Nothing Then
                    picGame.PaintPicture picElevator, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), COLOR_ELEVATOR, BF
                End If

            Case CELL_ROAD
                If Not picRoad Is Nothing Then
                    picGame.PaintPicture picRoad, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), COLOR_ROAD, BF
                End If

            Case CELL_WATER
                If Not picWater Is Nothing Then
                    picGame.PaintPicture picWater, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), COLOR_WATER, BF
                End If

            Case CELL_GRANITE
                If Not picRock Is Nothing Then
                    picGame.PaintPicture picRock, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), COLOR_ROCK, BF
                End If

            Case CELL_CAVE
                If Not picCave Is Nothing Then
                    picGame.PaintPicture picCave, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), &H404080, BF
                End If

            Case CELL_WHIRLPOOL
                If Not picWhirlpool Is Nothing Then
                    picGame.PaintPicture picWhirlpool, ScreenX, ScreenY
                Else
                    picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), &HFF8000, BF
                End If
        End Select
    Else
        ' Draw using colored rectangles (fallback)
        Select Case CellType
            Case CELL_AIR: DrawColor = COLOR_SKY
            Case CELL_DOOR: DrawColor = &H8000&
            Case CELL_DIRT: DrawColor = COLOR_DIRT
            Case CELL_DUG: DrawColor = COLOR_DUG
            Case CELL_ELEVATOR, CELL_ELEVATOR_CAR: DrawColor = COLOR_ELEVATOR
            Case CELL_ROAD: DrawColor = COLOR_ROAD
            Case CELL_WATER: DrawColor = COLOR_WATER
            Case CELL_GRANITE: DrawColor = COLOR_ROCK
            Case CELL_CAVE: DrawColor = &H404080
            Case CELL_WHIRLPOOL: DrawColor = &HFF8000
            Case Else: DrawColor = 0
        End Select

        picGame.Line (ScreenX, ScreenY)-(ScreenX + CELL_WIDTH - 1, ScreenY + CELL_HEIGHT - 1), DrawColor, BF

        ' Show modifier hint
        If CellType = CELL_DIRT And CanSeeModifier(GridX, GridY) And Modifier <> MOD_NONE Then
            Call DrawModifierHint(ScreenX, ScreenY, Modifier)
        End If
    End If
End Sub

Private Sub DrawModifierHint(ByVal ScreenX As Integer, ByVal ScreenY As Integer, ByVal Modifier As Integer)
    ' Draw a small colored dot to indicate hidden modifier
    Dim HintColor As Long
    Dim CenterX As Integer, CenterY As Integer

    CenterX = ScreenX + CELL_WIDTH \ 2
    CenterY = ScreenY + CELL_HEIGHT \ 2

    Select Case Modifier
        Case MOD_SILVER: HintColor = COLOR_SILVER
        Case MOD_GOLD: HintColor = COLOR_GOLD
        Case MOD_PLATINUM: HintColor = COLOR_PLATINUM
        Case MOD_CAVEIN: HintColor = &H0000FF  ' Red warning
        Case MOD_WATER: HintColor = COLOR_WATER
        Case MOD_WHIRLPOOL: HintColor = &HFF8000
        Case MOD_GRANITE: HintColor = COLOR_ROCK
        Case Else: Exit Sub
    End Select

    picGame.Circle (CenterX, CenterY), 3, HintColor
End Sub

Private Sub DrawPlayer()
    Dim ScreenX As Integer, ScreenY As Integer

    ScreenX = (Player.X - ViewportX) * CELL_WIDTH
    ScreenY = (Player.Y - ViewportY) * CELL_HEIGHT

    If UseImages Then
        If Player.Facing = FACING_LEFT Then
            If Not picPlayerLeft Is Nothing Then
                picGame.PaintPicture picPlayerLeft, ScreenX, ScreenY
            End If
        Else
            If Not picPlayerRight Is Nothing Then
                picGame.PaintPicture picPlayerRight, ScreenX, ScreenY
            End If
        End If
    Else
        ' Draw player as colored rectangle
        picGame.Line (ScreenX + 2, ScreenY + 2)-(ScreenX + CELL_WIDTH - 3, ScreenY + CELL_HEIGHT - 3), COLOR_PLAYER, BF
    End If
End Sub

' ============================================================================
' Status Display
' ============================================================================
Private Sub UpdateStatus()
    Dim HealthColor As Long
    Dim StatusText As String
    Dim Depth As Integer

    Depth = GetDepthInFeet(Player.Y)

    ' Health color (red if low)
    If Player.Health < 20 Then
        lblStatus.ForeColor = vbRed
    Else
        lblStatus.ForeColor = vbGreen
    End If

    StatusText = "Health: " & Player.Health & "%  |  "
    StatusText = StatusText & "Cash: $" & Player.Cash & "  |  "
    StatusText = StatusText & "Depth: " & Depth & " ft  |  "
    StatusText = StatusText & "Ag:" & Player.Silver & " Au:" & Player.Gold & " Pt:" & Player.Platinum

    lblStatus.Caption = StatusText
End Sub

' ============================================================================
' Screen Displays
' ============================================================================
Private Sub ShowTitleScreen()
    picGame.Cls
    picGame.CurrentX = 200
    picGame.CurrentY = 100
    picGame.ForeColor = vbYellow
    picGame.FontSize = 24
    picGame.FontBold = True
    picGame.Print "MINER VGA"

    picGame.FontSize = 12
    picGame.FontBold = False
    picGame.ForeColor = vbWhite
    picGame.CurrentX = 150
    picGame.CurrentY = 180
    picGame.Print "A Visual Basic 6 Clone"

    picGame.CurrentX = 100
    picGame.CurrentY = 250
    picGame.Print "Collect $20,000 and a Diamond Ring"
    picGame.CurrentX = 100
    picGame.CurrentY = 280
    picGame.Print "to win Miss Mimi's hand in marriage!"

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 380
    picGame.Print "Press any key to start..."

    picGame.Refresh

    lblStatus.Caption = "MinerVGA - Press any key to begin"
End Sub

Private Sub ShowDeathScreen()
    picGame.Cls
    picGame.ForeColor = vbRed
    picGame.FontSize = 24
    picGame.FontBold = True
    picGame.CurrentX = 200
    picGame.CurrentY = 150
    picGame.Print "GAME OVER"

    picGame.FontSize = 14
    picGame.FontBold = False
    picGame.ForeColor = vbWhite
    picGame.CurrentX = 150
    picGame.CurrentY = 250
    picGame.Print "You have died in the mines!"

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 350
    picGame.Print "Press any key to try again..."

    picGame.Refresh
End Sub

Private Sub ShowWinScreen()
    picGame.Cls
    picGame.ForeColor = vbYellow
    picGame.FontSize = 24
    picGame.FontBold = True
    picGame.CurrentX = 150
    picGame.CurrentY = 100
    picGame.Print "CONGRATULATIONS!"

    picGame.FontSize = 14
    picGame.FontBold = False
    picGame.ForeColor = vbWhite
    picGame.CurrentX = 100
    picGame.CurrentY = 200
    picGame.Print "You have won Miss Mimi's heart!"

    picGame.CurrentX = 100
    picGame.CurrentY = 250
    picGame.Print "With $" & Player.Cash & " and a diamond ring,"
    picGame.CurrentX = 100
    picGame.CurrentY = 280
    picGame.Print "you can now retire in style!"

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 380
    picGame.Print "Press any key to play again..."

    picGame.Refresh
    GameState = STATE_WON
    tmrGame.Enabled = False
End Sub

' ============================================================================
' Building Entry
' ============================================================================
Private Sub EnterBuilding()
    Dim DoorTarget As Integer

    ' Check if player is at a door
    If Grid(Player.X, Player.Y).CellType <> CELL_DOOR Then
        Exit Sub
    End If

    DoorTarget = Grid(Player.X, Player.Y).DoorTarget

    tmrGame.Enabled = False

    Select Case DoorTarget
        Case 1  ' Bank
            frmBank.Show vbModal

        Case 2  ' Store
            frmStore.Show vbModal

        Case 3  ' Hospital
            frmHospital.Show vbModal

        Case 4  ' Saloon
            frmSaloon.Show vbModal
    End Select

    tmrGame.Enabled = True
    Call RenderGame
End Sub

' ============================================================================
' Help Screen
' ============================================================================
Private Sub ShowHelp()
    frmHelp.Show vbModal
End Sub

' ============================================================================
' Exit Game
' ============================================================================
Private Sub ExitGame()
    Dim Result As VbMsgBoxResult
    Result = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "MinerVGA")

    If Result = vbYes Then
        Unload Me
        End
    End If
End Sub
