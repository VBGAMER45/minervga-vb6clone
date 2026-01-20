VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "MinerVGA"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9720
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   600
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sav"
      DialogTitle     =   "MinerVGA Save Game"
      Filter          =   "Save Files (*.sav)|*.sav|All Files (*.*)|*.*"
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   6000
   End
   Begin VB.PictureBox picTilesetHolder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox picGame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6120
      Left            =   0
      ScaleHeight     =   404
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   0
      Top             =   0
      Width           =   9720
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   "Press H for Help"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6180
      Width           =   9735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadGame 
         Caption         =   "Load Game"
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "Save Game"
      End
      Begin VB.Menu mnuSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "High Scores"
      End
      Begin VB.Menu mnuSEP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
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
' MinerVGA - Main Game Form
' ============================================================================

' --- Windows API for Transparent Drawing ---
Private Declare Function TransparentBlt Lib "msimg32.dll" ( _
    ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
    ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
    ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
    ByVal crTransparent As Long) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

' --- Common Dialog Flags ---
Private Const cdlOFNFileMustExist = &H1000
Private Const cdlOFNOverwritePrompt = &H2
Private Const cdlOFNHideReadOnly = &H4

' --- Transparent color (white) ---
Private Const TRANSPARENT_COLOR As Long = &HFFFFFF

' --- Sprite Array Storage ---
Private Sprites(0 To SPR_COUNT - 1) As StdPicture

' --- Colors for drawing (fallback if no images) ---
' NOTE: Use & suffix on hex values to ensure Long type (prevents sign issues)
Private Const COLOR_SKY As Long = &HFFFF00      ' Cyan
Private Const COLOR_DIRT As Long = &H4080&      ' Brown
Private Const COLOR_DUG As Long = &H404040      ' Dark gray
Private Const COLOR_ROCK As Long = &H808080     ' Gray
Private Const COLOR_WATER As Long = &HFF0000    ' Blue
Private Const COLOR_ROAD As Long = &H606060     ' Medium gray
Private Const COLOR_ELEVATOR As Long = &H8080&  ' Dark yellow
Private Const COLOR_PLAYER As Long = &HFFFF&    ' Yellow
Private Const COLOR_SILVER As Long = &HC0C0C0   ' Silver
Private Const COLOR_GOLD As Long = &HD7FF&      ' Gold (BGR: 0, 215, 255)
Private Const COLOR_PLATINUM As Long = &HFFFFFF  ' White/Platinum
Private Const COLOR_SPRING As Long = &HFFFF80   ' Light cyan
Private Const COLOR_VOLCANIC As Long = &H404080  ' Dark red-brown
Private Const COLOR_SANDSTONE As Long = &H80C0FF  ' Sandy color

' --- Rendering State ---
Private UseImages As Boolean
Private LastDirection As Integer
Private SpritesLoaded As Boolean

' --- Sprite filename mapping ---
Private SpriteFiles(0 To SPR_COUNT - 1) As String
' --- Tileset Storage ---
Private picTileset As StdPicture
Private picTitleScreen As StdPicture
Private TilesetCols As Integer
Private TilesetRows As Integer


' ============================================================================
' Form Events
' ============================================================================
Private Sub Form_Load()
    ' Initialize random seed
    Randomize Timer

    ' Initialize sprite filenames
    Call InitSpriteFilenames

    ' Try to load sprites
    Call LoadSprites

    ' Initialize game
    Call InitPlayer
    Call InitializeGrid
    Call ClearMessages
    Call InitializePrices  ' Initialize mineral market prices
    Call InitHighScores    ' Initialize high score system

    ' Set game state
    GameState = STATE_TITLE
    SoundEnabled = True

    ' Get twips conversion
    Dim TwipsX As Single, TwipsY As Single
    TwipsX = Screen.TwipsPerPixelX
    TwipsY = Screen.TwipsPerPixelY
    If TwipsX <= 0 Then TwipsX = 15
    If TwipsY <= 0 Then TwipsY = 15

    ' Initialize game area (includes sidebar area)
    picGame.ScaleMode = vbPixels
    picGame.AutoRedraw = True
    picGame.BackColor = vbBlack
    picGame.Visible = True
    picGame.Move 0, 0, SCREEN_WIDTH * TwipsX, GAME_HEIGHT * TwipsY

    ' Show title screen
    Call ShowTitleScreen
End Sub

Private Sub Form_Resize()
    ' Fixed layout
    Dim TwipsX As Single, TwipsY As Single

    ' Don't resize if form is minimized
    If Me.WindowState = vbMinimized Then Exit Sub

    TwipsX = Screen.TwipsPerPixelX
    TwipsY = Screen.TwipsPerPixelY

    ' Guard against invalid values
    If TwipsX <= 0 Then TwipsX = 15
    If TwipsY <= 0 Then TwipsY = 15

    ' Position game area (includes sidebar)
    picGame.Move 0, 0, SCREEN_WIDTH * TwipsX, GAME_HEIGHT * TwipsY

    ' Position status bar below
    lblStatus.Move 0, GAME_HEIGHT * TwipsY + (5 * TwipsY), SCREEN_WIDTH * TwipsX
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Handle input based on game state
    If GameState = STATE_TITLE Then
        ' Any key starts the game
        GameState = STATE_PLAYING
        tmrGame.Enabled = True
        Call AddMessage("Digging...")
        Call RenderGame
        Exit Sub
    End If

    If GameState = STATE_DEAD Or GameState = STATE_WON Or GameState = STATE_BANKRUPT Then
        ' Show high scores before returning to title
        Dim FinalScore As Long
        FinalScore = Player.Cash
        If FinalScore < 0 Then FinalScore = 0

        ' Reset game first
        Call InitPlayer
        Call InitializeGrid
        Call ClearMessages
        GameState = STATE_TITLE

        ' Now show high scores (modal - blocks until closed)
        Call ShowHighScores(FinalScore)

        ' Then show title
        Call ShowTitleScreen
        Exit Sub
    End If

    ' Game is playing - handle input
    Select Case KeyCode
        Case KEY_UP
            ' If in elevator, move elevator up; otherwise normal movement
            If IsInElevator() Then
                Call ElevatorUp
            Else
                Call MovePlayer(KeyCode)
            End If
            LastDirection = KeyCode

        Case KEY_DOWN
            ' If in elevator, move elevator down; otherwise normal movement
            If IsInElevator() Then
                Call ElevatorDown
            Else
                Call MovePlayer(KeyCode)
            End If
            LastDirection = KeyCode

        Case KEY_LEFT, KEY_RIGHT
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
            ' Try to clear cave-in/whirlpool first, otherwise do big blast
            If Not UseDynamiteOnTarget(LastDirection) Then
                Call UseDynamite
            End If

        Case KEY_Q
            SoundEnabled = Not SoundEnabled
            If SoundEnabled Then
                Call AddMessage("Sound ON")
            Else
                Call AddMessage("Sound OFF")
            End If
    End Select

    ' Check for death
    If Player.Health <= 0 Then
        GameState = STATE_DEAD
        tmrGame.Enabled = False
        Call ShowDeathScreen
        Exit Sub
    End If

    ' Check for bankruptcy
    If Player.Cash < BANKRUPTCY_LIMIT Then
        GameState = STATE_BANKRUPT
        tmrGame.Enabled = False
        Call ShowBankruptScreen
        Exit Sub
    End If

    ' Check for win condition
    Call CheckWinCondition
End Sub

Private Sub mnuExit_Click()
    Call ExitGame
End Sub

Private Sub mnuHelp_Click()
    Call ShowHelp
End Sub

Private Sub mnuHighScores_Click()
    Call ShowHighScores(0)
End Sub

Private Sub mnuLoadGame_Click()
    On Error GoTo LoadCancelled

    ' Configure dialog for loading
    dlgFile.DialogTitle = "Load MinerVGA Game"
    dlgFile.Filter = "Save Files (*.sav)|*.sav|All Files (*.*)|*.*"
    dlgFile.FilterIndex = 1
    dlgFile.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgFile.InitDir = App.Path
    dlgFile.FileName = ""

    ' Show Open dialog (will raise error if cancelled)
    dlgFile.ShowOpen

    ' User selected a file - load it
    If dlgFile.FileName <> "" Then
        Call LoadGame(dlgFile.FileName)
        GameState = STATE_PLAYING
        tmrGame.Enabled = True
        Call RenderGame
    End If
    Exit Sub

LoadCancelled:
    ' User cancelled the dialog - do nothing
End Sub

Private Sub mnuSaveGame_Click()
    On Error GoTo SaveCancelled

    ' Configure dialog for saving
    dlgFile.DialogTitle = "Save MinerVGA Game"
    dlgFile.Filter = "Save Files (*.sav)|*.sav|All Files (*.*)|*.*"
    dlgFile.FilterIndex = 1
    dlgFile.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
    dlgFile.InitDir = App.Path
    dlgFile.FileName = "MinerVGA_Save.sav"

    ' Show Save dialog (will raise error if cancelled)
    dlgFile.ShowSave

    ' User selected a file - save to it
    If dlgFile.FileName <> "" Then
        Call SaveGame(dlgFile.FileName)
    End If
    Exit Sub

SaveCancelled:
    ' User cancelled the dialog - do nothing
End Sub

' ============================================================================
' Game Timer
' ============================================================================
Private Sub tmrGame_Timer()
    Call RenderGame
End Sub

' ============================================================================
' Sprite Loading
' ============================================================================

Private Sub InitSpriteFilenames()
    ' Initialize tileset dimensions (8 columns x 7 rows)
    TilesetCols = 8
    TilesetRows = 7
End Sub

Private Sub LoadSprites()
    On Error GoTo NoSprites

    Dim TilesetPath As String
    Dim TitlePath As String

    TilesetPath = App.Path & "\tileset.bmp"
    TitlePath = App.Path & "\title-screen.bmp"

    ' Load the tileset into the hidden PictureBox for TransparentBlt
    picTilesetHolder.Picture = LoadPicture(TilesetPath)

    ' Also keep StdPicture for compatibility
    Set picTileset = LoadPicture(TilesetPath)

    ' Load title screen
    On Error Resume Next
    Set picTitleScreen = LoadPicture(TitlePath)
    On Error GoTo 0

    SpritesLoaded = True
    UseImages = True
    Exit Sub

NoSprites:
    ' Fall back to colored rectangles
    SpritesLoaded = False
    UseImages = False
End Sub

Private Sub DrawSprite(ByVal SpriteIndex As Integer, ByVal DestX As Integer, ByVal DestY As Integer)
    ' Draw a sprite from the tileset to the game picture box (normal, no transparency)
    Dim SrcX As Integer, SrcY As Integer
    Dim Col As Integer, Row As Integer

    If Not SpritesLoaded Then Exit Sub
    If SpriteIndex < 0 Or SpriteIndex >= SPR_COUNT Then Exit Sub

    ' Calculate source position in tileset
    Col = SpriteIndex Mod TilesetCols
    Row = SpriteIndex \ TilesetCols

    SrcX = Col * CELL_WIDTH
    SrcY = Row * CELL_HEIGHT

    ' Use PaintPicture for normal sprites (no transparency)
    picGame.PaintPicture picTileset, DestX, DestY, CELL_WIDTH, CELL_HEIGHT, _
                         SrcX, SrcY, CELL_WIDTH, CELL_HEIGHT
End Sub

Private Sub DrawSpriteTransparent(ByVal SpriteIndex As Integer, ByVal DestX As Integer, ByVal DestY As Integer)
    ' Draw a sprite with transparency (white = transparent) - used for player only
    Dim SrcX As Long, SrcY As Long
    Dim Col As Integer, Row As Integer
    Dim Result As Long

    If Not SpritesLoaded Then Exit Sub
    If SpriteIndex < 0 Or SpriteIndex >= SPR_COUNT Then Exit Sub

    ' Calculate source position in tileset
    Col = SpriteIndex Mod TilesetCols
    Row = SpriteIndex \ TilesetCols

    SrcX = Col * CELL_WIDTH
    SrcY = Row * CELL_HEIGHT

    ' Use TransparentBlt for transparency (white = transparent)
    Result = TransparentBlt(picGame.hDC, DestX, DestY, CELL_WIDTH, CELL_HEIGHT, _
                            picTilesetHolder.hDC, SrcX, SrcY, CELL_WIDTH, CELL_HEIGHT, _
                            TRANSPARENT_COLOR)
End Sub

' ============================================================================
' Town Sprite Mapping
' ============================================================================
Private Function GetTownSprite(ByVal X As Integer, ByVal Y As Integer) As Integer
    ' Returns sprite index for town area, or -1 if default sky
    GetTownSprite = -1

    ' Row 0: Sky with bird and clouds
    If Y = 0 Then
        Select Case X
            Case 2: GetTownSprite = SPR_BIRD
            Case 9: GetTownSprite = SPR_CLOUD_LEFT
            Case 10: GetTownSprite = SPR_CLOUD_MIDDLE
            Case 11: GetTownSprite = SPR_CLOUD_RIGHT
        End Select
        Exit Function
    End If

    ' Row 1: Sky with clouds and bird
    If Y = 1 Then
        Select Case X
            Case 18: GetTownSprite = SPR_CLOUD_LEFT
            Case 19: GetTownSprite = SPR_CLOUD_RIGHT
            Case 23: GetTownSprite = SPR_BIRD
        End Select
        Exit Function
    End If

    ' Row 2: Building roofs
    If Y = 2 Then
        Select Case X
            ' Hospital roof (cols 3-6)
            Case 3: GetTownSprite = SPR_ROOF_LEFT
            Case 4, 5: GetTownSprite = SPR_ROOF_MIDDLE
            Case 6: GetTownSprite = SPR_ROOF_RIGHT
            ' Bank roof (cols 10-13)
            Case 10: GetTownSprite = SPR_ROOF_LEFT
            Case 11, 12: GetTownSprite = SPR_ROOF_MIDDLE
            Case 13: GetTownSprite = SPR_ROOF_RIGHT
            ' Saloon roof (cols 17-21)
            Case 17: GetTownSprite = SPR_ROOF_LEFT
            Case 18, 19, 20: GetTownSprite = SPR_ROOF_MIDDLE
            Case 21: GetTownSprite = SPR_ROOF_RIGHT
            ' Store roof (cols 24-27)
            Case 24: GetTownSprite = SPR_ROOF_LEFT
            Case 25, 26: GetTownSprite = SPR_ROOF_MIDDLE
            Case 27: GetTownSprite = SPR_ROOF_RIGHT
        End Select
        Exit Function
    End If

    ' Row 3: Building walls, doors, and decorations
    If Y = 3 Then
        Select Case X
            Case 0: GetTownSprite = SPR_CACTUS
            Case 2: GetTownSprite = SPR_OUTHOUSE
            ' Hospital (cols 3-6)
            Case 3: GetTownSprite = SPR_WALL_HOSPITAL
            Case 4: GetTownSprite = SPR_WALL
            Case 5: GetTownSprite = SPR_DOOR  ' Hospital door
            Case 6: GetTownSprite = SPR_WALL
            ' Bank (cols 10-13)
            Case 10: GetTownSprite = SPR_WALL
            Case 11: GetTownSprite = SPR_DOOR  ' Bank door
            Case 12: GetTownSprite = SPR_WALL
            Case 13: GetTownSprite = SPR_WALL_BANK
            ' Saloon (cols 17-21)
            Case 17: GetTownSprite = SPR_WALL_DRINK
            Case 18: GetTownSprite = SPR_WALL
            Case 19: GetTownSprite = SPR_DOOR_SALOON  ' Saloon door
            Case 20: GetTownSprite = SPR_WALL
            Case 21: GetTownSprite = SPR_WALL_BROTHEL
            ' Store (cols 24-27)
            Case 24: GetTownSprite = SPR_WALL
            Case 25: GetTownSprite = SPR_DOOR  ' Store door
            Case 26: GetTownSprite = SPR_WALL
            Case 27: GetTownSprite = SPR_WALL_WHEEL
            Case 29: GetTownSprite = SPR_HITCHINGPOST
        End Select
        Exit Function
    End If
End Function

' ============================================================================
' Rendering
' ============================================================================
Private Sub RenderGame()
    Dim X As Integer, Y As Integer
    Dim ScreenX As Integer, ScreenY As Integer
    Dim CellType As Integer, Modifier As Integer

    ' Update viewport
    Call UpdateViewport

    ' Clear entire screen (game area + sidebar)
    picGame.Cls

    ' Draw grid
    For Y = ViewportY To ViewportY + VIEWPORT_ROWS - 1
        For X = ViewportX To ViewportX + VIEWPORT_COLS - 1
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

    ' Draw sidebar directly on picGame (at right side)
    Call DrawSidebar

    ' Refresh display
    picGame.Refresh
End Sub

Private Sub DrawCell(ByVal ScreenX As Integer, ByVal ScreenY As Integer, _
                     ByVal CellType As Integer, ByVal Modifier As Integer, _
                     ByVal GridX As Integer, ByVal GridY As Integer)

    Dim DrawColor As Long
    Dim SpriteIdx As Integer
    Dim TownSprite As Integer

    If UseImages And SpritesLoaded Then
        ' Draw using tileset sprites
        Select Case CellType
            Case CELL_AIR
                ' Check if this is in the town area (rows 0-3)
                TownSprite = GetTownSprite(GridX, GridY)
                If TownSprite >= 0 Then
                    Call DrawSprite(TownSprite, ScreenX, ScreenY)
                Else
                    Call DrawSprite(SPR_SKY, ScreenX, ScreenY)
                End If

            Case CELL_DOOR
                ' Check which building door this is
                TownSprite = GetTownSprite(GridX, GridY)
                If TownSprite >= 0 Then
                    Call DrawSprite(TownSprite, ScreenX, ScreenY)
                Else
                    Call DrawSprite(SPR_DOOR, ScreenX, ScreenY)
                End If

            Case CELL_DIRT
                ' Alternate dirt textures based on position
                If (GridX + GridY) Mod 2 = 0 Then
                    Call DrawSprite(SPR_DIRT, ScreenX, ScreenY)
                Else
                    Call DrawSprite(SPR_DIRT2, ScreenX, ScreenY)
                End If

                ' Show modifier hint if player has light
                If CanSeeModifier(GridX, GridY) And Modifier <> MOD_NONE Then
                    Call DrawModifierHint(ScreenX, ScreenY, Modifier)
                End If

            Case CELL_DUG
                Call DrawSprite(SPR_CLEARED, ScreenX, ScreenY)

            Case CELL_ELEVATOR, CELL_ELEVATOR_CAR
                ' Draw elevator shaft or car based on current elevator position
                If GridY = ElevatorY Then
                    ' Elevator car is here
                    Call DrawSprite(SPR_ELEVATOR_MIDDLE, ScreenX, ScreenY)
                ElseIf GridY = 3 Then
                    ' Top of shaft (above ground)
                    Call DrawSprite(SPR_ELEVATOR_TOP_ABOVE, ScreenX, ScreenY)
                ElseIf GridY = 4 Then
                    ' Road level shaft
                    Call DrawSprite(SPR_SHAFT, ScreenX, ScreenY)
                Else
                    ' Underground shaft
                    Call DrawSprite(SPR_SHAFT, ScreenX, ScreenY)
                End If

            Case CELL_ROAD
                Call DrawSprite(SPR_BORDER, ScreenX, ScreenY)

            Case CELL_WATER
                Call DrawSprite(SPR_WATER, ScreenX, ScreenY)

            Case CELL_GRANITE
                Call DrawSprite(SPR_GRANITE, ScreenX, ScreenY)

            Case CELL_CAVE
                Call DrawSprite(SPR_CAVEIN, ScreenX, ScreenY)

            Case CELL_WHIRLPOOL
                Call DrawSprite(SPR_WATER, ScreenX, ScreenY)

            Case CELL_SPRING
                Call DrawSprite(SPR_SPRING, ScreenX, ScreenY)

            Case CELL_SANDSTONE
                Call DrawSprite(SPR_SANDSTONE, ScreenX, ScreenY)

            Case CELL_VOLCANIC
                Call DrawSprite(SPR_VOLCANIC, ScreenX, ScreenY)

            Case CELL_WALL
                Call DrawSprite(SPR_WALL, ScreenX, ScreenY)

            Case CELL_ROOF
                Call DrawSprite(SPR_ROOF_MIDDLE, ScreenX, ScreenY)

            Case Else
                Call DrawSprite(SPR_BLACK, ScreenX, ScreenY)
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
            Case CELL_WATER, CELL_WHIRLPOOL: DrawColor = COLOR_WATER
            Case CELL_GRANITE: DrawColor = COLOR_ROCK
            Case CELL_CAVE: DrawColor = &H404080
            Case CELL_SPRING: DrawColor = COLOR_SPRING
            Case CELL_SANDSTONE: DrawColor = COLOR_SANDSTONE
            Case CELL_VOLCANIC: DrawColor = COLOR_VOLCANIC
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
    ' Draw sprite to show hidden modifier (using spritesheet images)
    Dim SpriteIdx As Integer

    ' Map modifier to sprite index from spritesheet
    Select Case Modifier
        Case MOD_SILVER: SpriteIdx = SPR_SILVER         ' Row 5, Col 4
        Case MOD_GOLD: SpriteIdx = SPR_GOLD             ' Row 5, Col 5
        Case MOD_PLATINUM: SpriteIdx = SPR_PLATINUM     ' Row 5, Col 6
        Case MOD_CAVEIN: SpriteIdx = SPR_CAVEIN         ' Row 5, Col 7
        Case MOD_WATER: SpriteIdx = SPR_WATER           ' Row 5, Col 3
        Case MOD_WHIRLPOOL: SpriteIdx = SPR_WATER       ' Row 5, Col 2 (use water sprite)
        Case MOD_GRANITE: SpriteIdx = SPR_GRANITE       ' Row 2, Col 5
        Case MOD_DIAMOND: SpriteIdx = SPR_DIAMOND       ' Row 6, Col 2
        Case MOD_PUMP: SpriteIdx = SPR_PUMP             ' Row 6, Col 2
        Case MOD_CLOVER: SpriteIdx = SPR_CLOVER         ' Row 1, Col 3
        Case MOD_SPRING: SpriteIdx = SPR_SPRING         ' Row 5, Col 2
        Case MOD_VOLCANIC: SpriteIdx = SPR_VOLCANIC     ' Row 6, Col 1
        Case MOD_SANDSTONE: SpriteIdx = SPR_SANDSTONE   ' Row 5, Col 8
        Case Else: Exit Sub
    End Select

    ' Draw the sprite at the cell position
    Call DrawSprite(SpriteIdx, ScreenX, ScreenY)
End Sub

Private Sub DrawPlayer()
    Dim ScreenX As Integer, ScreenY As Integer

    ScreenX = (Player.X - ViewportX) * CELL_WIDTH
    ScreenY = (Player.Y - ViewportY) * CELL_HEIGHT

    If UseImages And SpritesLoaded Then
        ' Use transparent drawing for player only (removes white background)
        If Player.Facing = FACING_LEFT Then
            Call DrawSpriteTransparent(SPR_PLAYER_LEFT, ScreenX, ScreenY)
        Else
            Call DrawSpriteTransparent(SPR_PLAYER_RIGHT, ScreenX, ScreenY)
        End If
    Else
        ' Draw player as colored rectangle
        picGame.Line (ScreenX + 2, ScreenY + 2)-(ScreenX + CELL_WIDTH - 3, ScreenY + CELL_HEIGHT - 3), COLOR_PLAYER, BF
    End If
End Sub

' ============================================================================
' Sidebar Drawing (draws directly to picGame at GAME_WIDTH offset)
' ============================================================================
Private Sub DrawSidebar()
    Dim SX As Integer  ' Sidebar X offset
    Dim Y As Integer
    Dim i As Integer
    Dim HealthColor As Long

    SX = GAME_WIDTH  ' Sidebar starts at x=512

    ' Draw sidebar background
    picGame.Line (SX, 0)-(SX + SIDEBAR_WIDTH - 1, GAME_HEIGHT - 1), &H202020, BF

    ' Set font
    picGame.FontName = "Consolas"
    picGame.FontSize = 8
    picGame.FontBold = True

    Y = 5

    ' Title
    picGame.ForeColor = vbYellow
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "MinerVGA"
    Y = Y + 20

    ' Separator
    picGame.Line (SX + 5, Y)-(SX + SIDEBAR_WIDTH - 10, Y), &H404040
    Y = Y + 8

    ' Minerals section
    picGame.ForeColor = COLOR_PLATINUM
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "PT: " & Player.Platinum
    Y = Y + 15

    picGame.ForeColor = COLOR_GOLD
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "AU: " & Player.Gold
    Y = Y + 15

    picGame.ForeColor = COLOR_SILVER
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "AG: " & Player.Silver
    Y = Y + 20

    ' Separator
    picGame.Line (SX + 5, Y)-(SX + SIDEBAR_WIDTH - 10, Y), &H404040
    Y = Y + 8

    ' Health section
    If Player.Health > 50 Then
        HealthColor = vbGreen
    ElseIf Player.Health > 20 Then
        HealthColor = vbYellow
    Else
        HealthColor = vbRed
    End If

    picGame.ForeColor = vbWhite
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "Health %"
    Y = Y + 15

    picGame.ForeColor = HealthColor
    picGame.FontSize = 12
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "   " & Player.Health
    picGame.FontSize = 8
    Y = Y + 25

    ' Separator
    picGame.Line (SX + 5, Y)-(SX + SIDEBAR_WIDTH - 10, Y), &H404040
    Y = Y + 8

    ' Bank Account
    picGame.ForeColor = vbCyan
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "Bank Accnt"
    Y = Y + 15

    If Player.Cash >= 0 Then
        picGame.ForeColor = vbCyan
    Else
        picGame.ForeColor = vbRed
    End If
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "$ " & Format(Player.Cash, "#,##0.00")
    Y = Y + 25

    ' Separator
    picGame.Line (SX + 5, Y)-(SX + SIDEBAR_WIDTH - 10, Y), &H404040
    Y = Y + 8

    ' Messages section
    picGame.ForeColor = vbWhite
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "Messages"
    Y = Y + 15

    picGame.FontSize = 7
    For i = 0 To 5
        If i < MessageCount And Messages(i) <> "" Then
            ' Color code messages
            If InStr(Messages(i), "Silver") > 0 Then
                picGame.ForeColor = COLOR_SILVER
            ElseIf InStr(Messages(i), "Gold") > 0 Then
                picGame.ForeColor = COLOR_GOLD
            ElseIf InStr(Messages(i), "Platinum") > 0 Then
                picGame.ForeColor = COLOR_PLATINUM
            ElseIf InStr(Messages(i), "Spring") > 0 Or InStr(Messages(i), "damage") > 0 Then
                picGame.ForeColor = vbRed
            ElseIf InStr(Messages(i), "found") > 0 Then
                picGame.ForeColor = vbYellow
            Else
                picGame.ForeColor = &H808080
            End If
            picGame.CurrentX = SX + 5
            picGame.CurrentY = Y
            picGame.Print Left(Messages(i), 16)
            Y = Y + 12
        End If
    Next i
    picGame.FontSize = 8
    Y = Y + 10

    ' Separator
    picGame.Line (SX + 5, Y)-(SX + SIDEBAR_WIDTH - 10, Y), &H404040
    Y = Y + 8

    ' Inventory section
    picGame.ForeColor = vbWhite
    picGame.CurrentX = SX + 5
    picGame.CurrentY = Y
    picGame.Print "You have"
    Y = Y + 18

    ' Draw inventory icons
    Call DrawSidebarInventory(SX, Y)
End Sub

Private Sub DrawSidebarInventory(ByVal SX As Integer, ByVal StartY As Integer)
    Dim X As Integer, Y As Integer
    Dim ItemCount As Integer
    Const ITEMS_PER_ROW As Integer = 5

    X = SX + 5
    Y = StartY
    ItemCount = 0

    ' Draw item icons for owned items in grid (5 per row)
    If HasShovel Then
        Call DrawSprite(SPR_SHOVEL, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasPickaxe Then
        Call DrawSprite(SPR_PICKAXE, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasDrill Then
        Call DrawSprite(SPR_DRILL, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasLantern Then
        Call DrawSprite(SPR_LAMP, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasBucket Then
        Call DrawSprite(SPR_BUCKET, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasTorch Then
        Call DrawSprite(SPR_TORCH, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasDynamite Then
        Call DrawSprite(SPR_DYNAMITE, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasRing Then
        Call DrawSprite(SPR_RING, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasCondom Then
        Call DrawSprite(SPR_CONDOM, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasPump Then
        Call DrawSprite(SPR_PUMP, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasClover Then
        Call DrawSprite(SPR_CLOVER, X, Y)
        ItemCount = ItemCount + 1
        If ItemCount Mod ITEMS_PER_ROW = 0 Then
            X = SX + 5
            Y = Y + CELL_HEIGHT + 2
        Else
            X = X + CELL_WIDTH + 2
        End If
    End If

    If HasDiamond Then
        Call DrawSprite(SPR_DIAMOND, X, Y)
        ItemCount = ItemCount + 1
    End If
End Sub

' ============================================================================
' Status Display
' ============================================================================
Private Sub UpdateStatus()
    Dim Depth As Integer

    Depth = GetDepthInFeet(Player.Y)

    ' Simple status on bottom bar
    lblStatus.ForeColor = vbGreen
    lblStatus.Caption = "Depth: " & Depth & " ft  |  Press H for Help  |  E to Enter Buildings"
End Sub

' ============================================================================
' Screen Displays
' ============================================================================
Private Sub ShowTitleScreen()
    picGame.Cls

    ' Draw title in game area
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

    ' Draw sidebar on title screen too
    Call DrawSidebar

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

    picGame.CurrentX = 150
    picGame.CurrentY = 280
    picGame.Print "Final wealth: $" & Format(Player.Cash, "#,##0")

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 350
    picGame.Print "Press any key to continue..."

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
    picGame.Print "With $" & Format(Player.Cash, "#,##0") & " and a diamond ring,"
    picGame.CurrentX = 100
    picGame.CurrentY = 280
    picGame.Print "you can now retire in style!"

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 380
    picGame.Print "Press any key to continue..."

    picGame.Refresh
    GameState = STATE_WON
    tmrGame.Enabled = False
End Sub

Private Sub ShowBankruptScreen()
    picGame.Cls
    picGame.ForeColor = vbRed
    picGame.FontSize = 24
    picGame.FontBold = True
    picGame.CurrentX = 150
    picGame.CurrentY = 100
    picGame.Print "BANKRUPT!"

    picGame.FontSize = 14
    picGame.FontBold = False
    picGame.ForeColor = vbWhite
    picGame.CurrentX = 100
    picGame.CurrentY = 200
    picGame.Print "You have run out of money!"

    picGame.CurrentX = 100
    picGame.CurrentY = 250
    picGame.Print "Final balance: $" & Format(Player.Cash, "#,##0")

    picGame.CurrentX = 100
    picGame.CurrentY = 280
    picGame.Print "The bank has foreclosed on your claim."

    picGame.ForeColor = vbGreen
    picGame.CurrentX = 150
    picGame.CurrentY = 380
    picGame.Print "Press any key to continue..."

    picGame.Refresh
End Sub

' ============================================================================
' Building Entry
' ============================================================================
Private Sub EnterBuilding()
    Dim DoorTarget As Integer

    ' Check if player is at a door
    If Grid(Player.X, Player.Y).CellType <> CELL_DOOR Then
        Call AddMessage("No door here!")
        Exit Sub
    End If

    DoorTarget = Grid(Player.X, Player.Y).DoorTarget

    tmrGame.Enabled = False

    Select Case DoorTarget
        Case BUILDING_OUTHOUSE
            frmOuthouse.Show vbModal
            Call AddMessage("Left outhouse")

        Case BUILDING_BANK
            frmBank.Show vbModal
            Call AddMessage("Left bank")

        Case BUILDING_STORE
            frmStore.Show vbModal
            Call AddMessage("Left store")

        Case BUILDING_HOSPITAL
            frmHospital.Show vbModal
            Call AddMessage("Left hospital")

        Case BUILDING_SALOON
            frmSaloon.Show vbModal
            Call AddMessage("Left saloon")

        Case Else
            Call AddMessage("Unknown building")
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
