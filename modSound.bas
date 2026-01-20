Attribute VB_Name = "modSound"
' ============================================================================
' MinerVGA VB6 Edition by vbgamer45
' https://github.com/VBGAMER45/minervga-vb6clone
' https://www.theprogrammingzone.com/
' ============================================================================
Option Explicit

' ============================================================================
' MinerVGA - Sound System Module
' Uses Windows Beep API for simple sound effects
' ============================================================================

' --- Windows API Declaration ---
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' --- Sound Frequencies (Hz) ---
Private Const FREQ_DIG As Long = 200
Private Const FREQ_MINERAL As Long = 800
Private Const FREQ_DAMAGE As Long = 100
Private Const FREQ_ITEM As Long = 1000
Private Const FREQ_SPRING As Long = 300
Private Const FREQ_CAVEIN As Long = 80
Private Const FREQ_WIN As Long = 1200
Private Const FREQ_LOSE As Long = 150

' --- Sound Durations (ms) ---
Private Const DUR_SHORT As Long = 50
Private Const DUR_MEDIUM As Long = 100
Private Const DUR_LONG As Long = 200

' ============================================================================
' Sound Effect Procedures
' ============================================================================

Public Sub PlayDigSound()
    ' Low frequency thud for digging
    If Not SoundEnabled Then Exit Sub
    Beep FREQ_DIG, DUR_SHORT
End Sub

Public Sub PlayMineralSound()
    ' Pleasant chime for finding minerals
    If Not SoundEnabled Then Exit Sub
    Beep FREQ_MINERAL, DUR_MEDIUM
    Beep FREQ_MINERAL + 200, DUR_SHORT
End Sub

Public Sub PlayDamageSound()
    ' Low rumble for taking damage
    If Not SoundEnabled Then Exit Sub
    Beep FREQ_DAMAGE, DUR_MEDIUM
    Beep FREQ_DAMAGE + 20, DUR_SHORT
End Sub

Public Sub PlayItemSound()
    ' High pitched chirp for picking up items
    If Not SoundEnabled Then Exit Sub
    Beep FREQ_ITEM, DUR_SHORT
    Beep FREQ_ITEM + 300, DUR_SHORT
End Sub

Public Sub PlaySpringSound()
    ' Swoosh sound for spring/water
    If Not SoundEnabled Then Exit Sub
    Dim i As Integer
    For i = 140 To 190 Step 10
        Beep i, 30
    Next i
End Sub

Public Sub PlayCaveInSound()
    ' Deep rumble for cave-in
    If Not SoundEnabled Then Exit Sub
    Dim i As Integer
    For i = 1 To 5
        Beep FREQ_CAVEIN + (i * 5), DUR_SHORT
    Next i
End Sub

Public Sub PlayWinSound()
    ' Victory fanfare
    If Not SoundEnabled Then Exit Sub
    Beep 523, 150  ' C
    Beep 659, 150  ' E
    Beep 784, 150  ' G
    Beep 1047, 300 ' High C
End Sub

Public Sub PlayLoseSound()
    ' Sad descending tones
    If Not SoundEnabled Then Exit Sub
    Beep 400, 200
    Beep 300, 200
    Beep 200, 400
End Sub

Public Sub PlayPurchaseSound()
    ' Ka-ching for purchases
    If Not SoundEnabled Then Exit Sub
    Beep 600, 50
    Beep 900, 100
End Sub

Public Sub PlayElevatorSound()
    ' Mechanical elevator sound
    If Not SoundEnabled Then Exit Sub
    Beep 150, 100
    Beep 180, 100
End Sub

Public Sub PlayDynamiteSound()
    ' Explosion
    If Not SoundEnabled Then Exit Sub
    Beep 100, 50
    Beep 80, 100
    Beep 60, 150
End Sub

Public Sub PlayDrillSound()
    ' Drilling through rock
    If Not SoundEnabled Then Exit Sub
    Beep 250, 50
    Beep 300, 50
    Beep 250, 50
End Sub

Public Sub PlayPumpSound()
    ' Pumping water
    If Not SoundEnabled Then Exit Sub
    Beep 180, 100
    Beep 220, 100
End Sub

Public Sub PlaySandstoneSound()
    ' Soft crumbly sound for sandstone
    If Not SoundEnabled Then Exit Sub
    Beep 350, 40
    Beep 300, 30
End Sub

Public Sub PlayVolcanicSound()
    ' Hard crunchy sound for volcanic rock
    If Not SoundEnabled Then Exit Sub
    Beep 150, 60
    Beep 120, 40
    Beep 180, 50
End Sub

Public Sub PlayGraniteSound()
    ' Hard drilling sound for granite
    If Not SoundEnabled Then Exit Sub
    Beep 280, 40
    Beep 320, 40
    Beep 260, 40
End Sub

Public Sub PlayWaterSplash()
    ' Splash sound for water
    If Not SoundEnabled Then Exit Sub
    Beep 400, 30
    Beep 350, 40
    Beep 300, 50
End Sub
