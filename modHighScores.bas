Attribute VB_Name = "modHighScores"
' ============================================================================
' MinerVGA VB6 Edition by vbgamer45
' https://github.com/VBGAMER45/minervga-vb6clone
' https://www.theprogrammingzone.com/
' ============================================================================
Option Explicit

' ============================================================================
' MinerVGA - High Scores Module
' ============================================================================

Public Const MAX_HIGH_SCORES As Integer = 10

Public Type HighScoreEntry
    PlayerName As String
    Score As Long
    ScoreDate As String
    Won As Boolean
End Type

Public HighScores(1 To 10) As HighScoreEntry

' ============================================================================
' Initialize High Scores
' ============================================================================
Public Sub InitHighScores()
    Dim i As Integer

    For i = 1 To MAX_HIGH_SCORES
        HighScores(i).PlayerName = ""
        HighScores(i).Score = 0
        HighScores(i).ScoreDate = ""
        HighScores(i).Won = False
    Next i

    ' Try to load existing scores
    Call LoadHighScores
End Sub

' ============================================================================
' Check if Score Qualifies for High Score List
' ============================================================================
Public Function IsHighScore(ByVal Score As Long) As Boolean
    ' A score qualifies if it's higher than the lowest score on the list
    ' or if the list isn't full yet
    If Score <= 0 Then
        IsHighScore = False
        Exit Function
    End If

    If HighScores(MAX_HIGH_SCORES).Score = 0 Then
        ' List isn't full
        IsHighScore = True
    ElseIf Score > HighScores(MAX_HIGH_SCORES).Score Then
        ' Better than lowest score
        IsHighScore = True
    Else
        IsHighScore = False
    End If
End Function

' ============================================================================
' Get Rank for a Score
' ============================================================================
Public Function GetScoreRank(ByVal Score As Long) As Integer
    Dim i As Integer

    If Score <= 0 Then
        GetScoreRank = 0
        Exit Function
    End If

    For i = 1 To MAX_HIGH_SCORES
        If Score > HighScores(i).Score Then
            GetScoreRank = i
            Exit Function
        End If
    Next i

    GetScoreRank = 0
End Function

' ============================================================================
' Insert a New High Score
' ============================================================================
Public Sub InsertHighScore(ByVal Rank As Integer, ByVal PlayerName As String, ByVal Score As Long)
    Dim i As Integer

    If Rank < 1 Or Rank > MAX_HIGH_SCORES Then Exit Sub

    ' Shift scores down
    For i = MAX_HIGH_SCORES To Rank + 1 Step -1
        HighScores(i) = HighScores(i - 1)
    Next i

    ' Insert new score
    HighScores(Rank).PlayerName = Left(PlayerName, 15)
    HighScores(Rank).Score = Score
    HighScores(Rank).ScoreDate = Format(Date, "mm/dd/yyyy")
    HighScores(Rank).Won = (GameState = STATE_WON)
End Sub

' ============================================================================
' Save High Scores to File
' ============================================================================
Public Sub SaveHighScores()
    Dim FilePath As String
    Dim FileNum As Integer
    Dim i As Integer

    FilePath = App.Path & "\HIGHSCORES.DAT"
    FileNum = FreeFile

    On Error GoTo SaveError

    Open FilePath For Output As #FileNum

    For i = 1 To MAX_HIGH_SCORES
        Print #FileNum, HighScores(i).PlayerName
        Print #FileNum, HighScores(i).Score
        Print #FileNum, HighScores(i).ScoreDate
        Print #FileNum, HighScores(i).Won
    Next i

    Close #FileNum
    Exit Sub

SaveError:
    Close #FileNum
End Sub

' ============================================================================
' Load High Scores from File
' ============================================================================
Public Sub LoadHighScores()
    Dim FilePath As String
    Dim FileNum As Integer
    Dim i As Integer

    FilePath = App.Path & "\HIGHSCORES.DAT"

    If Dir(FilePath) = "" Then Exit Sub

    FileNum = FreeFile

    On Error GoTo LoadError

    Open FilePath For Input As #FileNum

    For i = 1 To MAX_HIGH_SCORES
        If EOF(FileNum) Then Exit For
        Input #FileNum, HighScores(i).PlayerName
        Input #FileNum, HighScores(i).Score
        Input #FileNum, HighScores(i).ScoreDate
        Input #FileNum, HighScores(i).Won
    Next i

    Close #FileNum
    Exit Sub

LoadError:
    Close #FileNum
End Sub

' ============================================================================
' Show High Scores (called after game ends)
' ============================================================================
Public Sub ShowHighScores(Optional ByVal NewScore As Long = 0)
    Dim Qualifies As Boolean
    Dim Rank As Integer
    Dim PlayerName As String

    Qualifies = IsHighScore(NewScore)
    Rank = GetScoreRank(NewScore)

    ' If this is a qualifying high score, prompt for name
    If NewScore > 0 And Qualifies And Rank > 0 Then
        PlayerName = InputBox("NEW HIGH SCORE!" & vbCrLf & vbCrLf & _
                              "Your score: $" & Format(NewScore, "#,##0") & vbCrLf & _
                              "Rank: #" & Rank & vbCrLf & vbCrLf & _
                              "Enter your name:", "Hall of Fame", "Player")

        If PlayerName = "" Then PlayerName = "Anonymous"

        ' Insert the score
        Call InsertHighScore(Rank, PlayerName, NewScore)

        ' Save to file
        Call SaveHighScores
    End If

    ' Show the high scores list
    frmHighScores.Show vbModal
End Sub
