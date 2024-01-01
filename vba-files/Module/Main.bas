Attribute VB_Name = "Main"
Option Explicit

Public speed As Integer
Public gameStarted As Boolean
Public endGame As Boolean
Public delay As Byte

Public tetrominoX As Integer, tetrominoY As Integer
Public tetrominoNr As Byte, tetrominoRot As Byte
Public nextTetrominoNr As Byte, nextTetrominoRot As Byte

Public prevXTetro As Integer, prevYTetro As Integer, prevTetroRot As Integer, prevTetroNr As Integer
Public newCycle As Boolean
Public completeLines As Integer

Public level As Byte
Public tetroCounter as Long

Public Const MAX_LEVEL = 5
Public delayLevels(1 To MAX_LEVEL) As Integer

Sub initDelayLevels()
    delayLevels(1) = 10
    delayLevels(2) = 8
    delayLevels(3) = 6
    delayLevels(4) = 4
    delayLevels(5) = 2
End Sub

Function randomNumber(ByVal upperbound As Integer, ByVal lowerbound As Integer) As Integer
    Randomize
    randomNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound) 'start + Round(Rnd * (end - 1), 0)
End Function

Sub resetTetrion()
    Dim r, c As Byte
    For r = 1 To HEIGHT_TETRION
        For c = 1 To WIDTH_TETRION
            Tetrion(r, c) = COLLISION_FREE_COLOR
        Next c
    Next r
End Sub

Sub drawTetrominoOnTetrion()
    Dim r, c As Integer
    Dim tetromino As TetrominoT
    tetromino = Tetrominoes(tetrominoNr).Rotations(tetrominoRot)

    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                Tetrion(r + tetrominoY - 1, c + tetrominoX - 1) = tetromino.Color
            End If
        Next c
    Next r
End Sub

Sub drawTetrominoOnSheet(ByVal tetroNr As Integer, ByVal tetroRot As Integer, _
                         ByVal Color As Long, ByVal tetroY As Integer, ByVal tetroX As Integer)
    Dim r, c As Integer
    Dim tetromino As TetrominoT
    tetromino = Tetrominoes(tetroNr).Rotations(tetroRot)

    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                Cells(r + tetroY - 1, tetroX + c - 1).Interior.Color = Color
            End If
        Next c
    Next r
End Sub

Sub drawTetroOnSheetRelTetrion(ByVal tetroNr As Integer, ByVal tetroRot As Integer, _
                                   ByVal Color As Long, ByVal tetroY As Integer, ByVal tetroX As Integer)
    drawTetrominoOnSheet tetroNr, tetroRot, Color, Y_TETRION + tetroY - 1, X_TETRION + tetroX - 1
End Sub

Sub drawRectOnSheet(ByVal width As Integer, ByVal height As Integer, _
                            ByVal y As Integer, ByVal x As Integer, ByVal Color As Long)
    Dim r, c As Byte
    For r = 1 To height
        For c = 1 To width
            Cells(y + r - 1, x + c - 1).Interior.Color = Color
        Next c
    Next r
End Sub

Sub drawBackgroundTetrion()
    drawRectOnSheet WIDTH_TETRION, HEIGHT_TETRION, Y_TETRION, X_TETRION, COLLISION_FREE_COLOR
End Sub

Sub drawBackgroundNextTetro()
    drawRectOnSheet 4, 4, Y_TETRION, X_TETRION + WIDTH_TETRION + 1, COLLISION_FREE_COLOR
End Sub

Sub drawTetrionOnSheet(ByVal x As Byte, ByVal y As Byte)
    Dim r, c As Integer
    'Application.ScreenUpdating = False
    For r = 1 To HEIGHT_TETRION
        For c = 1 To WIDTH_TETRION
            Cells(Y_TETRION + r - 1, X_TETRION + c - 1).Interior.Color = Tetrion(r, c)
        Next c
    Next r
    'Application.ScreenUpdating = True
End Sub

Function checkCollision(ByVal tetroX As Integer, ByVal tetroY As Integer, tetroRot)
    Dim tetrionX, tetrionY As Integer
    Dim tetromino As TetrominoT
    Dim r, c As Integer
    tetromino = Tetrominoes(tetrominoNr).Rotations(tetroRot)
    
    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                tetrionX = tetroX + c - 1
                tetrionY = tetroY + r - 1
                If tetrionX > WIDTH_TETRION Or tetrionY > HEIGHT_TETRION Or tetrionX < 1 Then
                    checkCollision = True
                    Exit Function
                ElseIf Tetrion(tetrionY, tetrionX) <> COLLISION_FREE_COLOR Then
                    checkCollision = True
                    Exit Function
                End If
            End If
        Next c
    Next r
    checkCollision = False
End Function

Sub Info(ByVal r As Byte, ByVal c As Byte)
    Cells(r, c) = tetrominoX
    Cells(r, c + 1) = tetrominoY
    Cells(r + 1, c) = tetrominoNr
    Cells(r + 1, c + 1) = tetrominoRot
    Cells(r + 2, c) = delay
    Cells(r + 2, c + 1) = completeLines
    Cells(r + 3, c) = level
    Cells(r + 3, c + 1) = delayLevels(level)
    Cells(r + 4, c) = tetroCounter
End Sub

Function completeLine(ByVal r As Byte)
    Dim c As Byte, collisionFreeColor As Boolean
    collisionFreeColor = False

    For c = 1 To WIDTH_TETRION
        If Tetrion(r, c) = COLLISION_FREE_COLOR Then
            collisionFreeColor = True
            Exit For
        End If
    Next c
    
    completeLine = Not collisionFreeColor
End Function

Sub delCompleteLines()
    Dim rs As Byte, rd As Byte, c As Byte
    rs = HEIGHT_TETRION
    rd = HEIGHT_TETRION
    
    Do
        If Not completeLine(rs) Then
            For c = 1 To WIDTH_TETRION
                Tetrion(rd, c) = Tetrion(rs, c)
            Next c
            rs = rs - 1
            rd = rd - 1
        Else
            completeLines = completeLines + 1
            rs = rs - 1
        End If
    Loop Until rs = 1
End Sub

Sub updateGame()
    If Not checkCollision(tetrominoX, tetrominoY + 1, tetrominoRot) Then
        tetrominoY = tetrominoY + 1
    Else
        drawTetrominoOnTetrion
        delCompleteLines
        drawTetrionOnSheet 1, 1

        tetrominoNr = nextTetrominoNr
        tetrominoRot = nextTetrominoRot

        nextTetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
        nextTetrominoRot = randomNumber(4, 1)

        tetrominoX = 1
        tetrominoY = 1
        newCycle = True
        tetroCounter = tetroCounter + 1

        Select Case tetroCounter
            Case 1 To 29
                level = 1
            Case 30 To 79
                level = 2
            Case 80 To 129
                level = 3
            Case 130 To 179
                level = 4
            Case Else
                level = 5
        End Select

        If checkCollision(tetrominoX, tetrominoY, tetrominoRot) Then
            endGame = True
        End If
    End If
End Sub

Sub drawNextTetromino(ByVal y As Integer, ByVal x As Integer)
    drawTetrominoOnSheet tetrominoNr, tetrominoRot, COLLISION_FREE_COLOR, y, x
    drawTetrominoOnSheet nextTetrominoNr, nextTetrominoRot, Tetrominoes(nextTetrominoNr).Rotations(nextTetrominoRot).Color, _
        y, x
End Sub

Sub mainLoop()
    Dim Color As Long

    If gameStarted And Not endGame Then
        If delay = delayLevels(level) Then
            updateGame
            delay = 0
        End If

        If newCycle Then
            If endGame Then
                Color = RGB(90, 90, 90)
            Else
                Color = Tetrominoes(tetrominoNr).Rotations(tetrominoRot).Color
            End If

            drawTetroOnSheetRelTetrion tetrominoNr, tetrominoRot, Color, tetrominoY, tetrominoX
            
            drawNextTetromino Y_TETRION, X_TETRION + WIDTH_TETRION + 1

            newCycle = False
        ElseIf prevXTetro <> tetrominoX Or prevYTetro <> tetrominoY Or prevTetroRot <> tetrominoRot Or prevTetroNr <> tetrominoNr Then
            drawTetroOnSheetRelTetrion prevTetroNr, prevTetroRot, COLLISION_FREE_COLOR, prevYTetro, prevXTetro
            drawTetroOnSheetRelTetrion tetrominoNr, tetrominoRot, Tetrominoes(tetrominoNr).Rotations(tetrominoRot).Color, tetrominoY, tetrominoX
        End If

        prevXTetro = tetrominoX
        prevYTetro = tetrominoY
        prevTetroRot = tetrominoRot
        prevTetroNr = tetrominoNr

        Info 22, 1
        delay = delay + 1
    Else
        StopTimer
        MsgBox "Koniec GRY!"
    End If
End Sub

Sub StartGame()
    initTetrominoes
    initDelayLevels
    resetTetrion
    drawBackgroundTetrion
    drawBackgroundNextTetro

    speed = 50
    level = 1
    tetroCounter = 1

    Keys.bindKeys
    tetrominoY = 1
    tetrominoX = 1
    delay = 0
    tetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
    tetrominoRot = randomNumber(4, 1)

    nextTetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
    nextTetrominoRot = randomNumber(4, 1)

    newCycle = True
    completeLines = 0

    'If TimerID = 0 Then
        endGame = False
        gameStarted = True
        StartTimer
    'End If
End Sub

Sub StopGame()
  gameStarted = False
  Timer.StopTimer
  Keys.freeKeys
End Sub

