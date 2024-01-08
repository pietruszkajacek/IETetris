Attribute VB_Name = "Main"
Option Explicit

Public speed As Integer
Public gameStarted As Boolean
Public endGame As Boolean
Public delay As Byte

Public TetrominoX As Integer, tetrominoY As Integer
Public tetrominoNr As Byte, tetrominoRot As Byte
Public nextTetrominoNr As Byte, nextTetrominoRot As Byte

Public prevXTetro As Integer, prevYTetro As Integer, prevTetroRot As Integer, prevTetroNr As Integer
Public newCycle As Boolean
Public completeLines As Integer

Public level As Byte
Public tetroCounter As Long

Public Const MAX_LEVEL = 5
Public delayLevels(1 To MAX_LEVEL) As Integer
Public statTetrominoes(1 To NUMBER_OF_TETROMINOES) As Long

Public Const tetrisSheet = "IE Tetris"

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
    Dim r As Byte, c As Byte
    For r = 1 To HEIGHT_TETRION
        For c = 1 To WIDTH_TETRION
            Tetrion(r, c) = COLLISION_FREE_COLOR
        Next c
    Next r
End Sub

Sub drawTetrominoOnTetrion()
    Dim r As Integer, c As Integer
    Dim tetromino As TetrominoT
    tetromino = Tetrominoes(tetrominoNr).Rotations(tetrominoRot)

    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                Tetrion(r + tetrominoY - 1, c + TetrominoX - 1) = tetromino.color
            End If
        Next c
    Next r
End Sub

Sub drawTetrominoOnSheet(ByVal TetroNr As Integer, ByVal TetroRot As Integer, _
                         ByVal color As Long, ByVal TetroY As Integer, ByVal TetroX As Integer)
    Dim r As Integer, c As Integer
    Dim tetromino As TetrominoT
    tetromino = Tetrominoes(TetroNr).Rotations(TetroRot)

    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                Worksheets(tetrisSheet).Cells(r + TetroY - 1, TetroX + c - 1).Interior.color = color
            End If
        Next c
    Next r
End Sub

Sub drawTetroOnSheetRelTetrion(ByVal TetroNr As Integer, ByVal TetroRot As Integer, _
                                   ByVal color As Long, ByVal TetroY As Integer, ByVal TetroX As Integer)
    drawTetrominoOnSheet TetroNr, TetroRot, color, Y_TETRION + TetroY - 1, X_TETRION + TetroX - 1
End Sub

Sub drawRectOnSheet(ByVal width As Integer, ByVal height As Integer, _
                            ByVal Y As Integer, ByVal X As Integer, ByVal color As Long)
    Dim r As Byte, c As Byte
    For r = 1 To height
        For c = 1 To width
            Worksheets(tetrisSheet).Cells(Y + r - 1, X + c - 1).Interior.color = color
        Next c
    Next r
End Sub

Sub drawBackgroundTetrion()
    drawRectOnSheet WIDTH_TETRION, HEIGHT_TETRION, Y_TETRION, X_TETRION, COLLISION_FREE_COLOR
End Sub

Sub drawBackgroundNextTetro()
    drawRectOnSheet 4, 4, Y_TETRION, X_TETRION + WIDTH_TETRION + 1, COLLISION_FREE_COLOR
End Sub

Sub drawTetrionOnSheet(ByVal X As Byte, ByVal Y As Byte)
    Dim r As Integer, c As Integer
    'Application.ScreenUpdating = False
    For r = 1 To HEIGHT_TETRION
        For c = 1 To WIDTH_TETRION
            Worksheets(tetrisSheet).Cells(Y_TETRION + r - 1, X_TETRION + c - 1).Interior.color = Tetrion(r, c)
        Next c
    Next r
    'Application.ScreenUpdating = True
End Sub

Function checkCollision(ByVal TetroX As Integer, ByVal TetroY As Integer, TetroRot)
    Dim tetrionX As Integer, tetrionY As Integer
    Dim tetromino As TetrominoT
    Dim r As Integer, c As Integer
    tetromino = Tetrominoes(tetrominoNr).Rotations(TetroRot)
    
    For r = 1 To tetromino.height
        For c = 1 To tetromino.width
            If tetromino.matrix(r, c) = 1 Then
                tetrionX = TetroX + c - 1
                tetrionY = TetroY + r - 1
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
    Cells(r, c) = TetrominoX
    Cells(r, c + 1) = tetrominoY
    Cells(r + 1, c) = tetrominoNr
    Cells(r + 1, c + 1) = tetrominoRot
    Cells(r + 2, c) = delay
    Cells(r + 2, c + 1) = completeLines
    Cells(r + 3, c) = level
    Cells(r + 3, c + 1) = delayLevels(level)
    Cells(r + 4, c) = tetroCounter
End Sub

Sub Info2(frm As UserForm)
    With frm
        .TetroX = TetrominoX
        .TetroX = TetrominoX
        .TetroY = tetrominoY
        .TetroNr = tetrominoNr
        .TetroRot = tetrominoRot
        .delay = delay
        .completeLines = completeLines
        .level = level
        .delayLevel = delayLevels(level)
        .tetroCounter = tetroCounter
    End With
End Sub

Sub statTetro(frm As UserForm)
    With frm
        .statOLabel.Caption = statTetrominoes(1)
        .statILabel.Caption = statTetrominoes(2)
        .statTLabel.Caption = statTetrominoes(3)
        .statJLabel.Caption = statTetrominoes(4)
        .statLLabel.Caption = statTetrominoes(5)
        .statSLabel.Caption = statTetrominoes(6)
        .statZLabel.Caption = statTetrominoes(7)
    End With
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
    If Not checkCollision(TetrominoX, tetrominoY + 1, tetrominoRot) Then
        tetrominoY = tetrominoY + 1
    Else
        drawTetrominoOnTetrion
        delCompleteLines
        drawTetrionOnSheet 1, 1

        tetrominoNr = nextTetrominoNr
        tetrominoRot = nextTetrominoRot

        nextTetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
        nextTetrominoRot = randomNumber(4, 1)

        statTetrominoes(tetrominoNr) = statTetrominoes(tetrominoNr) + 1

        TetrominoX = 5
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

        If checkCollision(TetrominoX, tetrominoY, tetrominoRot) Then
            endGame = True
        End If
    End If
End Sub

Sub drawNextTetromino(ByVal Y As Integer, ByVal X As Integer)
    drawTetrominoOnSheet tetrominoNr, tetrominoRot, COLLISION_FREE_COLOR, Y, X
    drawTetrominoOnSheet nextTetrominoNr, nextTetrominoRot, Tetrominoes(nextTetrominoNr).Rotations(nextTetrominoRot).color, _
        Y, X
End Sub

Sub mainLoop()
    Dim color As Long

    If gameStarted And Not endGame Then
        If delay = delayLevels(level) Then
            updateGame
            delay = 0
        End If

        If newCycle Then
            If endGame Then
                color = RGB(90, 90, 90)
            Else
                color = Tetrominoes(tetrominoNr).Rotations(tetrominoRot).color
            End If

            drawTetroOnSheetRelTetrion tetrominoNr, tetrominoRot, color, tetrominoY, TetrominoX
            
            drawNextTetromino Y_TETRION, X_TETRION + WIDTH_TETRION + 1

            newCycle = False
        ElseIf prevXTetro <> TetrominoX Or prevYTetro <> tetrominoY Or prevTetroRot <> tetrominoRot Or prevTetroNr <> tetrominoNr Then
            drawTetroOnSheetRelTetrion prevTetroNr, prevTetroRot, COLLISION_FREE_COLOR, prevYTetro, prevXTetro
            drawTetroOnSheetRelTetrion tetrominoNr, tetrominoRot, Tetrominoes(tetrominoNr).Rotations(tetrominoRot).color, tetrominoY, TetrominoX
        End If

        prevXTetro = TetrominoX
        prevYTetro = tetrominoY
        prevTetroRot = tetrominoRot
        prevTetroNr = tetrominoNr

        'Info 22, 1
        Info2 MainUserForm
        statTetro MainUserForm
        delay = delay + 1
    Else
        stopTimer
        freeKeys

        MsgBox "KONIEC GRY!", _
            VBA.vbMsgBoxStyle.vbInformation, _
            "IE Tetris"
        
        If completeLines > 0 Then
            If MsgBox("Tw�j wynik pozwala na dopisanie do listy najlepszych." & vbCrLf & _
                "Czy chcesz to zrobi�?", vbYesNo Or vbDefaultButton1, "Lista najlepszych graczy") = vbYes Then
                BestPlayersUserForm.show
            End If
        End If
    End If
End Sub

Sub startGame()
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
    TetrominoX = 5
    delay = 0

    tetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
    tetrominoRot = randomNumber(4, 1)

    nextTetrominoNr = randomNumber(NUMBER_OF_TETROMINOES, 1)
    nextTetrominoRot = randomNumber(4, 1)

    Erase statTetrominoes
    statTetrominoes(tetrominoNr) = statTetrominoes(tetrominoNr) + 1

    newCycle = True
    completeLines = 0
    endGame = False
    gameStarted = True
    startTimer
End Sub

Sub stopGame()
  gameStarted = False
  stopTimer
  freeKeys
  'InfoUserForm.Hide
  'MainUserForm.show
End Sub

