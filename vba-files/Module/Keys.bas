Attribute VB_Name = "Keys"
Option Explicit

Sub bindKeys()
    'Application.ScreenUpdating = False
    Application.OnKey "{LEFT}", "tetroLeft"
    Application.OnKey "{UP}", "tetroRotate"
    Application.OnKey "{DOWN}", "tetroDown"
    Application.OnKey "{RIGHT}", "tetroRight"
    Application.OnKey "{ESC}", "StopGame"
    'Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub tetroLeft()
    If gameStarted = True Then
        If Not checkCollision(TetrominoX - 1, tetrominoY, tetrominoRot) Then
            TetrominoX = TetrominoX - 1
        End If
    End If
End Sub

Private Sub tetroRotate()
    Dim TetroRot As Byte

    If gameStarted = True Then
        If tetrominoRot + 1 <= 4 Then
            TetroRot = tetrominoRot + 1
        Else
            TetroRot = 1
        End If

        If Not checkCollision(TetrominoX, tetrominoY, TetroRot) Then
            tetrominoRot = TetroRot
        End If
    End If
End Sub

Private Sub tetroRight()
    If gameStarted = True Then
        If Not checkCollision(TetrominoX + 1, tetrominoY, tetrominoRot) Then
            TetrominoX = TetrominoX + 1
        End If
    End If
End Sub

Private Sub tetroDown()
    If gameStarted = True Then
        If Not checkCollision(TetrominoX, tetrominoY + 1, tetrominoRot) Then
            While Not checkCollision(TetrominoX, tetrominoY + 1, tetrominoRot)
                tetrominoY = tetrominoY + 1
            Wend
            delay = 0
        End If
    End If
End Sub

Sub freeKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{ESC}"
    'Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

