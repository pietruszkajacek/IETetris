Attribute VB_Name = "Keys"
Option Explicit

Sub bindKeys()
    Application.ScreenUpdating = False
    Application.OnKey "{LEFT}", "tetroLeft"
    Application.OnKey "{UP}", "tetroRotate"
    Application.OnKey "{DOWN}", "tetroDown"
    Application.OnKey "{RIGHT}", "tetroRight"
    Application.OnKey "{ESC}", "StopGame"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub tetroLeft()
    If gameStarted = True Then
        If Not checkCollision(tetrominoX - 1, tetrominoY, tetrominoRot) Then
            tetrominoX = tetrominoX - 1
        End If
    End If
End Sub

Private Sub tetroRotate()
    Dim tetroRot As Byte

    If gameStarted = True Then
        If tetrominoRot + 1 <= 4 Then
            tetroRot = tetrominoRot + 1
        Else
            tetroRot = 1
        End If

        If Not checkCollision(tetrominoX, tetrominoY, tetroRot) Then
            tetrominoRot = tetroRot
        End If
    End If
End Sub

Private Sub tetroRight()
    If gameStarted = True Then
        If Not checkCollision(tetrominoX + 1, tetrominoY, tetrominoRot) Then
            tetrominoX = tetrominoX + 1
        End If
    End If
End Sub

Private Sub tetroDown()
    If gameStarted = True Then
        While Not checkCollision(tetrominoX, tetrominoY + 1, tetrominoRot)
            tetrominoY = tetrominoY + 1
        Wend
        delay = 0
    End If
End Sub

Sub freeKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{ESC}"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

