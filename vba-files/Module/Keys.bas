Attribute VB_Name = "Keys"
Option Explicit

Sub bindKeys()
    Application.ScreenUpdating = False
    Application.OnKey "{LEFT}", "przesuniecieWLewo"
    Application.OnKey "{UP}", "obrot"
    Application.OnKey "{DOWN}", "przesuniecieWDol"
    Application.OnKey "{RIGHT}", "przesuniecieWPrawo"
    Application.OnKey "{ESC}", "StopGame"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub przesuniecieWLewo()
    If gameStarted = True Then
        If Not CzyKolizja(xFigury - 1, yFigury, nrPozFigury) Then
            xFigury = xFigury - 1
        End If
    End If
End Sub

Private Sub obrot()
    Dim delPozFigury As Byte

    If gameStarted = True Then
        If nrPozFigury + 1 <= 4 Then
            delPozFigury = nrPozFigury + 1
        Else
            delPozFigury = 1
        End If

        If Not CzyKolizja(xFigury, yFigury, delPozFigury) Then
            nrPozFigury = delPozFigury
        End If
    End If
End Sub

Private Sub przesuniecieWPrawo()
    If gameStarted = True Then
        If Not CzyKolizja(xFigury + 1, yFigury, nrPozFigury) Then
            xFigury = xFigury + 1
        End If
    End If
End Sub

Private Sub przesuniecieWDol()
    If gameStarted = True Then
        While Not CzyKolizja(xFigury, yFigury + 1, nrPozFigury)
            yFigury = yFigury + 1
        Wend
        Opoznienie = 0
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

