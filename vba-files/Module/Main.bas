Attribute VB_Name = "Main"
Option Explicit

Public speed As Integer
Public gameStarted As Boolean
Public koniecGry As Boolean
Public Opoznienie As Byte

Public xFigury As Integer, yFigury As Integer
Public nrFigury As Byte, nrPozFigury As Byte
Public nrNastFigury As Byte, nrNastPozFigury As Byte

Public tmpx As Integer, tmpy As Integer, tmppoz As Integer, tmpnrfig As Integer
Public nowyCykl As Boolean
Public wierszePelne As Integer

Function liczbaLosowa(ByVal koniec As Byte, ByVal poczatek As Byte) As Integer
    Randomize
    liczbaLosowa = poczatek + Round(Rnd * (koniec - 1), 0)
End Function

Sub ResetPlanszy()
    Dim w, k As Byte
    For w = 1 To WYS_PLANSZY
        For k = 1 To SZER_PLANSZY
            Plansza(w, k) = KOLOR_BEZKOL
        Next k
    Next w
End Sub

Sub rysujFigureNaPlanszy()
    Dim w, k As Integer
    Dim figura As FiguraT
    figura = Tetromino(nrFigury).Obroty(nrPozFigury)

    For w = 1 To figura.wysFigury
        For k = 1 To figura.szerFigury
            If figura.Matrix(w, k) = 1 Then
                Plansza(w + yFigury - 1, k + xFigury - 1) = figura.kolor
            End If
        Next k
    Next w
End Sub

Sub rysujFigureNaArkuszu(ByVal nrFig As Integer, ByVal pozFig As Integer, _
                         ByVal kolor As Long, ByVal yFig As Integer, ByVal xFig As Integer)
    Dim w, k As Integer
    Dim figura As FiguraT
    figura = Tetromino(nrFig).Obroty(pozFig)

    For w = 1 To figura.wysFigury
        For k = 1 To figura.szerFigury
            If figura.Matrix(w, k) = 1 Then
                Cells(w + yFig - 1, xFig + k - 1).Interior.Color = kolor
            End If
        Next k
    Next w
End Sub

Sub rysujFigureNaArkuszuWzgPlanszy(ByVal nrFig As Integer, ByVal pozFig As Integer, _
                                   ByVal kolor As Long, ByVal yFig As Integer, ByVal xFig As Integer)
    rysujFigureNaArkuszu nrFig, pozFig, kolor, Y_PLANSZY + yFig - 1, X_PLANSZY + xFig - 1
End Sub

Sub RysujProstokatNaArkuszu(ByVal szer As Integer, ByVal wys As Integer, _
                            ByVal y As Integer, ByVal x As Integer, ByVal kolor As Long)
    Dim w, k As Byte
    For w = 1 To wys
        For k = 1 To szer
            Cells(y + w - 1, x + k - 1).Interior.Color = kolor
        Next k
    Next w
End Sub

Sub RysujTloPlanszy()
    RysujProstokatNaArkuszu SZER_PLANSZY, WYS_PLANSZY, Y_PLANSZY, X_PLANSZY, KOLOR_BEZKOL
End Sub

Sub RysujTloNastFigury()
    RysujProstokatNaArkuszu 4, 4, Y_PLANSZY, X_PLANSZY + SZER_PLANSZY + 1, KOLOR_BEZKOL
End Sub

Sub RysujPlansze(ByVal x As Byte, ByVal y As Byte)
    Dim w, k As Integer
    'Application.ScreenUpdating = False
    For w = 1 To WYS_PLANSZY
        For k = 1 To SZER_PLANSZY
            Cells(Y_PLANSZY + w - 1, X_PLANSZY + k - 1).Interior.Color = Plansza(w, k)
        Next k
    Next w
    'Application.ScreenUpdating = True
End Sub

Function CzyKolizja(ByVal xDelFigury As Integer, ByVal yDelFigury As Integer, pozDelFigury)
    Dim xPlanszy, yPlanszy As Integer
    Dim figura As FiguraT
    Dim y, x As Integer
    figura = Tetromino(nrFigury).Obroty(pozDelFigury)
    
    For y = 1 To figura.wysFigury
        For x = 1 To figura.szerFigury
            If figura.Matrix(y, x) = 1 Then
                xPlanszy = xDelFigury + x - 1
                yPlanszy = yDelFigury + y - 1
                If xPlanszy > SZER_PLANSZY Or yPlanszy > WYS_PLANSZY Or xPlanszy < 1 Then
                    CzyKolizja = True
                    Exit Function
                ElseIf Plansza(yPlanszy, xPlanszy) <> KOLOR_BEZKOL Then
                    CzyKolizja = True
                    Exit Function
                End If
            End If
        Next x
    Next y
    CzyKolizja = False
End Function

Sub Info(ByVal w As Byte, ByVal k As Byte)
    Cells(w, k) = xFigury
    Cells(w, k + 1) = yFigury
    Cells(w + 1, k) = nrFigury
    Cells(w + 1, k + 1) = nrPozFigury
    Cells(w + 2, k) = Opoznienie
    Cells(w + 2, k + 1) = wierszePelne
End Sub

Function czyWierszPelny(ByVal w As Byte)
    Dim k As Byte, wystapilKolorBezkol As Boolean
    wystapilKolorBezkol = False

    For k = 1 To SZER_PLANSZY
        If Plansza(w, k) = KOLOR_BEZKOL Then
            wystapilKolorBezkol = True
            Exit For
        End If
    Next k
    
    czyWierszPelny = Not wystapilKolorBezkol
End Function

Sub redukujPelneWiersze()
    Dim wz As Byte, wd As Byte, k As Byte
    wz = WYS_PLANSZY
    wd = WYS_PLANSZY
    
    Do
        If Not czyWierszPelny(wz) Then
            For k = 1 To SZER_PLANSZY
                Plansza(wd, k) = Plansza(wz, k)
            Next k
            wz = wz - 1
            wd = wd - 1
        Else
            wierszePelne = wierszePelne + 1
            wz = wz - 1
        End If
    Loop Until wz = 1
End Sub

Sub aktualizujStanGry()
    If Not CzyKolizja(xFigury, yFigury + 1, nrPozFigury) Then
        yFigury = yFigury + 1
    Else
        rysujFigureNaPlanszy
        redukujPelneWiersze
        RysujPlansze 1, 1

        nrFigury = nrNastFigury
        nrPozFigury = nrNastPozFigury

        nrNastFigury = liczbaLosowa(LICZBA_FIGUR, 1)
        nrNastPozFigury = liczbaLosowa(4, 1)

        xFigury = 1
        yFigury = 1
        nowyCykl = True
        
        If CzyKolizja(xFigury, yFigury, nrPozFigury) Then
            koniecGry = True
        End If
    End If
End Sub

Sub rysujNastepnaFigure(ByVal y As Integer, ByVal x As Integer)
    rysujFigureNaArkuszu nrFigury, nrPozFigury, KOLOR_BEZKOL, y, x
    rysujFigureNaArkuszu nrNastFigury, nrNastPozFigury, Tetromino(nrNastFigury).Obroty(nrNastPozFigury).kolor, _
        y, x
End Sub

Sub Petla()
    Dim kolor As Long

    If gameStarted And Not koniecGry Then
        If Opoznienie = 10 Then
            aktualizujStanGry
            Opoznienie = 0
        End If

        If nowyCykl Then
            If koniecGry Then
                kolor = RGB(90, 90, 90)
            Else
                kolor = Tetromino(nrFigury).Obroty(nrPozFigury).kolor
            End If

            rysujFigureNaArkuszuWzgPlanszy nrFigury, nrPozFigury, kolor, yFigury, xFigury
            
            rysujNastepnaFigure Y_PLANSZY, X_PLANSZY + SZER_PLANSZY + 1

            nowyCykl = False
        ElseIf tmpx <> xFigury Or tmpy <> yFigury Or tmppoz <> nrPozFigury Or tmpnrfig <> nrFigury Then
            rysujFigureNaArkuszuWzgPlanszy tmpnrfig, tmppoz, KOLOR_BEZKOL, tmpy, tmpx
            rysujFigureNaArkuszuWzgPlanszy nrFigury, nrPozFigury, Tetromino(nrFigury).Obroty(nrPozFigury).kolor, yFigury, xFigury
        End If

        tmpx = xFigury
        tmpy = yFigury
        tmppoz = nrPozFigury
        tmpnrfig = nrFigury

        Info 22, 1
        Opoznienie = Opoznienie + 1
    Else
        StopTimer
        MsgBox "Koniec GRY!"
    End If
End Sub

Sub StartGame()
    initTetromino
    ResetPlanszy
    RysujTloPlanszy
    RysujTloNastFigury

    speed = 50
    Keys.bindKeys
    yFigury = 1
    xFigury = 1
    Opoznienie = 0
    nrFigury = liczbaLosowa(LICZBA_FIGUR, 1)
    nrPozFigury = liczbaLosowa(4, 1)

    nrNastFigury = liczbaLosowa(LICZBA_FIGUR, 1)
    nrNastPozFigury = liczbaLosowa(4, 1)

    nowyCykl = True
    wierszePelne = 0

    'If TimerID = 0 Then
        koniecGry = False
        gameStarted = True
        StartTimer
    'End If
End Sub

Sub StopGame()
  gameStarted = False
  Timer.StopTimer
  Keys.freeKeys
End Sub

