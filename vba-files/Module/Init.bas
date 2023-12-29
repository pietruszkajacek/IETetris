Attribute VB_Name = "Init"
Option Explicit

Type FiguraT
    Matrix(1 To 4, 1 To 4) As Byte
    szerFigury As Byte
    wysFigury As Byte
    kolor As Long
End Type

Type FiguraObrotyT
    Obroty(1 To 4) As FiguraT
End Type

Public Const LICZBA_FIGUR = 7
Public Const SZER_PLANSZY = 10
Public Const WYS_PLANSZY = 20

Public Const X_PLANSZY = 2
Public Const Y_PLANSZY = 2

Public Const KOLOR_BEZKOL = 0 '16777215 'RGB(255, 255, 255)

Dim O0 As FiguraT
Dim Oobr As FiguraObrotyT

Dim I0 As FiguraT, I90 As FiguraT, I270 As FiguraT, I360 As FiguraT
Dim Iobr As FiguraObrotyT

Dim T0 As FiguraT, T90 As FiguraT, T270 As FiguraT, T360 As FiguraT
Dim Tobr As FiguraObrotyT

Dim J0 As FiguraT, J90 As FiguraT, J270 As FiguraT, J360 As FiguraT
Dim Jobr As FiguraObrotyT

Dim L0 As FiguraT, L90 As FiguraT, L270 As FiguraT, L360 As FiguraT
Dim Lobr As FiguraObrotyT

Dim S0 As FiguraT, S90 As FiguraT, S270 As FiguraT, S360 As FiguraT
Dim Sobr As FiguraObrotyT

Dim Z0 As FiguraT, Z90 As FiguraT, Z270 As FiguraT, Z360 As FiguraT
Dim Zobr As FiguraObrotyT

Public Tetromino(1 To LICZBA_FIGUR) As FiguraObrotyT
Public Plansza(1 To WYS_PLANSZY, 1 To SZER_PLANSZY) As Long

Sub initTetromino()
    With O0
        .Matrix(1, 1) = 1
        .Matrix(1, 2) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 255, 0)
    End With
        
    With Oobr
        .Obroty(1) = O0
        .Obroty(2) = O0
        .Obroty(3) = O0
        .Obroty(4) = O0
    End With

    With I0
        .Matrix(1, 2) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .Matrix(4, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 255)
    End With

    With I90
        .Matrix(3, 1) = 1
        .Matrix(3, 2) = 1
        .Matrix(3, 3) = 1
        .Matrix(3, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 255)
    End With

    With I270
        .Matrix(1, 3) = 1
        .Matrix(2, 3) = 1
        .Matrix(3, 3) = 1
        .Matrix(4, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 255)
    End With

    With I360
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .Matrix(2, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 255)
    End With

    With Iobr
        .Obroty(1) = I0
        .Obroty(2) = I90
        .Obroty(3) = I270
        .Obroty(4) = I360
    End With

    With T0
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .Matrix(3, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 255)
    End With

    With T90
        .Matrix(1, 2) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .Matrix(2, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 255)
    End With

    With T270
        .Matrix(1, 2) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 255)
    End With

    With T360
        .Matrix(1, 2) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 255)
    End With

    With Tobr
        .Obroty(1) = T0
        .Obroty(2) = T90
        .Obroty(3) = T270
        .Obroty(4) = T360
    End With

    With J0
        .Matrix(1, 3) = 1
        .Matrix(2, 3) = 1
        .Matrix(3, 3) = 1
        .Matrix(4, 3) = 1
        .Matrix(4, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 0, 255)
    End With

    With J90
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .Matrix(2, 4) = 1
        .Matrix(3, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 0, 255)
    End With

    With J270
        .Matrix(1, 2) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .Matrix(4, 2) = 1
        .Matrix(1, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 0, 255)
    End With

    With J360
        .Matrix(2, 1) = 1
        .Matrix(3, 1) = 1
        .Matrix(3, 2) = 1
        .Matrix(3, 3) = 1
        .Matrix(3, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 0, 255)
    End With

    With Jobr
        .Obroty(1) = J0
        .Obroty(2) = J90
        .Obroty(3) = J270
        .Obroty(4) = J360
    End With

    With L0
        .Matrix(1, 2) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .Matrix(4, 2) = 1
        .Matrix(4, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 129, 0)
    End With

    With L90
        .Matrix(3, 1) = 1
        .Matrix(3, 2) = 1
        .Matrix(3, 3) = 1
        .Matrix(3, 4) = 1
        .Matrix(2, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 129, 0)
    End With

    With L270
        .Matrix(1, 2) = 1
        .Matrix(1, 3) = 1
        .Matrix(2, 3) = 1
        .Matrix(3, 3) = 1
        .Matrix(4, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 129, 0)
    End With

    With L360
        .Matrix(3, 1) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .Matrix(2, 4) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 129, 0)
    End With

    With Lobr
        .Obroty(1) = L0
        .Obroty(2) = L90
        .Obroty(3) = L270
        .Obroty(4) = L360
    End With

    With S0
        .Matrix(1, 2) = 1
        .Matrix(1, 3) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 0)
    End With

    With S90
        .Matrix(1, 1) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 0)
    End With

    With S270
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .Matrix(3, 1) = 1
        .Matrix(3, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 0)
    End With

    With S360
        .Matrix(1, 1) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(0, 255, 0)
    End With

    With Sobr
        .Obroty(1) = S0
        .Obroty(2) = S90
        .Obroty(3) = S270
        .Obroty(4) = S360
    End With

    With Z0
        .Matrix(1, 1) = 1
        .Matrix(1, 2) = 1
        .Matrix(2, 2) = 1
        .Matrix(2, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 0)
    End With

    With Z90
        .Matrix(1, 2) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 1) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 0)
    End With

    With Z270
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 2) = 1
        .Matrix(3, 3) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 0)
    End With

    With Z360
        .Matrix(1, 2) = 1
        .Matrix(2, 1) = 1
        .Matrix(2, 2) = 1
        .Matrix(3, 1) = 1
        .szerFigury = 4
        .wysFigury = 4
        .kolor = RGB(255, 0, 0)
    End With

    With Zobr
        .Obroty(1) = Z0
        .Obroty(2) = Z90
        .Obroty(3) = Z270
        .Obroty(4) = Z360
    End With

    Tetromino(1) = Oobr
    Tetromino(2) = Iobr
    Tetromino(3) = Tobr
    Tetromino(4) = Jobr
    Tetromino(5) = Lobr
    Tetromino(6) = Sobr
    Tetromino(7) = Zobr
End Sub



