Attribute VB_Name = "Init"
Option Explicit

Type TetrominoT
    matrix(1 To 4, 1 To 4) As Byte
    width As Byte
    height As Byte
    Color As Long
End Type

Type TetrominoRotationsT
    Rotations(1 To 4) As TetrominoT
End Type

Public Const NUMBER_OF_TETROMINOES = 7
Public Const WIDTH_TETRION = 10
Public Const HEIGHT_TETRION = 20

Public Const X_TETRION = 2
Public Const Y_TETRION = 2

Public Const COLLISION_FREE_COLOR = 16777215 'RGB(255, 255, 255)

Dim O0 As TetrominoT
Dim Oobr As TetrominoRotationsT

Dim I0 As TetrominoT, I90 As TetrominoT, I270 As TetrominoT, I360 As TetrominoT
Dim Iobr As TetrominoRotationsT

Dim T0 As TetrominoT, T90 As TetrominoT, T270 As TetrominoT, T360 As TetrominoT
Dim Tobr As TetrominoRotationsT

Dim J0 As TetrominoT, J90 As TetrominoT, J270 As TetrominoT, J360 As TetrominoT
Dim Jobr As TetrominoRotationsT

Dim L0 As TetrominoT, L90 As TetrominoT, L270 As TetrominoT, L360 As TetrominoT
Dim Lobr As TetrominoRotationsT

Dim S0 As TetrominoT, S90 As TetrominoT, S270 As TetrominoT, S360 As TetrominoT
Dim Sobr As TetrominoRotationsT

Dim Z0 As TetrominoT, Z90 As TetrominoT, Z270 As TetrominoT, Z360 As TetrominoT
Dim Zobr As TetrominoRotationsT

Public Tetrominoes(1 To NUMBER_OF_TETROMINOES) As TetrominoRotationsT

Public Tetrion(1 To HEIGHT_TETRION, 1 To WIDTH_TETRION) As Long

Sub initTetrominoes()
    With O0
        .matrix(1, 1) = 1
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 255, 0)
    End With
        
    With Oobr
        .Rotations(1) = O0
        .Rotations(2) = O0
        .Rotations(3) = O0
        .Rotations(4) = O0
    End With

    With I0
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(4, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 255)
    End With

    With I90
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .matrix(3, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 255)
    End With

    With I270
        .matrix(1, 3) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .matrix(4, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 255)
    End With

    With I360
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(2, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 255)
    End With

    With Iobr
        .Rotations(1) = I0
        .Rotations(2) = I90
        .Rotations(3) = I270
        .Rotations(4) = I360
    End With

    With T0
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 255)
    End With

    With T90
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(2, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 255)
    End With

    With T270
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 255)
    End With

    With T360
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 255)
    End With

    With Tobr
        .Rotations(1) = T0
        .Rotations(2) = T90
        .Rotations(3) = T270
        .Rotations(4) = T360
    End With

    With J0
        .matrix(1, 3) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .matrix(4, 3) = 1
        .matrix(4, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 0, 255)
    End With

    With J90
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(2, 4) = 1
        .matrix(3, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 0, 255)
    End With

    With J270
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(4, 2) = 1
        .matrix(1, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 0, 255)
    End With

    With J360
        .matrix(2, 1) = 1
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .matrix(3, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 0, 255)
    End With

    With Jobr
        .Rotations(1) = J0
        .Rotations(2) = J90
        .Rotations(3) = J270
        .Rotations(4) = J360
    End With

    With L0
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(4, 2) = 1
        .matrix(4, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 129, 0)
    End With

    With L90
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .matrix(3, 4) = 1
        .matrix(2, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 129, 0)
    End With

    With L270
        .matrix(1, 2) = 1
        .matrix(1, 3) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .matrix(4, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 129, 0)
    End With

    With L360
        .matrix(3, 1) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(2, 4) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 129, 0)
    End With

    With Lobr
        .Rotations(1) = L0
        .Rotations(2) = L90
        .Rotations(3) = L270
        .Rotations(4) = L360
    End With

    With S0
        .matrix(1, 2) = 1
        .matrix(1, 3) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 0)
    End With

    With S90
        .matrix(1, 1) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 0)
    End With

    With S270
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 0)
    End With

    With S360
        .matrix(1, 1) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .Color = RGB(0, 255, 0)
    End With

    With Sobr
        .Rotations(1) = S0
        .Rotations(2) = S90
        .Rotations(3) = S270
        .Rotations(4) = S360
    End With

    With Z0
        .matrix(1, 1) = 1
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 0)
    End With

    With Z90
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 1) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 0)
    End With

    With Z270
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 0)
    End With

    With Z360
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 1) = 1
        .width = 4
        .height = 4
        .Color = RGB(255, 0, 0)
    End With

    With Zobr
        .Rotations(1) = Z0
        .Rotations(2) = Z90
        .Rotations(3) = Z270
        .Rotations(4) = Z360
    End With

    Tetrominoes(1) = Oobr
    Tetrominoes(2) = Iobr
    Tetrominoes(3) = Tobr
    Tetrominoes(4) = Jobr
    Tetrominoes(5) = Lobr
    Tetrominoes(6) = Sobr
    Tetrominoes(7) = Zobr
End Sub



