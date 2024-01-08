Attribute VB_Name = "Init"
Option Explicit

Type TetrominoT
    matrix(1 To 4, 1 To 4) As Byte
    width As Byte
    height As Byte
    color As Long
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

Dim I0 As TetrominoT, I90 As TetrominoT, I180 As TetrominoT, I270 As TetrominoT
Dim Iobr As TetrominoRotationsT

Dim T0 As TetrominoT, T90 As TetrominoT, T180 As TetrominoT, T270 As TetrominoT
Dim Tobr As TetrominoRotationsT

Dim J0 As TetrominoT, J90 As TetrominoT, J180 As TetrominoT, J270 As TetrominoT
Dim Jobr As TetrominoRotationsT

Dim L0 As TetrominoT, L90 As TetrominoT, L180 As TetrominoT, L270 As TetrominoT
Dim Lobr As TetrominoRotationsT

Dim S0 As TetrominoT, S90 As TetrominoT, S180 As TetrominoT, S270 As TetrominoT
Dim Sobr As TetrominoRotationsT

Dim Z0 As TetrominoT, Z90 As TetrominoT, Z180 As TetrominoT, Z270 As TetrominoT
Dim Zobr As TetrominoRotationsT

Public Tetrominoes(1 To NUMBER_OF_TETROMINOES) As TetrominoRotationsT

Public Tetrion(1 To HEIGHT_TETRION, 1 To WIDTH_TETRION) As Long

Sub initTetrominoes()
    With O0
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .width = 4
        .height = 4
        .color = RGB(255, 255, 0)
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
        .color = RGB(0, 255, 255)
    End With

    With I90
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .matrix(3, 4) = 1
        .width = 4
        .height = 4
        .color = RGB(0, 255, 255)
    End With

    With I180
        .matrix(1, 3) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .matrix(4, 3) = 1
        .width = 4
        .height = 4
        .color = RGB(0, 255, 255)
    End With

    With I270
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(2, 4) = 1
        .width = 4
        .height = 4
        .color = RGB(0, 255, 255)
    End With

    With Iobr
        .Rotations(1) = I0
        .Rotations(2) = I90
        .Rotations(3) = I180
        .Rotations(4) = I270
    End With

    With T0
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .color = RGB(255, 0, 255)
    End With

    With T90
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(2, 3) = 1
        .width = 4
        .height = 4
        .color = RGB(255, 0, 255)
    End With

    With T180
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 4
        .height = 4
        .color = RGB(255, 0, 255)
    End With

    With T270
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 4
        .height = 4
        .color = RGB(255, 0, 255)
    End With

    With Tobr
        .Rotations(1) = T0
        .Rotations(2) = T90
        .Rotations(3) = T180
        .Rotations(4) = T270
    End With

    With J0
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 0, 255)
    End With

    With J90
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 0, 255)
    End With

    With J180
        .matrix(1, 2) = 1
        .matrix(1, 3) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 0, 255)
    End With

    With J270
        .matrix(1, 1) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 0, 255)
    End With

    With Jobr
        .Rotations(1) = J0
        .Rotations(2) = J90
        .Rotations(3) = J180
        .Rotations(4) = J270
    End With

    With L0
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 129, 0)
    End With

    With L90
        .matrix(1, 3) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 129, 0)
    End With

    With L180
        .matrix(1, 1) = 1
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 129, 0)
    End With

    With L270
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 1) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 129, 0)
    End With

    With Lobr
        .Rotations(1) = L0
        .Rotations(2) = L90
        .Rotations(3) = L180
        .Rotations(4) = L270
    End With

    With S0
        .matrix(1, 2) = 1
        .matrix(1, 3) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 255, 0)
    End With

    With S90
        .matrix(1, 1) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 255, 0)
    End With

    With S180
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 1) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 255, 0)
    End With

    With S270
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(0, 255, 0)
    End With

    With Sobr
        .Rotations(1) = S0
        .Rotations(2) = S90
        .Rotations(3) = S180
        .Rotations(4) = S270
    End With

    With Z0
        .matrix(1, 1) = 1
        .matrix(1, 2) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 0, 0)
    End With

    With Z90
        .matrix(1, 2) = 1
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 1) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 0, 0)
    End With

    With Z180
        .matrix(2, 1) = 1
        .matrix(2, 2) = 1
        .matrix(3, 2) = 1
        .matrix(3, 3) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 0, 0)
    End With

    With Z270
        .matrix(1, 3) = 1
        .matrix(2, 2) = 1
        .matrix(2, 3) = 1
        .matrix(3, 2) = 1
        .width = 3
        .height = 3
        .color = RGB(255, 0, 0)
    End With

    With Zobr
        .Rotations(1) = Z0
        .Rotations(2) = Z90
        .Rotations(3) = Z180
        .Rotations(4) = Z270
    End With

    Tetrominoes(1) = Oobr
    Tetrominoes(2) = Iobr
    Tetrominoes(3) = Tobr
    Tetrominoes(4) = Jobr
    Tetrominoes(5) = Lobr
    Tetrominoes(6) = Sobr
    Tetrominoes(7) = Zobr
End Sub



