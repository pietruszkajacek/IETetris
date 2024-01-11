VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BestPlayersUserForm 
   Caption         =   "Podaj swoje imiê / pseudonim do listy najlepszych"
   ClientHeight    =   3708
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4836
   OleObjectBlob   =   "BestPlayersUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BestPlayersUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim saved As Boolean
Dim FRow As Range

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub SaveButton_Click()
    If NickTextBox.Text <> "" Then
        Worksheets("Najlepsi").Cells(FRow.Row, 2) = NickTextBox.Text
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
        saved = True
        Unload Me
    Else
        MsgBox "Pole nie mo¿e byæ puste...", vbOKOnly Or vbExclamation, "Uzupe³nij pole..."
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim maxID As LongLong, playedId As LongLong
    Dim lastRow As LongLong
    Dim idsRang As Range

    With Worksheets("Najlepsi")
        maxID = WorksheetFunction.Max(.Range("A2", .Range("A2").End(xlDown)))

        Set idsRang = .Range("A1", .Range("A1").End(xlDown))
        
        lastRow = idsRang.Rows.Count

        .Cells(lastRow + 1, 1) = maxID + 1
        .Cells(lastRow + 1, 3) = level
        .Cells(lastRow + 1, 4) = completeLines
        .Cells(lastRow + 1, 5) = statTetrominoes(1)
        .Cells(lastRow + 1, 6) = statTetrominoes(2)
        .Cells(lastRow + 1, 7) = statTetrominoes(3)
        .Cells(lastRow + 1, 8) = statTetrominoes(4)
        .Cells(lastRow + 1, 9) = statTetrominoes(5)
        .Cells(lastRow + 1, 10) = statTetrominoes(6)
        .Cells(lastRow + 1, 11) = statTetrominoes(7)
    
        'Set sortRang = .Range("A1", .Range("D1").End(xlDown))

        .Range("A1", .Range("K1").End(xlDown)).Sort Key1:=.Range("C1"), Key2:=.Range("D1"), Header:=xlYes, _
            Order1:=xlDescending, Order2:=xlDescending

        Set idsRang = .Range("A1", .Range("A1").End(xlDown))
        Set FRow = idsRang.Find(maxID + 1, LookIn:=xlValues, LookAt:=xlWhole)
        posRangeLabel.Caption = Str(FRow.Row - 1)
    End With
End Sub

Private Sub UserForm_Terminate()
    If Not saved Then
        With Worksheets("Najlepsi")
            .Rows(FRow.Row).Delete
        End With
    End If
End Sub

