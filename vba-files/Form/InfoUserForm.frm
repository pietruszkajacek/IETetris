VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoUserForm 
   Caption         =   "Informacje o projekcie"
   ClientHeight    =   4032
   ClientLeft      =   2112
   ClientTop       =   2460
   ClientWidth     =   9276.001
   OleObjectBlob   =   "InfoUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    'HideTitleBar Me
    SystemButtonSettings Me, False

    With InfoUserForm
        .Left = 271
        .Top = 162
    End With
End Sub
