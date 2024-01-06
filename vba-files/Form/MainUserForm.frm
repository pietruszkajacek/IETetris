VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainUserForm 
   Caption         =   "IE Tetris © 2023-2024 Jacek Pietruszka"
   ClientHeight    =   5892
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8400.001
   OleObjectBlob   =   "MainUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CreditsButton_Click()
   InfoUserForm.show
End Sub

Private Sub EndGameButton_Click()
    Main.stopGame
End Sub

Private Sub StartGameButton_Click()
    'Me.Hide
    'InfoUserForm.show
    AppActivate Application.Caption
    'InfoUserForm.Enabled = False
    
    Main.startGame
End Sub

Private Sub UserForm_Initialize()
    'HideTitleBar Me
    SystemButtonSettings Me, False
End Sub
