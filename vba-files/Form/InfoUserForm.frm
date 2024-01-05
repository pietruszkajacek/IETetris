VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoUserForm 
   Caption         =   "Credits"
   ClientHeight    =   4272
   ClientLeft      =   2112
   ClientTop       =   2460
   ClientWidth     =   9540.001
   OleObjectBlob   =   "InfoUserForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "InfoUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    'HideTitleBar Me
    SystemButtonSettings Me, False

    With InfoUserForm
        .Left = 271
        .Top = 162
    End With
End Sub
