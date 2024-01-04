VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5244
   OleObjectBlob   =   "SplashUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    HideTitleBar Me
End Sub

Private Sub UserForm_Activate()
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label1.Caption = "Loading data..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label1.Caption = "Creating forms..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    SplashUserForm.Label1.Caption = "Opening..."
    SplashUserForm.Repaint
    Application.Wait (Now + TimeValue("00:00:01"))
    Unload SplashUserForm
End Sub
