VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashScreenUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6276
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9840.001
   OleObjectBlob   =   "SplashScreenUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashScreenUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
                       
Private Declare PtrSafe Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
 
Private Declare PtrSafe Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
 
Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As Long) As Long
 
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
 
'Constants for title bar
Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   'Controls if the window has an icon
 
'Constants for transparency
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1                  'Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     'Only needed if you want to fade the entire userform
 
Private Sub UserForm_Activate()
    HideTitleBarAndBorder Me 'hide the titlebar and border
    MakeUserFormTransparent Me 'make certain color transparent
End Sub
 
Sub MakeUserFormTransparent(frm As Object, Optional Color As Variant)
    'set transparencies on userform
    Dim formhandle As Long
    Dim bytOpacity As Byte
 
    formhandle = FindWindow(vbNullString, Me.Caption)
    If IsMissing(Color) Then Color = vbWhite 'default to vbwhite
    bytOpacity = 100 ' variable keeping opacity setting
 
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED

    'The following line makes only a certain color transparent so the
    ' background of the form and any object whose BackColor you've set to match
    ' vbColor (default vbWhite) will be transparent.
    Me.BackColor = Color
    SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Sub
 
Sub HideTitleBarAndBorder(frm As Object)
    'Hide title bar and border around userform
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    'Build window and set window until you remove the caption, title bar and frame around the window
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub
