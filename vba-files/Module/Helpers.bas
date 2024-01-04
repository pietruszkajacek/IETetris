Attribute VB_Name = "Helpers"
    
Option Explicit
Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000

Public Declare PtrSafe Function GetWindowLong _
                       Lib "user32" Alias "GetWindowLongA" ( _
                       ByVal hwnd As Long, _
                       ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong _
                       Lib "user32" Alias "SetWindowLongA" ( _
                       ByVal hwnd As Long, _
                       ByVal nIndex As Long, _
                       ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function DrawMenuBar _
                       Lib "user32" ( _
                       ByVal hwnd As Long) As Long
Public Declare PtrSafe Function FindWindowA _
                       Lib "user32" (ByVal lpClassName As String, _
                       ByVal lpWindowName As String) As Long

Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub

Sub DisplayMonitorInfo()
    Dim w As LongLong, h As LongLong
    w = GetSystemMetrics(0) ' width in points
    h = GetSystemMetrics(1) ' height in points
    MsgBox Format(w, "#,##0") & " x " & Format(h, "#,##0"), _
    vbInformation, "Monitor Size (width x height)"
End Sub

Public Sub SystemButtonSettings(frm As Object, show As Boolean)
    Dim windowStyle As Long
    Dim windowHandle As Long

    windowHandle = FindWindowA(vbNullString, frm.Caption)
    windowStyle = GetWindowLong(windowHandle, GWL_STYLE)

    If show = False Then
        SetWindowLong windowHandle, GWL_STYLE, (windowStyle And Not WS_SYSMENU)
    Else
        SetWindowLong windowHandle, GWL_STYLE, (windowStyle + WS_SYSMENU)
    End If

    DrawMenuBar (windowHandle)
End Sub
