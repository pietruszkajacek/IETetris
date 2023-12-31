Attribute VB_Name = "Timer"
Option Explicit

#If Win64 Then
  Public Declare PtrSafe Function SetTimer Lib "User32" ( _
      ByVal hwnd As LongLong, _
      ByVal nIDEvent As LongLong, _
      ByVal uElapse As LongLong, _
      ByVal lpTimerFunc As LongLong) As LongLong
  Public Declare PtrSafe Function KillTimer Lib "User32" ( _
      ByVal hwnd As LongLong, _
      ByVal nIDEvent As LongLong) As LongLong
  Public TimerID As LongLong
#Else
  Public Declare PtrSafe Function SetTimer Lib "User32" ( _
      ByVal hwnd As Long, _
      ByVal nIDEvent As Long, _
      ByVal uElapse As Long, _
      ByVal lpTimerFunc As Long) As Long
  Public Declare PtrSafe Function KillTimer Lib "User32" ( _
      ByVal hwnd As Long, _
      ByVal nIDEvent As Long) As Long
  Public TimerID As Long
#End If

Sub StartTimer()
  If gameStarted = True Then
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
    TimerID = SetTimer(0, 0, speed, AddressOf TimerEvent)
  End If
End Sub

Sub TimerEvent()
  On Error Resume Next
  If gameStarted = True Then Main.mainLoop
  'Exit Sub
End Sub

Sub StopTimer()
  KillTimer 0, TimerID
  TimerID = 0
End Sub

