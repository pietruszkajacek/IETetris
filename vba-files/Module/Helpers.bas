Attribute VB_Name = "Helpers"

Sub DisplayMonitorInfo()
    Dim w As LongLong, h As LongLong
    w = GetSystemMetrics(0) ' width in points
    h = GetSystemMetrics(1) ' height in points
    MsgBox Format(w, "#,##0") & " x " & Format(h, "#,##0"), _
    vbInformation, "Monitor Size (width x height)"
End Sub

