VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ten_skoroszyt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    ActiveWindow.Visible = False
    SplashUserForm.show
    'Application.WindowState = xlMaximized
    'Application.DisplayFullScreen = True
    Windows(ThisWorkbook.Name).Visible = True
    Application.ScreenUpdating = True

    MainUserForm.show
End Sub

