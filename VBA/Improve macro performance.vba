Sub PerformanceSettings(State As Boolean)
    ' Call this function with State=True whenever you are about to perform
    ' a memory/performance-intensive task. When the task is finished or an error
    ' is raised while it's executing call this function again with State=False.
    ' Requires public variable <CalculationSetting> to be declared beforehand:
    ' Public CalculationSetting As Integer
    
    Application.DisplayStatusBar = True ' always show statusbar
    Application.StatusBar = False ' give control over statusbar to excel
    Application.ScreenUpdating = Not State ' disable/enable screen updates e.g. for long execution times
    Application.DisplayAlerts = Not State ' disable/enable alerts to allow uninterrupted execution
    CalculationSetting = Application.Calculation ' store current setting in public variable
    If State Then
        Application.Calculation = xlCalculationManual
    Else
        If Application.Calculation <> CalculationSetting Then
            Application.Calculation = CalculationSetting
        End If
    End If
End Sub