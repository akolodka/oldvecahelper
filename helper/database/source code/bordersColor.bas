Attribute VB_Name = "bordersColor"
Option Explicit

Const THEME_C = 1
Const TINT = -0.499984740745262

Sub Borders_Color()
    
    Notifications_OFF

        Dim myRange As Range
        Set myRange = Range(ActiveWorkbook.ActiveSheet.PageSetup.PrintArea)

        Dim rCell As Range
        For Each rCell In myRange
        
            With rCell
                
'                If .Value <> vbNullString Then
'
'                    With .Font
'                        .ThemeColor = xlThemeColorLight1
'                        .TintAndShade = 0.349986266670736
'                    End With
'                End If
                
                With .Borders(xlEdgeBottom)
                    If .LineStyle <> xlNone Then
                        .ThemeColor = THEME_C
                        .TintAndShade = TINT
                    End If
                End With
                
                With .Borders(xlEdgeTop)
                    If .LineStyle <> xlNone Then
                        .ThemeColor = THEME_C
                        .TintAndShade = TINT
                    End If
                End With
                
                With .Borders(xlEdgeLeft)
                    If .LineStyle <> xlNone Then
                        .ThemeColor = THEME_C
                        .TintAndShade = TINT
                    End If
                End With
                
                With .Borders(xlEdgeRight)
                    If .LineStyle <> xlNone Then
                        .ThemeColor = THEME_C
                        .TintAndShade = TINT
                    End If
                End With
                
            End With
    
        Next
    
    Notifications_ON
    
End Sub

    Private Sub Notifications_OFF()
        With Application
        
            .DisplayAlerts = False
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .StatusBar = False
            
        End With
    End Sub
    Private Sub Notifications_ON()
        With Application
            
            .DisplayAlerts = True
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            .StatusBar = True
            
        End With
    End Sub
