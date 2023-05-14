VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
#End If

Const SW_RESTORE As Byte = 9
Const SHIFT_SHIFT = 1
Const SHIFT_CTRL = 2

Private fso As New FileSystemObject
Private oShell As New Shell32.Shell
Private oWindow As Object

Friend Sub OpenActiveWbPath( _
    Optional ByVal KeyCode As MSForms.ReturnInteger, _
    Optional ByVal Shift As Integer _
    )
    
    Dim targetPath As String
    targetPath = ActiveWorkbook.path
    
    If Shift = SHIFT_SHIFT Or Shift = SHIFT_CTRL Then _
        targetPath = Base.configDataPath
    
    Explorer.OpenFolder _
        path:=targetPath, _
        isFocusOnWindow:=True
    
    If UMenu.isLoaded Then _
        UMenu.Unload
        
End Sub

Friend Sub OpenFolder( _
    path As String, _
    Optional isFocusOnWindow As Boolean = False _
    )
    
    If Not IsCorrect(path) Then _
        Exit Sub
    
    Dim targetWindowName As String
    targetWindowName = WindowName(path)
    
    Dim isFinded As Boolean
    isFinded = IsWindowExists(targetWindowName)
    
    If isFinded Then
        
        If Not isFocusOnWindow Then _
            Exit Sub

        ActivateWindow
        
    Else 'среди открытых окон не найдено нужного окна
        
        If isFocusOnWindow Then _
            ActiveWorkbook.FollowHyperlink path: _
            Exit Sub
        
        If IsDesktopPath(path) Then _
            Exit Sub
        
        ShellOpen path
    
    End If

End Sub
    Private Function IsCorrect( _
        path As String _
        ) As Boolean
        
        IsCorrect = path <> "\"
        
    End Function
    Private Function WindowName( _
        path As String _
        ) As String
        
        Dim result As String
        result = fso.GetBaseName(path)
        
        If result = "Desktop" Then _
            result = "Рабочий стол"
        
        WindowName = result
        
    End Function
    Private Function IsWindowExists( _
        name As String _
        ) As Boolean
        
        IsWindowExists = False
        
        For Each oWindow In oShell.Windows 'пройтись по открытым окнам
            
            If oWindow.locatioNname = name Then
                
                IsWindowExists = True
                Exit Function
                
            End If
            
        Next

    End Function
    Private Sub ActivateWindow()
    
        oWindow.visible = False
        oWindow.visible = True
        
        If CBool(IsIconic(oWindow.hwnd)) Then _
            ShowWindow oWindow.hwnd, SW_RESTORE
        
    End Sub
    Private Function IsDesktopPath( _
        path As String _
        ) As Boolean
        
        IsDesktopPath = path = Base.desktopPath
        
    End Function
    Private Sub ShellOpen( _
        path As String _
        )
        
        Shell "explorer.exe " & path, vbNormalNoFocus
            
    End Sub
