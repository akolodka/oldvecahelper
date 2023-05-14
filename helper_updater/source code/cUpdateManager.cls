VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 17.04.2023 22:23
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit
Private fso As New FileSystemObject
' ----------------------------------------------------------------
' Дата: 06.12.2021 14:01, 17.04.2023 22:53
' ----------------------------------------------------------------
Public Sub Check( _
    )
    If Not fso.FolderExists(Config.pathUpdateFolder) Then _
        Exit Sub
    
    If Cache.IsUpdateAvailable Then _
        Exit Sub

    Dim status As Boolean: status = IsUpdateAvailable(Config.pathUpdateFolder)
    ' ----------------------------------------------------------------
    Cache.IsUpdateAvailable = status
    Cache.Save

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 17.04.2023 22:53
    ' ----------------------------------------------------------------
    Private Function IsUpdateAvailable( _
        updateFolderPath As String _
        ) As Boolean
        
        IsUpdateAvailable = False
        
        If updateFolderPath = Base.defaultValue Then _
            Exit Function
        'todo: [-] isUpdateAvailable сделать через collection subFiles
        Dim fileName As String
        fileName = Dir(updateFolderPath & "*")
        
        Dim currFile As cHelperVBProject
        Do While fileName <> vbNullString
            
            Set currFile = New cHelperVBProject
            currFile.GetData updateFolderPath, fileName
            
            If currFile.IsUpdateAvailable Then _
                IsUpdateAvailable = True: _
                Exit Do
            
            fileName = Dir
        Loop
        
        Set currFile = Nothing
        
    End Function
    
' ----------------------------------------------------------------
' Дата: 17.04.2023 22:54
' ----------------------------------------------------------------
Public Sub Install( _
    Optional isReload As Boolean = False _
    )
    
    If Config.pathUpdateFolder = Base.defaultValue Then _
        Exit Sub
    'todo: [-] Install сделать через collection subFiles
    Dim fileName As String
    fileName = Dir(Config.pathUpdateFolder & "*")
    
    Dim currFile As cHelperVBProject
    Do While fileName <> vbNullString
        
        If fileName <> ThisWorkbook.name Then
        
            Set currFile = New cHelperVBProject
            currFile.GetData Config.pathUpdateFolder, fileName
            currFile.Update isReload
            
        End If
        
        fileName = Dir
    Loop
    
    Cache.IsUpdateAvailable = False
    Cache.Save
    
End Sub


