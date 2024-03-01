Attribute VB_Name = "Program"
Option Explicit

Private fso As New FileSystemObject

Private cfg As New cConfigManagerContracts
Private cs As New cConfigSource

Private keys As New cConfigKeysContracts

Public Function ConfigFilePath() As String
    
    cs.Initialize _
        sourceType:=configData, _
        nameKey:=keys.Header
        
    ConfigFilePath = cs.filePath
    
End Function
Public Sub OpenConfig()
    
    Dim path As String
    path = ConfigFilePath

    If Not fso.FileExists(path) Then _
        cfg.Save

    Shell "explorer.exe " & ConfigFilePath
End Sub
