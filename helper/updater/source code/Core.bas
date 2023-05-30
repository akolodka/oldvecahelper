Attribute VB_Name = "Core"
Option Explicit

Private Type TSingleTones
    
    configManger As cConfigManagerUpdater
    cacheManager As cCacheManagerUpdater
    updateManager As cUpdateManager
    
End Type

Private this As TSingleTones

Public Function Config() As cConfigManagerUpdater
    
    If this.configManger Is Nothing Then _
        Set this.configManger = New cConfigManagerUpdater
        
    Set Config = this.configManger
    
End Function
Public Function Cache() As cCacheManagerUpdater
    
    If this.cacheManager Is Nothing Then _
        Set this.cacheManager = New cCacheManagerUpdater
        
    Set Cache = this.cacheManager
    
End Function

Public Function Update() As cUpdateManager

    If this.updateManager Is Nothing Then _
        Set this.updateManager = New cUpdateManager

    Set Update = this.updateManager

End Function

Public Sub ClearSingletone()
    
    Set this.configManger = Nothing
    Set this.cacheManager = Nothing
    Set this.updateManager = Nothing

End Sub
