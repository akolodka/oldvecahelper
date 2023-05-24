Attribute VB_Name = "Core"
Option Explicit

Private Type TSingleTones
    
    configManger As cConfigManagerDatabase
    dbManager As cDatabaseManager
    cacheManager As cCacheManagerDatabase
    
End Type

Private this As TSingleTones

Public Function Config() As cConfigManagerDatabase
    
    If this.configManger Is Nothing Then _
        Set this.configManger = New cConfigManagerDatabase
        
    Set Config = this.configManger
    
End Function
Public Function Cache() As cCacheManagerDatabase
    
    If this.cacheManager Is Nothing Then _
        Set this.cacheManager = New cCacheManagerDatabase
        
    Set Cache = this.cacheManager
    
End Function
Public Function DataBase() As cDatabaseManager

    If this.dbManager Is Nothing Then _
        Set this.dbManager = New cDatabaseManager

    Set DataBase = this.dbManager

End Function

Public Sub ClearSingletone()
    
    Set this.configManger = Nothing
    Set this.dbManager = Nothing
    Set this.cacheManager = Nothing

End Sub
