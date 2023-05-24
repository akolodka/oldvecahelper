Attribute VB_Name = "Core"
Option Explicit

Private m_config As cConfigManagerNumName, _
        m_numName As cNumNameManager, _
        m_protocol As cNumNameProtocol, _
        m_cache As cCacheManagerNumName

Public Function Config() As cConfigManagerNumName
    
    If m_config Is Nothing Then _
        Set m_config = New cConfigManagerNumName
        
    Set Config = m_config
    
End Function

Public Function Cache() As cCacheManagerNumName
    
    If m_cache Is Nothing Then _
        Set m_cache = New cCacheManagerNumName
        
    Set Cache = m_cache
    
End Function

Public Function NumName() As cNumNameManager
    
    If m_numName Is Nothing Then _
        Set m_numName = New cNumNameManager
    
    Set NumName = m_numName
    
End Function

Public Function Protocol() As cNumNameProtocol
    
    If m_protocol Is Nothing Then _
        Set m_protocol = New cNumNameProtocol
    
    Set Protocol = m_protocol
    
End Function

Public Sub ClearSingletone()
    
    Set m_config = Nothing
    Set m_numName = Nothing
    Set m_protocol = Nothing
    Set m_cache = Nothing
    
End Sub
