Attribute VB_Name = "BaseInterfaces"
Option Explicit

Public Enum Colors

    black = &H80000008
    white = &HFFFFFF
    
    grey = &H8000000C
    greyLight = &HC0C0C0
    
    blue = &HFF0000
    blueDark = &HC00000
    
    green = &H8000&
    greenPastel = &HC0FFC0
    greenLight = &HC000&
    
    red = &H80&
    redLips = &HFF&
    
    yellowPastel = &HC0FFFF
    yellowGold = &HC0C0&
    
    turquoise = &H808000 'בטנ‏חמגûי
    orangePastel = &HC0E0FF
    
    
End Enum

Private m_addin As cAddin, _
        m_base As cBase, _
        m_properties As cProperties, _
        m_userMenu As cUserMenu, _
        m_explorer As cExplorer, _
        m_handler As cHandler
        
Public Function Addin() As cAddin
    
    If m_addin Is Nothing Then _
        Set m_addin = New cAddin
        
    Set Addin = m_addin
    
End Function
Public Function Base() As cBase
    
    If m_base Is Nothing Then _
        Set m_base = New cBase
        
    Set Base = m_base
    
End Function
Public Function Explorer() As cExplorer
    
    If m_explorer Is Nothing Then _
        Set m_explorer = New cExplorer
        
    Set Explorer = m_explorer
    
End Function
Public Function Properties() As cProperties

    If m_properties Is Nothing Then _
        Set m_properties = New cProperties

    Set Properties = m_properties

End Function
Public Function UMenu() As cUserMenu

    If m_userMenu Is Nothing Then _
        Set m_userMenu = New cUserMenu

    Set UMenu = m_userMenu
    
End Function
Public Function Handler() As cHandler
    
    If m_handler Is Nothing Then _
        Set m_handler = New cHandler
        
    Set Handler = m_handler
    
End Function
