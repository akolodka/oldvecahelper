VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheManagerNumName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 09.03.2023 17:19
' Автор: akolodka
' Назначение: Управление кешем
' ----------------------------------------------------------------
Option Explicit

Private cs As New cConfigSource
Private fso As New FileSystemObject
Private keys As New cCacheKeysNumName

Private m_cachedNumber As String, _
        m_empOther As String
        
Private m_useEmpSecond As String, _
        m_useEmpThird As String, _
        m_useEmpOther As String, _
        m_saveAsCopy As String

Private Sub Class_Initialize()

    cs.Initialize _
        sourceType:=cacheData, _
        nameKey:=keys.header
    
    m_cachedNumber = cs.FindValue(keys.cachedNumber)
    
    m_useEmpSecond = cs.FindValue(keys.useEmployeeSecond, boolValue)
    m_useEmpThird = cs.FindValue(keys.useEmployeeThird, boolValue)
    m_useEmpOther = cs.FindValue(keys.useEmployeeOther, boolValue)
    
    m_saveAsCopy = cs.FindValue(keys.saveAsCopy, boolValue)
    m_empOther = cs.FindValue(keys.employeeOther)
    
End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Friend Sub Save()

    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи

    cs.PrepareToWrite keys.cachedNumber, m_cachedNumber
    cs.PrepareToWrite keys.useEmployeeSecond, CStr(m_useEmpSecond)
    cs.PrepareToWrite keys.useEmployeeThird, CStr(m_useEmpThird)
    cs.PrepareToWrite keys.useEmployeeOther, CStr(m_useEmpOther)
    cs.PrepareToWrite keys.employeeOther, m_empOther
        
    cs.PrepareToWrite keys.saveAsCopy, CStr(m_saveAsCopy)
    cs.WriteToFile

End Sub
' ----------------------------------------------------------------
Friend Property Get cachedNumber( _
    ) As String
    
    cachedNumber = m_cachedNumber
        
End Property
Friend Property Let cachedNumber( _
    ByVal NewValue As String _
    )

    m_cachedNumber = Base.defaultValue

    If NewValue <> vbNullString Then _
        m_cachedNumber = NewValue
        
End Property
' ----------------------------------------------------------------
Public Property Get useEmpSecond() As Boolean
    useEmpSecond = m_useEmpSecond
End Property
Public Property Let useEmpSecond(ByVal NewValue As Boolean)
    m_useEmpSecond = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get useEmpThird() As Boolean
    useEmpThird = m_useEmpThird
End Property
Public Property Let useEmpThird(ByVal NewValue As Boolean)
    m_useEmpThird = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get useEmpOther() As Boolean
    useEmpOther = m_useEmpOther
End Property
Public Property Let useEmpOther(ByVal NewValue As Boolean)
    m_useEmpOther = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get empOther() As String
    empOther = m_empOther
End Property
Public Property Let empOther(ByVal NewValue As String)
    m_empOther = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get saveAsCopy() As Boolean
    saveAsCopy = m_saveAsCopy
End Property
Public Property Let saveAsCopy(ByVal NewValue As Boolean)
    m_saveAsCopy = NewValue
End Property


