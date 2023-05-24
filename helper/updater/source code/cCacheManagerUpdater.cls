VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheManagerUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 25.03.2023 09:28
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Private Type TLocalVariables
    
    dateOfCheck As String
    IsUpdateAvailable As Boolean
    
End Type
Private this As TLocalVariables, _
    cs As New cConfigSource, _
    fso As New FileSystemObject, _
    keys As New cCacheKeysUpdater

Private Sub Class_Initialize()
    
    cs.Initialize _
        sourceType:=cacheData, _
        nameKey:=keys.Header
    
    this.dateOfCheck = cs.FindValue(keys.dateOfCheck)
    this.IsUpdateAvailable = cs.FindValue(keys.IsUpdateAvailable, boolValue)
    
End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Friend Sub Save()
    
    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.dateOfCheck, this.dateOfCheck
    cs.PrepareToWrite keys.IsUpdateAvailable, CStr(this.IsUpdateAvailable)
    
    cs.WriteToFile
    
End Sub
' ----------------------------------------------------------------
Public Property Get dateOfCheck() As String
    dateOfCheck = this.dateOfCheck
End Property
Public Property Let dateOfCheck(ByVal NewValue As String)
    this.dateOfCheck = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get IsUpdateAvailable() As Boolean
    IsUpdateAvailable = this.IsUpdateAvailable
End Property
Public Property Let IsUpdateAvailable(ByVal NewValue As Boolean)
    this.IsUpdateAvailable = NewValue
End Property


