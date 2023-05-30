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
    isUpdateAvailable As Boolean
    
'    modulesExcel As String
'    modulesWord As String
'    modulesOther As String
    
    modulesUpdateAvailable As String
    
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
    this.isUpdateAvailable = cs.FindValue(keys.isUpdateAvailable, boolValue)
'
'    this.modulesExcel = cs.FindValue(keys.modulesExcel)
'    this.modulesWord = cs.FindValue(keys.modulesWord)
'    this.modulesOther = cs.FindValue(keys.modulesOther)
    
    this.modulesUpdateAvailable = cs.FindValue(keys.modulesUpdateAvailable)
    
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
    cs.PrepareToWrite keys.isUpdateAvailable, CStr(this.isUpdateAvailable)
    
'    cs.PrepareToWrite keys.modulesExcel, this.modulesExcel
'    cs.PrepareToWrite keys.modulesWord, this.modulesWord
'    cs.PrepareToWrite keys.modulesOther, this.modulesOther
    
    cs.PrepareToWrite keys.modulesUpdateAvailable, this.modulesUpdateAvailable
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
Public Property Get isUpdateAvailable() As Boolean
    isUpdateAvailable = this.isUpdateAvailable
End Property
Public Property Let isUpdateAvailable(ByVal NewValue As Boolean)
    this.isUpdateAvailable = NewValue
End Property
' ----------------------------------------------------------------
'Public Property Get modulesExcel() As String
'    modulesExcel = this.modulesExcel
'End Property
'Public Property Let modulesExcel(ByVal NewValue As String)
'    this.modulesExcel = NewValue
'End Property
'' ----------------------------------------------------------------
'Public Property Get modulesWord() As String
'    modulesWord = this.modulesWord
'End Property
'Public Property Let modulesWord(ByVal NewValue As String)
'    this.modulesWord = NewValue
'End Property
'' ----------------------------------------------------------------
'Public Property Get modulesOther() As String
'    modulesOther = this.modulesOther
'End Property
'Public Property Let modulesOther(ByVal NewValue As String)
'    this.modulesOther = NewValue
'End Property
' ----------------------------------------------------------------
Public Property Get modulesUpdateAvailable() As String
    modulesUpdateAvailable = this.modulesUpdateAvailable
End Property
Public Property Let modulesUpdateAvailable(ByVal NewValue As String)
    this.modulesUpdateAvailable = NewValue
End Property
