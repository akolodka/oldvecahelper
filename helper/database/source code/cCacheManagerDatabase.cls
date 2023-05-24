VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheManagerDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 25.03.2023 09:28
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Private Type TCacheFields
    
    organisationLastSearch As String
    instrumentLastSearch As String
    etalonLastSearch As String
    personLastSearch As String
    
End Type
Private this As TCacheFields, _
    cs As New cConfigSource, _
    fso As New FileSystemObject, _
    keys As New cCacheKeysDatabase

Private Sub Class_Initialize()
    
    cs.Initialize _
        sourceType:=cacheData, _
        nameKey:=keys.Header
    
    this.organisationLastSearch = cs.FindValue(keys.organisationLastSearch)
    this.instrumentLastSearch = cs.FindValue(keys.instrumentLastSearch)
    this.etalonLastSearch = cs.FindValue(keys.etalonLastSearch)
    this.personLastSearch = cs.FindValue(keys.personLastSearch)
    
End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Friend Sub Save()
    
    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.organisationLastSearch, this.organisationLastSearch
    cs.PrepareToWrite keys.instrumentLastSearch, this.instrumentLastSearch
    cs.PrepareToWrite keys.etalonLastSearch, this.etalonLastSearch
    cs.PrepareToWrite keys.personLastSearch, this.personLastSearch
    
    cs.WriteToFile
    
End Sub
' ----------------------------------------------------------------
Public Property Get organisationLastSearch() As String
    organisationLastSearch = this.organisationLastSearch
End Property
Public Property Let organisationLastSearch(ByVal NewValue As String)
    this.organisationLastSearch = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get instrumentLastSearch() As String
    instrumentLastSearch = this.instrumentLastSearch
End Property
Public Property Let instrumentLastSearch(ByVal NewValue As String)
    this.instrumentLastSearch = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get etalonLastSearch() As String
    etalonLastSearch = this.etalonLastSearch
End Property
Public Property Let etalonLastSearch(ByVal NewValue As String)
    this.etalonLastSearch = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get personLastSearch() As String
    personLastSearch = this.personLastSearch
End Property
Public Property Let personLastSearch(ByVal NewValue As String)
    this.personLastSearch = NewValue
End Property




