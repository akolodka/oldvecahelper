VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheKeysDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 25.03.2023 09:24
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Const CACHE_HEADER = "dataBase"

Const KEY_ORGANISATIONS = "organisations"
Const KEY_INSTRUMENTS = "instruments"
Const KEY_ETALONS = "etalons"
Const KEY_PERSON = "persons"

Friend Property Get Header() As String
    Header = CACHE_HEADER
End Property

Public Property Get organisationLastSearch() As String
    organisationLastSearch = KEY_ORGANISATIONS
End Property

Public Property Get instrumentLastSearch() As String
    instrumentLastSearch = KEY_INSTRUMENTS
End Property

Public Property Get etalonLastSearch() As String
    etalonLastSearch = KEY_ETALONS
End Property

Public Property Get personLastSearch() As String
    personLastSearch = KEY_PERSON
End Property
