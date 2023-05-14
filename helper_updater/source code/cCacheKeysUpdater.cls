VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCacheKeysUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 17.04.2023 21:53
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Const CACHE_HEADER = "updater"

Const CHECK_DATE_KEY = "dateOfCheck"
Const STATUS_KEY = "isUpdateAvailable"


Friend Property Get Header() As String
    Header = CACHE_HEADER
End Property

Friend Property Get dateOfCheck() As String
    dateOfCheck = CHECK_DATE_KEY
End Property

Friend Property Get IsUpdateAvailable() As String
    IsUpdateAvailable = STATUS_KEY
End Property
