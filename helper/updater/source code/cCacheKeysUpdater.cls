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

'Const XL_KEY = "modulesExcel"
'Const WD_KEY = "modulesWord"
'Const OTHER_KEY = "modulesOther"

Const READY_UPD_KEY = "modulesUpdateAvailable"

Friend Property Get Header() As String
    Header = CACHE_HEADER
End Property

Friend Property Get dateOfCheck() As String
    dateOfCheck = CHECK_DATE_KEY
End Property

Friend Property Get isUpdateAvailable() As String
    isUpdateAvailable = STATUS_KEY
End Property

'Friend Property Get modulesExcel() As String
'    modulesExcel = XL_KEY
'End Property
'
'Friend Property Get modulesWord() As String
'    modulesWord = WD_KEY
'End Property
'
'Friend Property Get modulesOther() As String
'    modulesOther = OTHER_KEY
'End Property

Friend Property Get modulesUpdateAvailable() As String
    modulesUpdateAvailable = READY_UPD_KEY
End Property
