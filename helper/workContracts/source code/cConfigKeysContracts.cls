VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigKeysContracts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cConfigKeysDatabase (Class Module)
' Назначение: ключи доступа к конфигурации библиотеки
' Автор: akolodka
' Date: 26.02.2023 11:14
' ------------------------------------------------------

Option Explicit

Const CONFIG_HEADER = "workContracts"

Const SOURCE_Folder_KEY = "sourceFolder"
Const DESTINATION_Folder_KEY = "destinationFolder"
Const FILTER_CRITERIA_KEY = "filterCriteria"


Friend Property Get Header() As String
    Header = CONFIG_HEADER
End Property

Friend Property Get sourceFolder() As String
    sourceFolder = SOURCE_Folder_KEY
End Property

Friend Property Get destinationFolder() As String
    destinationFolder = DESTINATION_Folder_KEY
End Property

Friend Property Get filterCriteria() As String
    filterCriteria = FILTER_CRITERIA_KEY
End Property
