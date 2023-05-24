VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigKeysDatabase"
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

Const CONFIG_HEADER = "dataBase"

Const SOURCEDATA_PATH_KEY = "sourceDataPath"
Const SANDBOX_PATH_KEY = "sandboxPath"

Const ARCHIVE_LOCALPATH_KEY = "archiveLocalPath"
Const ARCHIVE_LOCALDIR_KEY = "archive"

Const ORGANISATIONS_DIR_KEY = "organisations"

Const INSTRUMENTS_DIR_KEY = "instruments"
'Const INSTRUMENTS_DIR_KEY = "instruments_TEST"

Const ETALONS_DIR_KEY = "etalons"
Const PERSONS_DIR_KEY = "persons"

Const TEMPLATES_BASE_DIR_KEY = "_templates base"

Const VERIFIER_KEY = "verifierKey"
Const EXECUTOR_KEY = "executorKey"

Friend Property Get Header() As String
    Header = CONFIG_HEADER
End Property

Friend Property Get sourceDataPath() As String
    sourceDataPath = SOURCEDATA_PATH_KEY
End Property

Friend Property Get organisationsDir() As String
    organisationsDir = ORGANISATIONS_DIR_KEY
End Property

Friend Property Get instrumentsDir() As String
    instrumentsDir = INSTRUMENTS_DIR_KEY
End Property

Friend Property Get etalonsDir() As String
    etalonsDir = ETALONS_DIR_KEY
End Property

Friend Property Get personsDir() As String
    personsDir = PERSONS_DIR_KEY
End Property

Friend Property Get sandboxPath() As String
    sandboxPath = SANDBOX_PATH_KEY
End Property

Friend Property Get archiveLocalPath() As String
    archiveLocalPath = ARCHIVE_LOCALPATH_KEY
End Property

Friend Property Get archiveLocalDir() As String
    archiveLocalDir = ARCHIVE_LOCALDIR_KEY
End Property

Friend Property Get verifierKey() As String
    verifierKey = VERIFIER_KEY
End Property

Friend Property Get executorKey() As String
    executorKey = EXECUTOR_KEY
End Property

Public Property Get templatesBaseDir() As String
    templatesBaseDir = TEMPLATES_BASE_DIR_KEY
End Property
