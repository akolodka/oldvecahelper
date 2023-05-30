VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigKeysUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 17.04.2023 21:41
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Const CONFIG_HEADER = "updater"

Const UPDATE_FOLDER_PATH_KEY = "updateFolderPath"
Const CHECK_AUTO_KEY = "isCheckAuto"
Const INSTALL_AUTO_KEY = "isInstallAuto"

Friend Property Get Header() As String
    Header = CONFIG_HEADER
End Property

Friend Property Get pathUpdateFolder() As String
    pathUpdateFolder = UPDATE_FOLDER_PATH_KEY
End Property

Friend Property Get isCheckAuto() As String
    isCheckAuto = CHECK_AUTO_KEY
End Property

Friend Property Get isInstallAuto() As String
    isInstallAuto = INSTALL_AUTO_KEY
End Property
