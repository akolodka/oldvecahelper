VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigManagerUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 09.03.2023 12:56
' Автор: akolodka
' Назначение: Обработка конфигурации
' ----------------------------------------------------------------
Option Explicit

Private Type TLocalVariables
    
    pathUpdateFolder As String
    isCheckAuto As Boolean
    isInstallAuto As Boolean
    
End Type

Private this As TLocalVariables, _
    cs As New cConfigSource, _
    fso As New FileSystemObject, _
    keys As New cConfigKeysUpdater
       
Private Sub Class_Initialize()

    cs.Initialize _
        sourceType:=configData, _
        nameKey:=keys.Header
    
    this.pathUpdateFolder = cs.FindValue(keys.pathUpdateFolder)
    this.isCheckAuto = cs.FindValue(keys.isCheckAuto, boolValue)
    this.isInstallAuto = cs.FindValue(keys.isInstallAuto, boolValue)

End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Public Property Get pathUpdateFolder() As String
    pathUpdateFolder = this.pathUpdateFolder
End Property
Friend Property Let pathUpdateFolder(ByVal NewValue As String)
    
    this.pathUpdateFolder = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        this.pathUpdateFolder = NewValue
    
End Property
' ----------------------------------------------------------------
Public Property Get isCheckAuto() As Boolean
    isCheckAuto = this.isCheckAuto
End Property
Friend Property Let isCheckAuto(ByVal NewValue As Boolean)
    this.isCheckAuto = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get isInstallAuto() As Boolean
    isInstallAuto = this.isInstallAuto
End Property
Friend Property Let isInstallAuto(ByVal NewValue As Boolean)
    this.isInstallAuto = NewValue
End Property
' ----------------------------------------------------------------

' ----------------------------------------------------------------
' Назначение: Выбор каталога
'    параметр initialSomniumPath:
' Дата: 09.03.2023 09:14
' ----------------------------------------------------------------
Friend Function ChooseFolderPath( _
    Optional initialSomniumPath As Boolean = False _
    ) As Boolean
    
    ChooseFolderPath = False
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    fd.Filters.Clear
    fd.InitialView = msoFileDialogViewDetails

    fd.Title = "Выбор каталога загрузки обновлений."
    fd.AllowMultiSelect = False

    fd.initialFileName = Base.desktopPath

    If initialSomniumPath Then _
        fd.initialFileName = Base.serverInitialPath

    If fd.Show = False Then _
        Set fd = Nothing: _
        Exit Function

    this.pathUpdateFolder = fd.SelectedItems(1) 'полный путь к файлу
    
    Set fd = Nothing
    ChooseFolderPath = True

End Function
' ----------------------------------------------------------------
Public Function IsCorrect() As Boolean
    
    IsCorrect = fso.FolderExists(this.pathUpdateFolder)
                
End Function
' ----------------------------------------------------------------
Friend Sub Save()

    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.pathUpdateFolder, this.pathUpdateFolder
    cs.PrepareToWrite keys.isCheckAuto, CStr(this.isCheckAuto)
    cs.PrepareToWrite keys.isInstallAuto, CStr(this.isInstallAuto)
    
    cs.WriteToFile
    
End Sub
