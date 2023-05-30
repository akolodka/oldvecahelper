VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHelperVBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const TYPE_MODULE_EXCEL = "Надстройка Microsoft Excel"
Const TYPE_MODULE_WORD = "Шаблон Microsoft Word с поддержкой макросов"

Const WORD_APPLICATION_NAME = "Word.Application"

Private Type TLocalVariables
    
    name As String
    type As String
    
    isInstalled As Boolean
    dateModuleInstalled As String
    dateModuleUpdate As String
    
    pathFileInstalled As String
    pathFileUpdate As String
    
End Type
        
Private fso As New FileSystemObject, _
        this As TLocalVariables

Private Sub Class_Initialize()

    this.name = Base.defaultValue
    this.type = Base.defaultValue
    
    this.isInstalled = False
    this.dateModuleInstalled = Base.defaultValue
    this.dateModuleUpdate = Base.defaultValue
    
    this.pathFileInstalled = Base.defaultValue
    this.pathFileUpdate = Base.defaultValue
    
End Sub

' ----------------------------------------------------------------
' Дата: 30.05.2023 20:48
' ----------------------------------------------------------------
Friend Sub Initialize( _
    updateFolderPath As String, _
    moduleName As String _
    )
    
    If updateFolderPath = Base.defaultValue Or _
        moduleName = Base.defaultValue Then _
        Exit Sub
        
    this.name = moduleName
    
    Dim modulePath As String: modulePath = fso.BuildPath(updateFolderPath, moduleName) 'путь к рассматриваемому модулю
    Dim md As file: Set md = fso.GetFile(modulePath)
    
    this.type = md.type
    this.dateModuleUpdate = md.DateLastModified
    
    this.dateModuleInstalled = LocalModuleDate() 'дата изменения установленного модуля
    this.isInstalled = this.dateModuleInstalled <> Base.defaultValue
    
    this.pathFileUpdate = fso.BuildPath(updateFolderPath, moduleName)
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 30.05.2023 20:57
    ' ----------------------------------------------------------------
    Private Function LocalModuleDate( _
        ) As String
        
        LocalModuleDate = Base.defaultValue
        
        Dim targetFolderPath As String: targetFolderPath = LocalModuleFolderPath()
        Dim localModulePath As String: localModulePath = fso.BuildPath(targetFolderPath, this.name)
        
        If Not fso.FileExists(localModulePath) Then _
            Exit Function
            
        this.pathFileInstalled = localModulePath
        ' ----------------------------------------------------------------
        Dim lmd As file: Set lmd = fso.GetFile(localModulePath)
        Dim result As String: result = lmd.DateLastModified
        
        LocalModuleDate = result
        
    End Function
        ' ----------------------------------------------------------------
        ' Дата: 30.05.2023 21:00
        ' ----------------------------------------------------------------
        Private Function LocalModuleFolderPath( _
            ) As String
            
            LocalModuleFolderPath = Base.configDataPath
            
            Select Case True
            
                Case this.type = TYPE_MODULE_EXCEL
                    LocalModuleFolderPath = Base.excelAddinsDir
                
                Case this.type = TYPE_MODULE_WORD
                    LocalModuleFolderPath = Base.wordAddinsDir
                
            End Select
            
        End Function
Friend Property Get dateModuleUpdate() As String
    dateModuleUpdate = this.dateModuleUpdate
End Property
Friend Property Get dateModuleInstalled() As String
    dateModuleInstalled = this.dateModuleInstalled
End Property
Friend Property Get isInstalled() As Boolean
    isInstalled = this.isInstalled
End Property
Friend Property Get name() As String
    name = this.name
End Property
Friend Property Get pathFileUpdate() As String
    pathFileUpdate = this.pathFileUpdate
End Property
Friend Property Get pathFileInstalled() As String
    pathFileInstalled = this.pathFileInstalled
End Property
' ----------------------------------------------------------------
' Дата: 30.05.2023 21:44 -- закрыть книгу Excel | документ Word
' ----------------------------------------------------------------
Friend Sub PrepareToInstall( _
    )

    Select Case True
    
        Case this.type = TYPE_MODULE_EXCEL
            UnloadExcelModule
        
        Case this.type = TYPE_MODULE_WORD
        
            If Base.IsWordExecuting Then _
                UnloadWordModule
            
    End Select
    
End Sub
    Private Sub UnloadExcelModule()
        Application.Workbooks(this.name).Close
    End Sub
    Private Sub UnloadWordModule()
        
        Dim wordApp As Object
        Set wordApp = GetObject(, WORD_APPLICATION_NAME)
        ' ----------------------------------------------------------------
        wordApp.AddIns(this.pathFileInstalled).Installed = False
        ' ----------------------------------------------------------------
        Set wordApp = Nothing
        
    End Sub

