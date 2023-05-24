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

Const MODULE_EXCEL = "Надстройка Microsoft Excel"
Const MODULE_WORD = "Шаблон Microsoft Word с поддержкой макросов"

Const WIN_PROCESS_WORD = "WINWORD.EXE"
Const WORD_APPLICATION_NAME = "Word.Application"

Private m_moduleDir As String, _
        m_moduleName As String, _
        m_isInstalled As Boolean, _
        m_isUpdateAvailable As Boolean

Private m_dateLastModified As String, _
        m_moduleType As String, _
        m_installedmoduleDir As String, _
        m_installedDateLastModified As String
        
Private fso As New FileSystemObject

Private Sub Class_Initialize()

    m_dateLastModified = Base.defaultValue
    m_moduleType = Base.defaultValue
    m_installedmoduleDir = Base.defaultValue
    m_installedDateLastModified = Base.defaultValue
    
End Sub
' ----------------------------------------------------------------
' Наименование: GetData (Public Sub)
' Назначение: Получить сведения об объекте обновления
'    параметр destDir:
'    параметр fileName:
' Дата: 04.12.2021 12:25
' ----------------------------------------------------------------
Public Sub GetData( _
    destDir As String, _
    fileName As String _
    )
    
    If destDir = Base.defaultValue Or _
       fileName = Base.defaultValue Then Exit Sub
    
    m_moduleDir = destDir
    m_moduleName = fileName
    
    Dim module As file
    Set module = fso.GetFile(fso.BuildPath(destDir, fileName))
    
    m_dateLastModified = module.dateLastModified
    m_moduleType = module.type
    
    GetInstallStatus
    GetInstalledData
    
    m_isUpdateAvailable = False
    If IsDate(m_dateLastModified) And IsDate(m_installedDateLastModified) Then
    
        If CDate(m_dateLastModified) > CDate(m_installedDateLastModified) Then _
            m_isUpdateAvailable = True
            
    End If
    
    Set module = Nothing
    
'Debug.Print
'Debug.Print "Каталог обновления = "; m_moduleDir
'Debug.Print "Модуль = "; fileName
'Debug.Print "Тип = "; m_moduleType
'Debug.Print "Установлено = "; m_isInstalled
'Debug.Print "Дата изменения модуля = "; m_dateLastModified
'Debug.Print "Дата изменения установленного модуля = "; m_installedDateLastModified
'Debug.Print "Доступность обновления = "; m_isUpdateAvailable

End Sub
    Private Sub GetInstallStatus( _
        )
        
        m_isInstalled = False

        Select Case m_moduleType
        
            Case MODULE_EXCEL
                m_installedmoduleDir = Base.excelAddinsDir
                
            Case MODULE_WORD
                m_installedmoduleDir = Base.wordAddinsDir
            
            Case Else
                m_installedmoduleDir = Base.configDataPath

        End Select
        
        If fso.FileExists(fso.BuildPath(m_installedmoduleDir, m_moduleName)) Then _
            m_isInstalled = True
                
    End Sub
    Private Sub GetInstalledData( _
        )
        If Not m_isInstalled Then _
            Exit Sub
        
        Dim installedModule As file
        Set installedModule = fso.GetFile(fso.BuildPath(m_installedmoduleDir, m_moduleName))
        m_installedDateLastModified = installedModule.dateLastModified
        
        Set installedModule = Nothing
    End Sub
Friend Property Get IsUpdateAvailable( _
    ) As Boolean
    
    IsUpdateAvailable = m_isUpdateAvailable

End Property

Public Sub Update( _
    Optional isReload As Boolean = False _
    )
    
    If Not m_isUpdateAvailable Then _
        Exit Sub
    
    
    UnloadModules
    
    fso.CopyFile fso.BuildPath(m_moduleDir, m_moduleName), _
                 fso.BuildPath(m_installedmoduleDir, m_moduleName), _
                 True
    
    If isReload Then _
        ReloadModules
  '  Debug.Print "----------------------------------------------------------------"
End Sub
    Private Sub UnloadModules( _
        )
            
        Select Case m_moduleType
        
            Case MODULE_EXCEL
                ' ----------------------------------------------------------------
                Application.Workbooks(m_moduleName).Close
                ' ----------------------------------------------------------------
                
            Case MODULE_WORD
                
                If isWordExecuting Then
                    
                    Dim wordApp As Object
                    Set wordApp = GetObject(, WORD_APPLICATION_NAME)
                    ' ----------------------------------------------------------------
                    wordApp.AddIns(fso.BuildPath(m_installedmoduleDir, m_moduleName)).Installed = False
                    ' ----------------------------------------------------------------
                    Set wordApp = Nothing
                    
                End If
                
        End Select
        
    End Sub
        Private Function isWordExecuting( _
            ) As Boolean
            
            isWordExecuting = False
            
            Dim winProcess As Object
            For Each winProcess In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process") 'пройтись по процессам компьютера
                
                If winProcess.caption Like WIN_PROCESS_WORD Then _
                    isWordExecuting = True: _
                    Exit Function
            Next
            
        End Function
    Private Sub ReloadModules( _
        )
        
        Select Case m_moduleType
    
            Case MODULE_EXCEL
                ' ----------------------------------------------------------------
                Application.Workbooks.Open (fso.BuildPath(m_installedmoduleDir, m_moduleName))
                ' ----------------------------------------------------------------
            Case MODULE_WORD
            
                If isWordExecuting Then
                    
                    Dim wordApp As Object
                    Set wordApp = GetObject(, WORD_APPLICATION_NAME)
                    ' ----------------------------------------------------------------
                    wordApp.AddIns.Add fso.BuildPath(m_installedmoduleDir, m_moduleName), True
                    ' ----------------------------------------------------------------
                    Set wordApp = Nothing
                    
                End If

        End Select
        
    End Sub
