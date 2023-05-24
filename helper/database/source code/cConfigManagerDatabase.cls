VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigManagerDatabase"
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

Public Enum EFolderTypes

    sandboxDir
    archiveLocalDir
    sourceDir

End Enum

Private Type TManagerProperties

    sourceDataPath As String
    sandboxPath As String
    archiveLocalPath As String
    verifierKey As String
    executorKey As String
    
End Type

Private this As TManagerProperties, _
    cs As New cConfigSource, _
    fso As New FileSystemObject, _
    keys As New cConfigKeysDatabase
       
Private Sub Class_Initialize()

    cs.Initialize _
        sourceType:=configData, _
        nameKey:=keys.Header
    
    this.sourceDataPath = cs.FindValue(keys.sourceDataPath)
    this.sandboxPath = cs.FindValue(keys.sandboxPath)
    this.archiveLocalPath = cs.FindValue(keys.archiveLocalPath)
    
    this.verifierKey = cs.FindValue(keys.verifierKey)
    this.executorKey = cs.FindValue(keys.executorKey)
    
    If this.sourceDataPath <> Base.defaultValue Then _
        CheckDirs
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 12.03.2023 14:01
    ' Назначение:
    ' ----------------------------------------------------------------
    Private Sub CheckDirs( _
        )
     
        If Not fso.FolderExists(organisationsPath) Then _
            fso.CreateFolder organisationsPath
        
        If Not fso.FolderExists(instrumentsPath) Then _
            fso.CreateFolder instrumentsPath
            
        If Not fso.FolderExists(etalonsPath) Then _
            fso.CreateFolder etalonsPath
        
        If Not fso.FolderExists(personsPath) Then _
            fso.CreateFolder personsPath
        
    End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Public Property Get sourceDataPath() As String
    sourceDataPath = this.sourceDataPath
End Property
Friend Property Let sourceDataPath(ByVal NewValue As String)
    
    this.sourceDataPath = Base.defaultValue
    
    If NewValue <> vbNullString Then
    
        this.sourceDataPath = NewValue
        CheckDirs
        
    End If
    
End Property
' ----------------------------------------------------------------
Friend Property Get sandboxPath() As String
    sandboxPath = this.sandboxPath
End Property
Friend Property Let sandboxPath(ByVal NewValue As String)
    
    this.sandboxPath = Base.defaultValue
    If NewValue <> vbNullString Then _
        this.sandboxPath = NewValue
        
End Property
' ----------------------------------------------------------------
Friend Property Get organisationsPath() As String
    organisationsPath = fso.BuildPath(this.sourceDataPath, keys.organisationsDir)
End Property
' ----------------------------------------------------------------
Public Property Get instrumentsPath() As String
    instrumentsPath = fso.BuildPath(this.sourceDataPath, keys.instrumentsDir)
End Property
Public Property Get templatesBasePath() As String
    templatesBasePath = fso.BuildPath(instrumentsPath, keys.templatesBaseDir)
End Property
' ----------------------------------------------------------------
Friend Property Get etalonsPath() As String
    etalonsPath = fso.BuildPath(this.sourceDataPath, keys.etalonsDir)
End Property
' ----------------------------------------------------------------
Public Property Get personsPath() As String
    personsPath = fso.BuildPath(this.sourceDataPath, keys.personsDir)
End Property
' ----------------------------------------------------------------
Friend Property Get archiveLocalPath() As String
    archiveLocalPath = this.archiveLocalPath
End Property
Friend Property Let archiveLocalPath(ByVal NewValue As String)
    
    this.archiveLocalPath = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        this.archiveLocalPath = NewValue
        
End Property
' ----------------------------------------------------------------
Public Property Get verifierKey() As String
    verifierKey = this.verifierKey
End Property
Friend Property Let verifierKey(ByVal NewValue As String)
    
    this.verifierKey = Base.defaultValue
    If NewValue <> vbNullString Then _
        this.verifierKey = NewValue
        
End Property
' ----------------------------------------------------------------
Friend Property Get executorKey() As String
    executorKey = this.executorKey
End Property
Friend Property Let executorKey(ByVal NewValue As String)
    
    this.executorKey = Base.defaultValue
    If NewValue <> vbNullString Then _
        this.executorKey = NewValue
        
End Property

' ----------------------------------------------------------------
' Назначение: Выбор каталога
'    параметр initialSomniumPath:
' Дата: 09.03.2023 09:14
' ----------------------------------------------------------------
Friend Function ChooseFolderPath( _
    folderType As EFolderTypes, _
    Optional initialSomniumPath As Boolean = False _
    ) As Boolean
    
    ChooseFolderPath = False
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    fd.Filters.Clear
    fd.InitialView = msoFileDialogViewDetails

    fd.Title = PikerTitle(folderType)
    fd.AllowMultiSelect = False

    fd.InitialFileName = Base.desktopPath

    If initialSomniumPath Then _
        fd.InitialFileName = Base.serverInitialPath

    If fd.Show = False Then _
        Set fd = Nothing: _
        Exit Function
    
    SetPath _
        folderType:=folderType, _
        path:=fd.SelectedItems(1) 'полный путь к файлу
    
    Set fd = Nothing
    ChooseFolderPath = True

End Function
    ' ----------------------------------------------------------------
    ' Дата: 09.03.2023 09:27
    ' Назначение: Заголовок меню выбора каталога
    ' ----------------------------------------------------------------
    Private Function PikerTitle( _
        folderType As EFolderTypes _
        ) As String
        
        Select Case True
                
            Case folderType = sourceDir
                PikerTitle = "Выбор исходного каталога данных"
            
            Case folderType = sandboxDir
                PikerTitle = "Выбор каталога песочницы загрузки протоколов"
                 
            Case folderType = archiveLocalDir
                PikerTitle = "Выбор каталога локального архива"
            
            Case Else
                PikerTitle = "Выборы в президенты 2024"
        
        End Select
    
    End Function
    Private Sub SetPath( _
        folderType As EFolderTypes, _
        path As String _
        )
        
        Select Case True
                
            Case folderType = sourceDir
                this.sourceDataPath = path
            
            Case folderType = sandboxDir
                this.sandboxPath = path
                
            Case folderType = archiveLocalDir
                this.archiveLocalPath = path
        
        End Select
        
    End Sub
' ----------------------------------------------------------------
Public Function IsCorrect() As Boolean
    
    IsCorrect = False
    
    If Not fso.FolderExists(this.sourceDataPath) Then _
        Exit Function
    
    If Not fso.FolderExists(this.sandboxPath) Then _
        Exit Function
    
    If Not fso.FolderExists(this.archiveLocalPath) Then _
        Exit Function
    
    IsCorrect = True
                
End Function
' ----------------------------------------------------------------
Friend Sub Save()

    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.sourceDataPath, this.sourceDataPath
    cs.PrepareToWrite keys.sandboxPath, this.sandboxPath
    cs.PrepareToWrite keys.archiveLocalPath, this.archiveLocalPath
    
    cs.PrepareToWrite keys.verifierKey, this.verifierKey
    cs.PrepareToWrite keys.executorKey, this.executorKey
    
    cs.WriteToFile
    
End Sub
