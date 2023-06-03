VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 17.04.2023 22:23
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit

Public Enum EInstallModes
    
    normalMode
    silentMode
    
End Enum

Const KEYWORD_ADDIN = "helper_"
Const WORD_APPLICATION_NAME = "Word.Application"

Const DELIMITER_PARSE = "#"
Private fso As New FileSystemObject
' ----------------------------------------------------------------
' Дата: 06.12.2021 14:01, 29.05.2023 22:32
' ----------------------------------------------------------------
Public Sub Check( _
    )

    Dim modulesUpdateAvailable As Collection
    Set modulesUpdateAvailable = UpdateList(Config.pathUpdateFolder)
    ' ----------------------------------------------------------------
    Cache.isUpdateAvailable = modulesUpdateAvailable.count > vbEmpty
    Cache.modulesUpdateAvailable = StringFromColl(modulesUpdateAvailable)
    
    Cache.dateOfCheck = Date
    Cache.Save

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 30.05.2023 21:16
    ' ----------------------------------------------------------------
    Private Function UpdateList( _
        path As String _
        ) As Collection

        If Not fso.FolderExists(Config.pathUpdateFolder) Then _
            Exit Function
        ' ----------------------------------------------------------------
        Dim updFolder As Folder
        Set updFolder = fso.GetFolder(path)
        
        Dim module As file, _
            listUpdateReady As New Collection, _
            md As cHelperVBProject
        ' ----------------------------------------------------------------
        For Each module In updFolder.Files
            
            Set md = New cHelperVBProject
            md.Initialize path, module.name 'получить данные рассматриваемого модуля
            
            Dim isUpdateAvailable As Boolean
            If md.isInstalled Then _
                isUpdateAvailable = md.dateModuleInstalled < md.dateModuleUpdate
                
            If isUpdateAvailable Then _
                listUpdateReady.Add module.name
            
        Next
        ' ----------------------------------------------------------------
        Set UpdateList = listUpdateReady
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 29.05.2023 22:17
    ' ----------------------------------------------------------------
    Private Function StringFromColl( _
        coll As Collection _
        ) As String
        
        StringFromColl = Base.defaultValue 'null по умолчанию
        
        If coll.count = vbEmpty Then _
            Exit Function
            
        Dim i As Integer, result As String
        For i = 1 To coll.count
        
            If result <> vbNullString Then _
                result = result & DELIMITER_PARSE
            
            result = result & coll(i)
            
        Next i
        
        StringFromColl = result
        
    End Function
' ----------------------------------------------------------------
' Дата: 30.05.2023 22:06
' ----------------------------------------------------------------
Public Sub Install( _
    installMode As EInstallModes _
    )
    
    If Not Cache.isUpdateAvailable Then _
        Exit Sub
    ' ----------------------------------------------------------------
    Dim UpdateList() As String: UpdateList = Split(Cache.modulesUpdateAvailable, DELIMITER_PARSE)
    
    Dim i As Integer, md As cHelperVBProject
    For i = LBound(UpdateList) To UBound(UpdateList)
        
        Set md = New cHelperVBProject
        md.Initialize Config.pathUpdateFolder, UpdateList(i)
        
        md.PrepareToInstall 'закрыть книгу Excel | документ Word перед установкой
        fso.CopyFile md.pathFileUpdate, md.pathFileInstalled
        
    Next
    
    If installMode = normalMode Then _
        ReloadModules
    
    Cache.isUpdateAvailable = False
    Cache.Save
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 30.05.2023 22:28
    ' ----------------------------------------------------------------
    Private Sub ReloadModules()
        
        UnloadExcelAddins
        LoadExcelAddins
        ' ----------------------------------------------------------------
        UnloadWordAddins
        LoadWordAddins
        
    End Sub
        Private Sub UnloadExcelAddins()
            
            Dim md As Addin
            For Each md In Application.AddIns2
                
                If IsHelperAddin(md.name) Then _
                    Application.Workbooks(md.name).Close
                        
            Next
            
        End Sub
            ' ----------------------------------------------------------------
            ' Дата: 30.05.2023 22:18
            ' ----------------------------------------------------------------
            Private Function IsHelperAddin( _
                addinName As String _
                ) As Boolean
                
                IsHelperAddin = False
                
                If addinName <> ThisWorkbook.name Then _
                    IsHelperAddin = InStr(addinName, KEYWORD_ADDIN) = 1
                
            End Function
        Private Sub LoadExcelAddins()
        
            Dim mdFolder As Folder
            Set mdFolder = fso.GetFolder(Base.excelAddinsDir)
            
            Dim md As file
            For Each md In mdFolder.Files
            
                If IsHelperAddin(md.name) Then _
                    Application.Workbooks.Open md.path
            
            Next
            
        End Sub
        Private Sub UnloadWordAddins()
            
            If Not Base.IsWordExecuting Then _
                Exit Sub
            ' ----------------------------------------------------------------
           
            Dim wordApp As Object
            Set wordApp = GetObject(, WORD_APPLICATION_NAME)
            '----------------------------------------------------------------
            Dim mdFolder As Folder
            Set mdFolder = fso.GetFolder(Base.wordAddinsDir)
            
            Dim md As file
            For Each md In mdFolder.Files
            
                If IsHelperAddin(md.name) Then _
                    wordApp.AddIns(md.name).Installed = False
            
            Next
            
            ' ----------------------------------------------------------------
            Set wordApp = Nothing
            
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 30.05.2023 22:28
        ' ----------------------------------------------------------------
        Private Sub LoadWordAddins()
            
            If Not Base.IsWordExecuting Then _
                Exit Sub
            ' ----------------------------------------------------------------
           
            Dim wordApp As Object
            Set wordApp = GetObject(, WORD_APPLICATION_NAME)
            '----------------------------------------------------------------
            Dim mdFolder As Folder
            Set mdFolder = fso.GetFolder(Base.wordAddinsDir)
            
            Dim md As file
            For Each md In mdFolder.Files
            
                If IsHelperAddin(md.name) Then _
                    wordApp.AddIns.Add md.path, True
            
            Next
            
            ' ----------------------------------------------------------------
            Set wordApp = Nothing
                        
        End Sub

