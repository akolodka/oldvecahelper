VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigManagerContracts"
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

    sourceF
    destinationF

End Enum

Private Type TManagerProperties

    sourceFolder As String
    destinationFolder As String
    filterCriteria As String
    
End Type

Private this As TManagerProperties, _
    cs As New cConfigSource, _
    fso As New FileSystemObject, _
    keys As New cConfigKeysContracts
       
Private Sub Class_Initialize()

    cs.Initialize _
        sourceType:=configData, _
        nameKey:=keys.Header
    
    this.sourceFolder = cs.FindValue(keys.sourceFolder)
    this.destinationFolder = cs.FindValue(keys.destinationFolder)
    this.filterCriteria = cs.FindValue(keys.filterCriteria)

End Sub

Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Public Property Get sourceFolder() As String
    sourceFolder = this.sourceFolder
End Property
Friend Property Let sourceFolder(ByVal NewValue As String)
    
    this.sourceFolder = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        this.sourceFolder = NewValue
        
    
End Property
' ----------------------------------------------------------------
Friend Property Get destinationFolder() As String
    destinationFolder = this.destinationFolder
End Property
Friend Property Let destinationFolder(ByVal NewValue As String)
    
    this.destinationFolder = Base.defaultValue
    If NewValue <> vbNullString Then _
        this.destinationFolder = NewValue
        
End Property

Friend Property Get filterCriteria() As String
    filterCriteria = this.filterCriteria
End Property
Friend Property Let filterCriteria(ByVal NewValue As String)
    
    this.filterCriteria = Base.defaultValue
    If NewValue <> vbNullString Then _
        this.filterCriteria = NewValue
        
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
                
            Case folderType = sourceF
                PikerTitle = "Выбор каталога исходных файлов"
            
            Case folderType = destinationF
                PikerTitle = "Выбор каталога назначения"
            
            Case Else
                PikerTitle = "Выборы в президенты 2024"
        
        End Select
    
    End Function
    Private Sub SetPath( _
        folderType As EFolderTypes, _
        path As String _
        )
        
        Select Case True
                
            Case folderType = sourceF
                this.sourceFolder = path
            
            Case folderType = destinationF
                this.destinationFolder = path
        
        End Select
        
    End Sub
' ----------------------------------------------------------------
Public Function IsCorrect() As Boolean
    
    IsCorrect = False
    
    If Not fso.FolderExists(this.sourceFolder) Then _
        Exit Function
    
    If Not fso.FolderExists(this.destinationFolder) Then _
        Exit Function
    
    IsCorrect = True
                
End Function
' ----------------------------------------------------------------
Friend Sub Save()

    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.sourceFolder, this.sourceFolder
    cs.PrepareToWrite keys.destinationFolder, this.destinationFolder
    cs.PrepareToWrite keys.filterCriteria, this.filterCriteria

    cs.WriteToFile
    
End Sub
