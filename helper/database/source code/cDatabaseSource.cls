VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabaseSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cItemOperator (Class Module)
' Назначение: получение сведений из файла конкретного элемента БД
' Автор: akolodka
' Date: 24.02.2023 19:29
' ------------------------------------------------------

Option Explicit

Const FILE_EXT = "txt"
Const NOCACHE_FILENAME = ".nocache"
Const BODIES_DIR = "_templates body"

Private Type TItemData

    filePath As String
    fileData As String
    fileNewData As String
    parentFolderPath As String
    key As String
    isHidden As Boolean
    
End Type

Private fso As New FileSystemObject, _
        this As TItemData

Private Sub Class_Initialize()
    
    this.filePath = Base.defaultValue
    this.fileData = Base.defaultValue
    this.fileNewData = Base.defaultValue
    this.parentFolderPath = Base.defaultValue
    this.key = Base.defaultValue
    this.isHidden = Base.defaultValue(boolDefault)
    
End Sub
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:35
' Назначение: получение первичных данных
' ----------------------------------------------------------------
Friend Sub Initialize( _
    sourcePath As String, _
    key As String _
    )
    
    this.key = key
    this.parentFolderPath = fso.BuildPath(sourcePath, key)
    this.filePath = fso.BuildPath(this.parentFolderPath, key & "." & FILE_EXT)
    this.fileData = Base.defaultValue
    this.fileNewData = Base.defaultValue
    
    Dim fPath As String
    fPath = fso.BuildPath(this.parentFolderPath, NOCACHE_FILENAME)
    this.isHidden = fso.FileExists(fPath)
    
End Sub
Public Sub OpenForEdit( _
    )
    Explorer.OpenFolder this.filePath, isFocusOnWindow:=True
End Sub
Public Property Get key() As String
    key = this.key
End Property
Public Property Get itemFolderPath() As String
    itemFolderPath = this.parentFolderPath
End Property
Public Property Get itemFilePath() As String
    itemFilePath = this.filePath
End Property

Public Property Get isHidden() As Boolean
    isHidden = this.isHidden
End Property
' ----------------------------------------------------------------
' Дата: 19.03.2023 08:30
' ----------------------------------------------------------------
Public Property Get bodyTemplatesPath() As String
    bodyTemplatesPath = fso.BuildPath(this.parentFolderPath, BODIES_DIR)
End Property
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:58
' Назначение:
' ----------------------------------------------------------------
Public Sub OpenItemDirectory()
    
    If this.parentFolderPath = Base.defaultValue Then _
        Handler.Notify "Каталог элемента базы данных недоступен.": _
        Exit Sub
    
    Explorer.OpenFolder this.parentFolderPath, True
    
End Sub
' ----------------------------------------------------------------
' Дата: 18.03.2023 18:05
' ----------------------------------------------------------------
Public Function GetInstrumentGroup( _
    ) As Collection
    
    If fso.FileExists(this.filePath) Then _
        this.fileData = Base.ContentFromFile(this.filePath)
    
    If this.fileData = Base.defaultValue Then _
        Exit Function
        
    Dim isContentCorrect As Boolean
    isContentCorrect = Not CBool(InStr(this.fileData, Base.delimiterKeyValue)) 'не должно быть никаких разделителей
    
    If Not isContentCorrect Then _
        Exit Function
        
    Dim splitArr() As String
    splitArr = Split(this.fileData, vbNewLine)
    
    Dim i As Integer, result As New Collection
    For i = LBound(splitArr) To UBound(splitArr)
        result.Add splitArr(i)
    Next
    
    Set GetInstrumentGroup = result
    this.fileData = Base.defaultValue

End Function
' ----------------------------------------------------------------
' Дата: 18.03.2023 18:14
' ----------------------------------------------------------------
Friend Function FindValue( _
    searchKey As String, _
    Optional searchDefault As EDefaultTypes = defaultValue, _
    Optional subKey As String _
    ) As String
    
    FindValue = Base.defaultValue
    
    If subKey <> vbNullString Then _
        this.filePath = fso.BuildPath(this.parentFolderPath, subKey & "." & FILE_EXT): _
        this.fileData = Base.ContentFromFile(this.filePath)
    
    If this.fileData = Base.defaultValue Then _
        this.fileData = Base.ContentFromFile(this.filePath)
    
    Dim SourceData As String
    SourceData = this.fileData
    
    If SourceData = Base.defaultValue Then _
        Exit Function
        
    Dim keyInstr As Integer
    keyInstr = InStr(SourceData, searchKey & Base.delimiterKeyValue) 'первое вхождение ключа параметра
    
    If keyInstr = 0 Then Exit Function 'ключ отсутствует
     
    Dim findedValue As String
    findedValue = findResult( _
        SourceData:=SourceData, _
        keyInstr:=keyInstr _
        )
    
    FindValue = Replace(findedValue, Base.defaultValuePrinted, Base.defaultValue)
    
End Function
    Private Function findResult( _
        ByVal SourceData As String, _
        ByVal keyInstr As Integer _
        ) As String
        
        Dim valueInStr As Integer, _
        valueEndStr As Integer
    
        valueInStr = InStr(keyInstr, SourceData, Base.delimiterKeyValue) + 1 'первое вхождение значения по ключу
        valueEndStr = InStr(valueInStr, SourceData, vbNewLine) 'окончание значения по ключу
        
        If valueEndStr = 0 Then _
            valueEndStr = Len(SourceData) + 1
        
        findResult = Mid(SourceData, valueInStr, valueEndStr - valueInStr)
        
    End Function
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:52
' Назначение: подготовить данные к записи
' ----------------------------------------------------------------
Friend Sub PrepareToWrite( _
    key As String, _
    value As String _
    )
    
    Dim result As String
    If this.fileNewData <> Base.defaultValue Then _
        result = this.fileNewData & vbNewLine
    
    result = result & Trim(key) & Base.delimiterKeyValue & Trim(value)
    this.fileNewData = result
    
End Sub
' ----------------------------------------------------------------
' Дата: 12.03.2023 11:59
' Назначение:
' ----------------------------------------------------------------
Public Sub CreateItemDir()
    
    If Not fso.FolderExists(this.parentFolderPath) Then _
        fso.CreateFolder (this.parentFolderPath)
        
End Sub
' ----------------------------------------------------------------
' Дата: 19.03.2023 08:25
' ----------------------------------------------------------------
Friend Sub WriteToFile( _
    Optional key As String _
    )

    Dim result As String
    result = Replace(this.fileNewData, Base.defaultValue, Base.defaultValuePrinted)
    
    If key <> vbNullString Then 'только для Instruments
        
        Dim targetPath As String
        targetPath = fso.BuildPath(this.parentFolderPath, key & "." & FILE_EXT) '21730-08.txt
        
        Base.WriteContent targetPath, result
        result = InstrumentBaseContent(key)
        ' ----------------------------------------------------------------
        Dim bodiesPath As String
        bodiesPath = fso.BuildPath(this.parentFolderPath, BODIES_DIR)
        
        If Not fso.FolderExists(bodiesPath) Then _
            fso.CreateFolder bodiesPath
        ' ----------------------------------------------------------------
        
    End If

    Base.WriteContent this.filePath, result
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 18.03.2023 15:29
    ' ----------------------------------------------------------------
    Private Function InstrumentBaseContent( _
        key As String _
        ) As String
        
        InstrumentBaseContent = Base.defaultValue
        
        If fso.FileExists(this.filePath) Then _
            this.fileData = Base.ContentFromFile(this.filePath)
        ' ----------------------------------------------------------------
        If this.fileData = Base.defaultValue Then
            this.fileNewData = key
        
        Else
            
            this.fileNewData = this.fileData
            
            If Not CBool(InStr(this.fileData, key)) Then _
                this.fileNewData = this.fileData & vbNewLine & Trim(key) 'ключ не опознан в файле
        
        End If
        ' ----------------------------------------------------------------
        Dim result As String
        result = Replace(this.fileNewData, Base.defaultValue, Base.defaultValuePrinted)
        InstrumentBaseContent = result
        
    End Function
'' ----------------------------------------------------------------
'Public Sub HideFromCaching()
'
'    Dim path As String
'    path = fso.BuildPath(this.parentFolderPath, NOCACHE_FILENAME)
'
'    Base.WriteContent path, vbNullString
'
'End Sub
'' ----------------------------------------------------------------
'Public Sub ShowForCaching()
'
'    Dim path As String
'    path = fso.BuildPath(this.parentFolderPath, NOCACHE_FILENAME)
'
'    If fso.FileExists(path) Then _
'        fso.DeleteFile (path)
'
'End Sub
