VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 09.03.2023 13:00
' Автор: akolodka
' Назначение: оператор файлов конфигуркции
' ----------------------------------------------------------------

Option Explicit

Const CONFIG_EXT = "ini"
Const CACHE_EXT = "txt"

Const SUBFOLDER_CONFIG = "config"
Const SUBFOLDER_CACHE = "cache"

Const FSO_READ = 1
Const FSO_WRITE = 2

Private Type TConfigData

    filePath As String
    fileData As String
    fileNewData As String
    parentFolderPath As String
    
End Type
 
Public Enum ESourceTypes

    configData
    cacheData
    
End Enum

Public Enum EDefaultTypes
    
    defaultValue
    boolValue
    nullStringValue
    
End Enum

Private fso As New FileSystemObject
Private this As TConfigData

Private Sub Class_Initialize()
       
    this.filePath = Base.defaultValue
    this.fileData = Base.defaultValue
    this.fileNewData = Base.defaultValue
    this.parentFolderPath = Base.defaultValue
    
End Sub
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:13
' Назначение: подготовка к работе с конфигурацией
' ----------------------------------------------------------------
Friend Sub Initialize( _
    sourceType As ESourceTypes, _
    nameKey As String _
    )
    
    Dim targetFolder As String, targetExt As String
    Select Case True
            
        Case sourceType = configData
        
            targetFolder = SUBFOLDER_CONFIG
            targetExt = CONFIG_EXT
        
        Case sourceType = cacheData
            
            targetFolder = SUBFOLDER_CACHE
            targetExt = CACHE_EXT
    
    End Select
    
    this.parentFolderPath = fso.BuildPath(Base.configDataPath, targetFolder)
    this.filePath = fso.BuildPath(this.parentFolderPath, nameKey & "." & targetExt)
    
    If fso.FileExists(this.filePath) Then _
        this.fileData = Base.ContentFromFile(this.filePath)
    
End Sub

Public Property Get filePath()
    
    filePath = this.filePath
    
End Property
' ----------------------------------------------------------------
' Наименование: FindValue (Friend Function)
' Назначение: поиск параметра по ключу в конфигурации или кеше
'    параметр searchKey:
'    параметр whereToFind: config или cache
'    параметр boolResult: если возвращаемый параметр должен быть True или False
' Возвращаемый тип: String
' Дата: 09.03.2023 12:31
' ----------------------------------------------------------------
Friend Function FindValue( _
    searchKey As String, _
    Optional searchDefault As EDefaultTypes = defaultValue _
    ) As String
    
    FindValue = DefaultFinded(searchDefault)
    
    If this.fileData = Base.defaultValue Then _
        Exit Function
        
    Dim keyInstr As Integer
    keyInstr = InStr(this.fileData, searchKey & Base.delimiterKeyValue) 'первое вхождение ключа параметра
    
    If keyInstr = 0 Then _
        Exit Function 'ключ отсутствует
     
    Dim findedValue As String
    findedValue = findResult( _
        searchData:=this.fileData, _
        keyInstr:=keyInstr _
        )
    
    FindValue = Replace(findedValue, Base.defaultValuePrinted, Base.defaultValue)
    
End Function
    Private Function DefaultFinded( _
        searchDefault As EDefaultTypes _
        ) As String

        DefaultFinded = Base.defaultValue
        
        Select Case True
                
            Case searchDefault = boolValue
                DefaultFinded = False
                
            Case searchDefault = nullStringValue
                DefaultFinded = vbNullString
        
        End Select

    End Function
    Private Function findResult( _
        ByVal searchData As String, _
        ByVal keyInstr As Integer _
        ) As String
        
        Dim valueInStr As Integer, _
        valueEndStr As Integer
    
        valueInStr = InStr(keyInstr, searchData, Base.delimiterKeyValue) + 1 'первое вхождение значения по ключу
        valueEndStr = InStr(valueInStr, searchData, vbNewLine) 'окончание значения по ключу
        
        If valueEndStr = 0 Then _
            valueEndStr = Len(searchData) + 1
        
        findResult = Mid(searchData, valueInStr, valueEndStr - valueInStr)
        
    End Function
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:43
' Назначение: формирование наполнения файла конфигурации
' ----------------------------------------------------------------
Friend Sub PrepareToWrite( _
    Optional key As String, _
    Optional value As String _
    )
    
    If key = vbNullString And value = vbNullString Then _
        this.fileNewData = Base.defaultValue: _
        Exit Sub
    
    Dim result As String
    If this.fileNewData <> Base.defaultValue Then _
        result = this.fileNewData & vbNewLine
    
    result = result & Trim(key) & Base.delimiterKeyValue & Trim(value)
    this.fileNewData = result
    
End Sub
' ----------------------------------------------------------------
Friend Sub WriteToFile( _
    )
    
    If Not fso.FolderExists(this.parentFolderPath) Then _
        MkDir this.parentFolderPath
    
    Dim result As String
    result = Replace(this.fileNewData, Base.defaultValue, Base.defaultValuePrinted)
    
    Base.WriteContent this.filePath, result
    
End Sub




