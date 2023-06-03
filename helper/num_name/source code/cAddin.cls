VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DEVELOPER_KEYWORD As String = "dev"
Const RELEASE_KEYWORD = "release"
Const IDE_KEYWORD = "VBA"

Const ENVIRON_KEYWORD = "USERPROFILE"
Const CLOUD_DIR = "Box\helper_updates"
Const SERVER_DIR = "\\somnium\irlab\Документы\Помощник ПКР\updates"
'Const SERVER_DIR = "C:\Users\Александр\Desktop\Server"

Enum e_updateTypes
    major
    medium
    micro
End Enum

Private fso As New FileSystemObject

Private m_addin As Workbook

Private Sub Class_Initialize()
    Set m_addin = ThisWorkbook
End Sub

Private Sub Class_Terminate()
    Set m_addin = Nothing
End Sub

Friend Property Get VersionCaption( _
    ) As String
    
    Properties.SetTargetWorkbook m_addin
    
    Dim strData As String
    strData = Replace(Properties.Comments, " — ", vbNewLine)
    
    If isDev Then _
        strData = "Dev build" & " " & Properties.Build & vbNewLine & _
        Format(Properties.Comments, "dd mmmm yyyy г.")
    
    Properties.SetTargetWorkbook ActiveWorkbook
    VersionCaption = strData
    
End Property
Friend Sub IncreaseBuild()
    
    Properties.SetTargetWorkbook m_addin
    
    If Not IsPropertiesCorrect Then _
        Initialialize
    
    IncreaseValue
    
    Properties.SetTargetWorkbook ActiveWorkbook
End Sub
    Private Function IsPropertiesCorrect( _
        )
        
        IsPropertiesCorrect = Properties.Version <> vbNullString _
                And Properties.Build <> vbNullString _
                And IsNumeric(Properties.Build)
        
    End Function
    Private Sub IncreaseValue( _
        )
        
        If Not isDev Then Exit Sub
                        
        Properties.Build = Properties.Build + 1
        Properties.Comments = Format(Date, "dd.mm.yyyy")
    
        Debug.Print "~~~ " & m_addin.name; " // "; _
                    Properties.Build; " // "; _
                    Properties.Version & " ~~~" '; " // "; version_date
                    
    End Sub
        Private Function isDev( _
            ) As Boolean
        
            isDev = False
            
            If Properties.Category = DEVELOPER_KEYWORD Then _
                isDev = True
            
        End Function
' ----------------------------------------------------------------
' Наименование: Initialialize (Friend Sub)
' Назначение: Принудительная очистка свойств надстройки
' Дата: 08.10.2022 14:18
' ----------------------------------------------------------------
Friend Sub Initialialize()
    
    Properties.SetTargetWorkbook m_addin
    
    Properties.Category = DEVELOPER_KEYWORD
    Properties.Keywords = IDE_KEYWORD
    Properties.Comments = Format(Date, "dd.mm.yyyy")
    Properties.Version = "0.1"
    Properties.Build = 0

    Debug.Print "~~~ Свойства успешно добавлены в " & m_addin.name & " ~~~"
    
    Properties.SetTargetWorkbook ActiveWorkbook
    
End Sub
Friend Sub Upload( _
    Optional updateType As e_updateTypes = micro _
    )
    
    Properties.SetTargetWorkbook m_addin
    
    Dim newVersionNum As String
    newVersionNum = NewVersion(updateType)
    
 '   newVersionNum = InputBox("Подтвердите изменение версии", m_addin.Name, newVersionNum)
    
    
    newVersionNum = InputBox("Текущая версия: " & Properties.Version & vbNewLine & _
            "Новая версия:", m_addin.name, newVersionNum)
            
'    newVersionNum = InputBox("Подтвердите изменение версии." & vbNewLine & vbNewLine & _
'            "Текущая версия: " & properties.version & vbNewLine & vbNewLine & _
'            "Новая версия:", m_addin.Name, newVersionNum)
'
    
    
    If newVersionNum = vbNullString Then _
        Debug.Print "~~~ !Версия не была введана, выполнение прервано ~~~": _
        Exit Sub
        
    Debug.Print
    
    DevUpload
    FreezeReleaseProperties newVersionNum 'подготовить свойства для релиза
    
    ReleaseUpload
    
    ReturnDevProperties
    m_addin.Save
    
    Properties.SetTargetWorkbook ActiveWorkbook
    Debug.Print "----------------------------"
    
End Sub
    Private Function NewVersion( _
        updateType As e_updateTypes _
        ) As String
        
        Dim currVersion As String
        currVersion = Properties.Version
        
        Dim arrVersion() As String
        arrVersion = Split(currVersion, ".")
        
        ReDim Preserve arrVersion(2)
        If arrVersion(LBound(arrVersion) + 2) = vbNullString Then _
            arrVersion(LBound(arrVersion) + 2) = 0
        
        Select Case True
        
            Case updateType = micro
                arrVersion(LBound(arrVersion) + 2) = CInt(arrVersion(LBound(arrVersion) + 2)) + 1
            
            Case updateType = medium
                arrVersion(LBound(arrVersion) + 1) = CInt(arrVersion(LBound(arrVersion) + 1)) + 1
                arrVersion(LBound(arrVersion) + 2) = vbNullString
                ReDim Preserve arrVersion(1)
                
            Case updateType = major
                arrVersion(LBound(arrVersion)) = CInt(arrVersion(LBound(arrVersion))) + 1
                arrVersion(LBound(arrVersion) + 1) = 0
                arrVersion(LBound(arrVersion) + 2) = vbNullString
                ReDim Preserve arrVersion(1)
                
        End Select
        
        NewVersion = Join(arrVersion, ".")
    End Function
    Private Sub DevUpload()
           
        If fso.FolderExists(CloudPath) Then
            
            m_addin.SaveCopyAs fso.BuildPath(CloudPath, m_addin.name)
            Debug.Print "~~~ Отправка в облако выполнена ~~~"
            
        Else
            Debug.Print "~~~ !Ошибка отправки в облако: каталог недоступен ~~~"
        End If

    End Sub
        Private Function CloudPath( _
            ) As String
            
            CloudPath = fso.BuildPath(Environ(ENVIRON_KEYWORD), CLOUD_DIR)
            
        End Function
    Private Sub FreezeReleaseProperties( _
        versionNumber As String _
        )
        
        Properties.Category = RELEASE_KEYWORD
        Properties.Version = versionNumber
        Properties.Comments = "v." & versionNumber & " (build " & Properties.Build & ") — " & Format(Date, "dd mmmm yyyy г.")
        
    End Sub
    Private Sub ReleaseUpload( _
        )
        If fso.FolderExists(SERVER_DIR) Then
        
            m_addin.SaveCopyAs fso.BuildPath(SERVER_DIR, m_addin.name)
            Debug.Print "~~~ Отправка на Somnium выполнена ~~~"
            
        Else
            m_addin.SaveCopyAs fso.BuildPath(Base.desktopPath, m_addin.name)
            Debug.Print "~~~ " & Time & " Отправлено на рабочий стол ~~~"
        End If
        
    End Sub
    Private Sub ReturnDevProperties()
        
        Properties.Category = DEVELOPER_KEYWORD
        Properties.Comments = Format(Date, "dd.mm.yyyy")
        Properties.Build = Properties.Build - 1
     
    End Sub
            
Friend Sub Download()

    Properties.SetTargetWorkbook ThisWorkbook

    Debug.Print
    If fso.FolderExists(CloudPath) Then
        
        Explorer.OpenFolder Properties.TargetWorkbook.path, True
        Explorer.OpenFolder CloudPath, True
        
    Else
        Debug.Print "~~~ Загрузка из облака... !Ошибка загрузки. Каталог недоступен ~~~"
        
        If fso.FolderExists(SERVER_DIR) Then
            
            Explorer.OpenFolder Properties.TargetWorkbook.path, True
            Explorer.OpenFolder SERVER_DIR, True
            
        Else
            Debug.Print "~~~ Загрузка c сервера... !Ошибка загрузки. Каталог недоступен ~~~"
        End If
        
    End If
    
    Properties.TargetWorkbook.Close False
    
End Sub
