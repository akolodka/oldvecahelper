VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCharTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 15.03.2023 18:31
' Автор: akolodka
' Назначение: для обработки строки независимо от раскладки клавиатуры
' ----------------------------------------------------------------

Option Explicit
Option Compare Text

Const CHARTABLE_FILENAME = "chartable.pkr"

Private Type TCharProps

    isAvailable As Boolean
    rusChars As String
    engChars As String
    
End Type

Private Enum EParseTypes
    
    rusToEng
    engToRus
    
End Enum

Private this As TCharProps, _
        fso As New FileSystemObject

Private Sub Class_Initialize()

    Dim filePath As String
    filePath = fso.BuildPath(Config.sourceDataPath, CHARTABLE_FILENAME)
    
    this.isAvailable = fso.FileExists(filePath)
    If Not this.isAvailable Then _
        Exit Sub
        
    LoadCharTable filePath
        
End Sub
    Private Sub LoadCharTable( _
        path As String _
        )
    
        Dim fileData As String
        fileData = Base.ContentFromFile(path)
        
        Dim langDelimiter As Integer
        langDelimiter = InStr(fileData, vbNewLine)
        
        this.rusChars = Left(fileData, langDelimiter - 1)
        this.engChars = Mid(fileData, langDelimiter + 2, Len(fileData) - langDelimiter) 'vbNewLine занимает 2 символа
        
    End Sub

Public Property Get isAvaiable( _
    ) As Boolean
    
    isAvaiable = this.isAvailable
    
End Property
' ----------------------------------------------------------------
' Дата: 17.03.2023 09:18
' ----------------------------------------------------------------
Public Function Reverse( _
    text As String _
    ) As String
        
    Reverse = text
    
    If Not this.isAvailable Then _
        Exit Function
    ' ----------------------------------------------------------------
    Dim i As Integer, result As String
    For i = 1 To Len(text)
    
        Dim ch As String: ch = Mid(text, i, 1)
        Dim revChar As String: revChar = ch
        
        Dim isParsed As Boolean
        isParsed = TryParse(revChar, rusToEng)
        
        If Not isParsed Then _
            isParsed = TryParse(revChar, engToRus)
            
        result = result & revChar
        
    Next

    Reverse = result

End Function
    ' ----------------------------------------------------------------
    ' Дата: 17.03.2023 09:00
    ' ----------------------------------------------------------------
    Private Function TryParse( _
        ByRef char As String, _
        pType As EParseTypes _
        ) As Boolean
        
        TryParse = False
        
        Dim source As String
        source = SourceCharData(pType)

        Dim charPos As Integer
        charPos = InStr(source, char)
        ' ----------------------------------------------------------------
        If Not CBool(charPos) Then _
            Exit Function
        ' ----------------------------------------------------------------
        Dim target As String
        target = TargetCharData(pType)
        
        char = Mid(target, charPos, 1)
        TryParse = True
        
    End Function
        ' ----------------------------------------------------------------
        ' Дата: 17.03.2023 08:57
        ' ----------------------------------------------------------------
        Private Function SourceCharData( _
            pType As EParseTypes _
            ) As String
            
            Dim result As String
            result = this.rusChars
            
            If pType = engToRus Then _
                result = this.engChars
                
            SourceCharData = result

        End Function
        ' ----------------------------------------------------------------
        ' Дата: 17.03.2023 08:58
        ' ----------------------------------------------------------------
        Private Function TargetCharData( _
            pType As EParseTypes _
            ) As String
            
            Dim result As String
            result = this.engChars
            
            If pType = engToRus Then _
                result = this.rusChars
                
            TargetCharData = result

        End Function
'#########################################################
'процедура заполняет массив кейкодов
Public Sub RefactorKeycodeFile()
    
    Dim targetPath As String
    targetPath = fso.BuildPath(OldSettingsPath, "keyCode.npDb")
    
    If Not fso.FileExists(targetPath) Then _
        Handler.Notify "Файл " & targetPath & " не найден.": _
        Exit Sub
    
    Dim fileData As String
    fileData = Base.ContentFromFile(targetPath)
    
    Dim splitArr() As String
    splitArr = Split(fileData, vbNewLine)
    
    Dim i As Integer, fS As String
    For i = LBound(splitArr) To 32
    
        Dim strData() As String
        strData = Split(splitArr(i), Chr(9))
        
        fS = fS & strData(2)
        
    Next

    fS = fS & vbNewLine

    For i = LBound(splitArr) To 32

        strData = Split(splitArr(i), Chr(9))
        fS = fS & strData(3)
        
    Next
    fS = Replace(fS, "{", "[")
    
    Base.WriteContent fso.BuildPath(Config.sourceDataPath, "chartable.pkr"), fS
End Sub
    Private Function OldSettingsPath() As String
        OldSettingsPath = Environ("APPDATA") & "\Microsoft\Помощник ПКР\"
    End Function

