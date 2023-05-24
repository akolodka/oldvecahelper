VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Z_clsmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sArrVal() As String, sArrParam() As String
'#######################################################
'страница работы с базами данных
Property Let startDir(value As String)
    LetProp "startDir", value
End Property
Property Get startDir() As String
    startDir = GetProp("startDir")
End Property
Property Let workDir(value As String)
    LetProp "workDir", value
End Property
Property Get workDir() As String
    workDir = GetProp("workDir")
End Property
Property Let cusDB(value As String)
    LetProp "cusDB", value
End Property
Property Get cusDB() As String
    cusDB = GetProp("cusDB")
End Property
Property Let measInstrDB(value As String)
    LetProp "measInstrDB", value
End Property
Property Get measInstrDB() As String
    measInstrDB = GetProp("measInstrDB")
End Property
Property Let etalDB(value As String)
    LetProp "etalDB", value
End Property
Property Get etalDB() As String
    etalDB = GetProp("etalDB")
End Property
Property Let empDB(value As String)
    LetProp "empDB", value
End Property
Property Get empDB() As String
    empDB = GetProp("empDB")
End Property
'#######################################################
'страница работы с данными поверки СИ
Property Let depPrefix(value As String)
    LetProp "depPrefix", value
End Property
Property Get depPrefix() As String
    depPrefix = GetProp("depPrefix")
End Property
Property Let labNum(value As String)
    LetProp "labNum", value
End Property
Property Get labNum() As String
    labNum = GetProp("labNum")
End Property
Property Let isFullName(value As String)
    LetProp "isFullName", value
End Property
Property Get isFullName() As String
    isFullName = GetProp("isFullName")
End Property

Property Let useArchiveDir(value As String)
    LetProp "useArchiveDir", value
End Property
Property Get useArchiveDir() As String
    useArchiveDir = GetProp("useArchiveDir")
End Property
Property Let ArchivePath(value As String)
    LetProp "ArchivePath", value
End Property
Property Get ArchivePath() As String
    ArchivePath = GetProp("ArchivePath")
End Property

Property Let empName(value As String)
    LetProp "empNAME", value
End Property
Property Get empName() As String
    empName = GetProp("empNAME")
End Property
Property Let empState(value As String)
    LetProp "empSTATE", value
End Property
Property Get empState() As String
    empState = GetProp("empSTATE")
End Property
Property Let headNAME(value As String)
    LetProp "headNAME", value
End Property
Property Get headNAME() As String
    headNAME = GetProp("headNAME")
End Property
Property Let headSTATE(value As String)
    LetProp "headSTATE", value
End Property
Property Get headSTATE() As String
    headSTATE = GetProp("headSTATE")
End Property
Property Let verNAME(value As String)
    LetProp "verNAME", value
End Property
Property Get verNAME() As String
    verNAME = GetProp("verNAME")
End Property

Property Let empSecName(value As String)
    LetProp "empSecNAME", value
End Property
Property Get empSecName() As String
    empSecName = GetProp("empSecNAME")
End Property
Property Let empSecState(value As String)
    LetProp "empSecSTATE", value
End Property
Property Get empSecState() As String
    empSecState = GetProp("empSecSTATE")
End Property

Property Let isPortable(value As String)
    LetProp "isPortable", value
End Property
Property Get isPortable() As String
    isPortable = GetProp("isPortable")
End Property

Property Let helpOption(value As String)
    LetProp "helpOption", value
End Property
Property Get helpOption() As String
    helpOption = GetProp("helpOption")
End Property

Property Let templatesDir(value As String)
    LetProp "templatesDir", value
End Property
Property Get templatesDir() As String
    templatesDir = GetProp("templatesDir")
End Property

'#######################################################
'страница работы с данными для экспорта
Property Let wdSvPath(value As String)
    LetProp "wdSvPath", value
End Property
Property Get wdSvPath() As String
    wdSvPath = GetProp("wdSvPath")
End Property
Property Let wdSrtPath(value As String)
    LetProp "wdSrtPath", value
End Property
Property Get wdSrtPath() As String
    wdSrtPath = GetProp("wdSrtPath")
End Property
Property Let wdInPath(value As String)
    LetProp "wdInPath", value
End Property
Property Get wdInPath() As String
    wdInPath = GetProp("wdInPath")
End Property
Property Let xlPrPath(value As String)
    LetProp "xlPrPath", value
End Property
Property Get xlPrPath() As String
    xlPrPath = GetProp("xlPrPath")
End Property
Property Let xlPrcPath(value As String)
    LetProp "xlPrcPath", value
End Property
Property Get xlPrcPath() As String
    xlPrcPath = GetProp("xlPrcPath")
End Property
Property Get myValues() As String()
    myValues = sArrVal
End Property
Property Get myParameters() As String()
    myParameters = sArrParam
End Property

'#######################################################
'функция заносит в класс значения с ключами
Function FillProperties(Parameters() As String, values() As String)
    sArrVal = values: sArrParam = Parameters
End Function
'#######################################################
'функция передаёт индекс значения по ключу
Private Function GetProp(keyProp As String) As String
    Dim i As Integer
    For i = LBound(sArrParam) To UBound(sArrParam)
        If sArrParam(i) Like "*" & keyProp & "*" Then GetProp = sArrVal(i): Exit For
    Next i
End Function
'#######################################################
'метод ищет ключ и передаёт по нему значение
Private Sub LetProp(keyProp As String, value As String)
    Dim i As Integer
    For i = LBound(sArrParam) To UBound(sArrParam)
        If sArrParam(i) Like "*" & keyProp & "*" Then sArrVal(i) = value: Exit For
    Next i
End Sub
'#######################################################
'функция формирует многомерный массив данных
Function FillDataBase(sArrFile() As String, Optional AllData As Boolean) As String()
    Dim i As Integer, j As Integer, sArrTemp() As String, sArrKeyC() As String
    If AllData Then 'импорт всех столбцов
        sArrTemp = Split(sArrFile(0), InStrDelimiter) 'разбить строку с данными о символах для поиска
        ReDim sArrKeyC(UBound(sArrTemp), UBound(sArrFile)) 'выделить память под массив
    Else 'заполнить массив кейкодов
        ReDim sArrKeyC(1, UBound(sArrFile)) 'выделить память под массив с двумя столбцами
        ' 0 - русскоязычная раскладка
        ' 1 - англоязычная раскладка
    End If
    For i = LBound(sArrFile) To UBound(sArrFile)
        sArrTemp = Split(sArrFile(i), InStrDelimiter) 'разбить строку с данными о символах для поиска
        For j = LBound(sArrKeyC()) To UBound(sArrKeyC())
            If AllData Then 'импорт всех столбцов
                sArrKeyC(j, i) = sArrTemp(j)
            Else 'заполнить массив кейкодов
                sArrKeyC(j, i) = sArrTemp(j + 2)
            End If
        Next j
    Next i
    FillDataBase = sArrKeyC
End Function
