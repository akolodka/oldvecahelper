VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Z_clsmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sArrValCfg() As String, sArrParamCfg() As String, sArrValSet() As String, sArrParamSet() As String
Property Get startDir() As String
    startDir = GetProp("startDir")
End Property
Property Get workDir() As String
    workDir = GetProp("workDir")
End Property
Property Get DbName() As String
    DbName = GetProp("DB")
End Property
Property Get templatesDir() As String
    templatesDir = GetProp("templatesDir")
End Property


Property Get headNAME() As String 'для заполнения данных эталона при загрузке шаблона - не удалять
    headNAME = GetProp("headNAME")
End Property
Property Get headSTATE() As String
    headSTATE = GetProp("headSTATE")
End Property
Property Get verNAME() As String
    verNAME = GetProp("verNAME")
End Property
Property Get empName() As String
    empName = GetProp("empNAME")
End Property
Property Get empState() As String
    empState = GetProp("empSTATE")
End Property
Property Get empSecName() As String
    empSecName = GetProp("empSecNAME")
End Property
Property Get empSecState() As String
    empSecState = GetProp("empSecSTATE")
End Property
Property Get helpOption() As String
    helpOption = GetProp("helpOption")
End Property

Property Let LastSearch(value As String)
    LetProp "Search", value
End Property
Property Get LastSearch() As String
    LastSearch = GetProp("Search")
End Property

Property Let LastIndex(value As String)
    LetProp "Index", value
End Property
Property Get LastIndex() As String
    LastIndex = GetProp("Index")
End Property

Property Get isFullName() As String
    isFullName = GetProp("isFullName")
End Property

Property Let FullFirstName(value As String)
    LetProp "FullFirstName", value
End Property
Property Get FullFirstName() As String
    FullFirstName = GetProp("FullFirstName")
End Property

Property Let constrSrch(value As String)
    LetProp "constrSrch", value
End Property

Property Get depPrefix() As String
    depPrefix = GetProp("depPrefix")
End Property
Property Get labNum() As String
    labNum = GetProp("labNum")
End Property

Property Get useArchiveDir() As String
    useArchiveDir = GetProp("useArchiveDir")
End Property
Property Get ArchivePath() As String
    ArchivePath = GetProp("ArchivePath")
End Property

Property Let normalCondition(value As String)
    LetProp "normalCondition", value
End Property
Property Get normalCondition() As String
    normalCondition = GetProp("normalCondition")
End Property

Property Get myValues() As String()
    myValues = sArrValSet
End Property
Property Get myParameters() As String()
    myParameters = sArrParamSet
End Property
'#######################################################
'функция заносит в класс значения с ключами
Function FillConfiguration(Parameters() As String, values() As String)
    sArrValCfg = values: sArrParamCfg = Parameters
End Function
'#######################################################
'функция заносит в класс значения с ключами
Function FillSettings(Parameters() As String, values() As String)
    sArrValSet = values: sArrParamSet = Parameters
End Function
'#######################################################
'функция передаёт индекс значения по ключу
Private Function GetProp(keyProp As String, Optional bolSecChoose As Boolean) As String
    
    Dim i As Integer, bChoseCnt As Byte
    For i = LBound(sArrParamCfg) To UBound(sArrParamCfg)
        If sArrParamCfg(i) Like "*" & keyProp & "*" Then _
            bChoseCnt = bChoseCnt + 1: GetProp = sArrValCfg(i)
            
        If bolSecChoose And bChoseCnt = 2 Then Exit Function
        If bolSecChoose = False And bChoseCnt = 1 Then Exit Function
    Next i
    
    For i = LBound(sArrParamSet) To UBound(sArrParamSet)
        If sArrParamSet(i) Like "*" & keyProp & "*" Then _
            bChoseCnt = bChoseCnt + 1: GetProp = sArrValSet(i)
           
        If bolSecChoose And bChoseCnt = 2 Then Exit Function
        If bolSecChoose = False And bChoseCnt = 1 Then Exit Function
    Next i
End Function
'#######################################################
'метод ищет ключ и передаёт по нему значение
Private Sub LetProp(keyProp As String, value As String, Optional bolSecChoose As Boolean)

    Dim i As Integer, bChoseCnt As Byte
    For i = LBound(sArrParamSet) To UBound(sArrParamSet)
        If sArrParamSet(i) Like "*" & keyProp & "*" Then bChoseCnt = bChoseCnt + 1
        If bolSecChoose And bChoseCnt = 2 Then Exit For
        If bolSecChoose = False And bChoseCnt = 1 Then Exit For
    Next i
    
    If bChoseCnt > 0 Then sArrValSet(i) = value
End Sub
'#######################################################
'функция формирует многомерный массив данных
Function FillDataBase(sArrFile() As String, Optional AllData As Boolean) As String()
    Dim i As Integer, j As Integer, sArrTemp() As String, sArrKeyC() As String
    If AllData Then 'импорт всех столбцов
        sArrTemp = Split(sArrFile(0), InStrDelimiter) 'разбить строку с данными о символах для поиска
        ReDim sArrKeyC(UBound(sArrTemp), UBound(sArrFile)) 'выделить память под массив
        
        If UMenu.typе = instrumentsOLD Then ReDim sArrKeyC(UBound(sArrTemp) + 1, UBound(sArrFile)) 'выделить память под массив
    Else 'заполнить массив кейкодов
        ReDim sArrKeyC(1, UBound(sArrFile)) 'выделить память под массив с двумя столбцами
        ' 0 - русскоязычная раскладка
        ' 1 - англоязычная раскладка
    End If
    For i = LBound(sArrFile) To UBound(sArrFile)
        sArrTemp = Split(sArrFile(i), InStrDelimiter) 'разбить строку с данными о символах для поиска
        For j = LBound(sArrTemp()) To UBound(sArrTemp())
            If AllData Then 'импорт всех столбцов
                sArrKeyC(j, i) = sArrTemp(j)
            Else 'заполнить массив кейкодов
                sArrKeyC(LBound(sArrKeyC), i) = sArrTemp(UBound(sArrTemp) - 1)
                sArrKeyC(UBound(sArrKeyC), i) = sArrTemp(UBound(sArrTemp))
                Exit For
            End If
        Next j
    Next i
    FillDataBase = sArrKeyC
End Function
