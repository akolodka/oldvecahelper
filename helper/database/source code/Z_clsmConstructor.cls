VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Z_clsmConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sArrValCfg() As String, sArrParamCfg() As String, sArrValSet() As String, sArrParamSet() As String
Property Get startDir() As String
    startDir = GetProp("startDir")
End Property
Property Get DbName() As String
    DbName = GetProp("DB")
End Property
Property Get constrSrch() As String
    constrSrch = GetProp("constrSrch") 'написание переменной "Search" недопустимо - в форме поиска перекрываются параметры
End Property
Property Get templatesDir() As String
    templatesDir = GetProp("templatesDir")
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
Private Function GetProp(keyProp As String) As String
    Dim i As Integer
    For i = LBound(sArrParamCfg) To UBound(sArrParamCfg)
        If sArrParamCfg(i) Like "*" & keyProp & "*" Then GetProp = sArrValCfg(i): Exit Function
    Next i
    For i = LBound(sArrParamSet) To UBound(sArrParamSet)
        If sArrParamSet(i) Like "*" & keyProp & "*" Then GetProp = sArrValSet(i): Exit Function
    Next i
End Function
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
