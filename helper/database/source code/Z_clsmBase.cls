VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Z_clsmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'класс обеспечивает импорт данных из файла, получение значений по ключу
Option Explicit
Private sArrParameters() As String, sArrFile() As String, sArrValues() As String, sSavePath As String
'#######################################################
'метод срабатывает при загрузке класса
Private Sub Class_Initialize()
    ReDim sArrParameters(0)
End Sub
'#######################################################
'функция формирует массив входных параметров
Function AddP(sItem1 As String, Optional sItem2 As String, Optional sItem3 As String)
    'если массив полон, расширить его
    If sItem1 <> "" Then
        If sArrParameters(UBound(sArrParameters)) <> "" Then ReDim Preserve sArrParameters(UBound(sArrParameters) + 1)
        sArrParameters(UBound(sArrParameters)) = sItem1 'передать значение в массив
    End If
    If sItem2 <> "" Then
        If sArrParameters(UBound(sArrParameters)) <> "" Then ReDim Preserve sArrParameters(UBound(sArrParameters) + 1)
        sArrParameters(UBound(sArrParameters)) = sItem2 'передать значение в массив
    End If
    If sItem3 <> "" Then
        If sArrParameters(UBound(sArrParameters)) <> "" Then ReDim Preserve sArrParameters(UBound(sArrParameters) + 1)
        sArrParameters(UBound(sArrParameters)) = sItem3 'передать значение в массив
    End If
End Function
'#######################################################
'метод очищает массив параметров
Sub ClearParameters()
    ReDim sArrParameters(0)
End Sub
Property Get Parameters() As String()
    Parameters = sArrParameters()
End Property
Property Get values() As String()
    values = sArrValues()
End Property
'#######################################################
'функция возвращает массив файла по указанному пути, а также формирует внутренний массив файла
Function GetArrFF( _
    sPath As String, _
    Optional sFName As String _
    ) As String()

    Dim i As Integer, sArrTemp() As String, iFileInd As Integer
    If sFName <> "" Then 'если указано отдельное имя файла в каталоге
        sSavePath = sPath & sFName
        If FileExist(sPath & sFName) = False Then ReDim sArrFile(0): Exit Function
        
    Else 'указан полный путь
        sSavePath = sPath
        If FileExist(sPath) = False Then ReDim sArrFile(0): Exit Function
    End If
    
    iFileInd = FreeFile 'получить номер первого свободного файла
    
    Open sSavePath For Input As iFileInd
        sArrTemp = Split(Input$(LOF(iFileInd), #iFileInd), vbNewLine) 'поместить в массив содержимое всего файла
    Close
    
     'обратный проход по массиву с выделением последней заполненной строки
    For i = UBound(sArrTemp) To LBound(sArrTemp) Step -1
        If sArrTemp(i) <> "" Then Exit For 'определена верхняя граница заполненного массива
    Next i
    
    If i < 0 Then i = 0 'на случай, если массив был пустым
    ReDim Preserve sArrTemp(i) 'удалить из рабочего массива пустые элементы в конце
    sArrFile = sArrTemp: GetArrFF = sArrTemp
    
End Function
'#######################################################
'функция возвращает массив значений из массива-источника
Function FillValues()
    Dim i As Integer, j As Integer, bolNotExitCycle As Boolean, sArrTemp() As String
    
    ReDim sArrValues(UBound(sArrParameters))
    
    If sArrParameters(UBound(sArrParameters)) <> "" Then 'если получены переменные параметров
        For i = LBound(sArrFile) To UBound(sArrFile) 'пройтись по всему массиву
            For j = LBound(sArrParameters) To UBound(sArrParameters)
                
                'если строка массива содержит нужную метку, передать параметр в текущий массив
                If sArrFile(i) Like sArrParameters(j) & "=*" Then
                    If sArrValues(j) = "" Or sArrValues(j) = "недоступно" Then 'если значение не было получено ранее
                    
                        sArrTemp = Split(sArrFile(i), "=")
                        If UBound(sArrTemp) = 1 Then sArrValues(j) = sArrTemp(UBound(sArrTemp))
                    End If
                End If
            Next j
            'выйти из цикла просмотра всех строк массива настроек, если получены все значения
            For j = LBound(sArrParameters) To UBound(sArrParameters)
                If sArrValues(j) = "" Then bolNotExitCycle = True: Exit For
            Next j
            If bolNotExitCycle = False Then Exit For 'если все параметры переданы в исходный массив
        Next i
        'повторный проход для заполнения пустых значений
        For j = LBound(sArrParameters) To UBound(sArrParameters)
            If sArrValues(j) = "" Then sArrValues(j) = "недоступно"
        Next j
    End If
End Function
'#######################################################
'метод задаёт массив значений из исходного массива
Sub SaveProperties(Parameters() As String, values() As String)

    If Dir(setDir, vbDirectory) = "" Then MkDir setDir      'чтобы предотвратить ошибку при разработке
    
    Dim i As Integer, j As Integer, iFillCount As Integer, sTempStr As String, iFileInd As Integer
    Dim bolArrFill() As Boolean
    
    sArrParameters = Parameters: sArrValues = values
    
    iFillCount = UBound(sArrValues) + 1 'переменная заполнения всех данных
    ReDim bolArrFill(UBound(sArrValues)) 'массив передачи даных в файл
    
    For i = LBound(sArrFile) To UBound(sArrFile) 'пройтись по всему исходному массиву
        If iFillCount > 0 Then 'переменная учёта передачи всех значений
        
            For j = LBound(sArrParameters) To UBound(sArrParameters)
            
                'если строка массива содержит нужную метку, передать параметр в текущий массив
                If sArrFile(i) Like sArrParameters(j) & "=*" Then
                    If sArrValues(j) = "" Then sArrValues(j) = "недоступно"
                    sArrFile(i) = sArrParameters(j) & "=" & sArrValues(j) 'передать значение в массив
                    
                    iFillCount = iFillCount - 1: bolArrFill(j) = True: Exit For
                End If
            Next j
        End If
        sTempStr = sTempStr & sArrFile(i)
        If i < UBound(sArrFile) Then sTempStr = sTempStr & vbNewLine
    Next i
    
    If iFillCount > 0 Then 'переданы не все переменные
        For i = LBound(bolArrFill) To UBound(bolArrFill)
            If bolArrFill(i) = False Then 'данная переменная не была передана
                
                If sTempStr <> "" Then _
                    sTempStr = sTempStr & vbNewLine 'чтобы не было пустых строк при первом сохранении
                
                sTempStr = sTempStr & sArrParameters(i) & "=" & sArrValues(i)
                bolArrFill(i) = True: iFillCount = iFillCount - 1
            End If
            If iFillCount < 0 Then Exit For 'дозаполнены все параметры
        Next i
    End If
    
    iFileInd = FreeFile 'получить номер первого свободного файла
    
    Open sSavePath For Output As iFileInd 'открыть файл для записи
        Print #iFileInd, sTempStr
    Close
End Sub
