Attribute VB_Name = "z_Old_CommonProject_Code"
Option Explicit 'запрет на использование неявных переменных

#If VBA7 Then
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
#Else
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
#End If

Const SW_RESTORE = 9
Const DEFAULT_VALUE = "nodata"
Const REFERENCE_FIFNUM_FILENAME = "fifRegNum.ref"

Public myMenu As frmLoadType, myMultiSel As MyMultiSelectProperties
Public bolUF_Set_Load As Boolean, bolUF_Cnstr_Load As Boolean, isWorkFormLoaded As Boolean
Public bolBackSpace As Boolean, bolAlreadySaved As Boolean

Type MyMultiSelectProperties 'группа переменных, описывающая данные мультивыбора
    sMSBaseDir As String: sMSfifNum As String: sMSAdditionalText As String: sMSFullFileName As String: bolThisWBSheetLoad As Boolean: sMSType As String
    templates As String
End Type

Type HeadVerName 'группа переменных, описывающая данные руководителей в выдаваемом документе
    sHeadState As String: sHeadName As String: sVerName As String: sVerSecName As String
End Type

Type EmpDivision 'группа переменных, описывающая номер отдела и лаборатории
    sDepPref As String: sLabNum As String
End Type

Type MeasInstrument 'группа переменных, описывающая 4 параметра средства измерений
    sName As String: sType As String: sFif As String: sMetodic As String: sModification As String: sRef As String
    bolEtal As Boolean
End Type

Type frmLoadType 'тип загружаемого меню
    type As Byte
End Type

Type StartWorkDirs 'начальная и рабочая директория
    sStartDir As String
    sWorkDir As String
    templates As String
    protocolBaseDir As String
End Type

Type myPrintShift
    iShiftX As Integer: iShiftY As Integer
End Type


'##########################################################################
'функция возвращает каталог настроек по умолчанию
Function setDir()
    setDir = Environ("APPDATA") & "\Microsoft\Помощник ПКР\"
End Function

Function configDirNew( _
    ) As String
    
    configDirNew = Environ("APPDATA") & "\Microsoft\Помощник ПКР\"
    
    If Dir(configDirNew, vbDirectory) = vbNullString Then _
        MkDir configDirNew
    
End Function

'##########################################################################
'разделитель внутри строки в загружаемых файлах
Function InStrDelimiter()
    InStrDelimiter = Chr(9)
End Function
'##########################################################################
'границы для корректирующих коэффициентов
Function BoundShiftX()
    BoundShiftX = 5
End Function
Function BoundShiftY()
    BoundShiftY = 7
End Function
'##########################################################################
'##########################################################################
'функция проверяет на существование файл, расположенный по указанному пути
Function FileExist(fPath As String, Optional fName As String) As Boolean
    FileExist = False

    If fName = "недоступно" Then Exit Function
    If fName <> "" Then 'передан отдельно путь и отдельно имя файла
        If Right(fPath, 1) <> "\" Then fPath = fPath & "\" 'на случай, если в конце пути нет слеша
        If Dir(fPath & fName) <> "" Then FileExist = True
    Else 'передан полный путь файла
        If fPath <> "недоступно" Then _
            If Dir(fPath) <> "" Then FileExist = True
    End If
End Function
'##########################################################################
'функция проверяет на существование файл, расположенный по указанному пути
Function FolderNotExist( _
    folderPath As String _
    ) As Boolean
    
    FolderNotExist = True
    
    If Dir(folderPath, vbDirectory) <> vbNullString Then _
        FolderNotExist = False
End Function
'##########################################################################
'функция предоставляет меню выбора файла
Function GetFileFPath( _
    ControlIndex As Byte, _
    sBaseDir As String, _
    Optional sTitle As String _
    ) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .Filters.Clear
        .InitialView = msoFileDialogViewDetails
        
        .Title = "Выбор источника данных"
        
        If sTitle <> vbNullString Then _
            .Title = sTitle
        
        .AllowMultiSelect = False
        .InitialFileName = sBaseDir
        
        Select Case ControlIndex
            
            Case 1, 2 'выбрать файл сведений заказчиков
                .Filters.Add "Сведения заказчиков", "*.cuDb; *.o13Db", 1
            
            Case 3, 4 'выбрать файл сведений СИ
                .Filters.Add "Сведения о средствах измерений", "*.miDb", 1
            
            Case 5, 6 'выбрать файл сведений эталонов
                .Filters.Add "Сведения об эталонах", "*.etDb", 1
            
            Case 7, 8 'выбрать файл сведений фамилий
                .Filters.Add "Сведения о фамилиях и должностях", "*.nmDb", 1
            
            Case 9 'конфигурация программы
                .Filters.Add "Файлы конфигураций", "*.uCfg", 1
            
            Case 10 'документ Word
                .Filters.Add "Документы Word", "*.doc*", 1
            
            Case 11 'шаблон протокола xl
                .Filters.Add "Книги Excel", "*.xls*", 1
        
        End Select
        
        If .Show = False Then _
            GetFileFPath = "NoPath": _
            Exit Function
            
        GetFileFPath = .SelectedItems(1) 'полный путь к файлу
    End With
    
End Function
'##########################################################################
'функция предоставляет меню выбора каталога
Function GetFolderFPath(Optional sTitle As String, Optional StartFolder As String, Optional Somnium As Boolean)
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Filters.Clear: .InitialView = msoFileDialogViewDetails
        .Title = "Выбор директории назначения"
        If sTitle <> "" Then .Title = sTitle
        
        .InitialFileName = StartFolder
        If StartFolder = "" Or StartFolder = "недоступно" Then
            .InitialFileName = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
            If Somnium Then .InitialFileName = "\\somnium\irlab\Документы\"
        End If

        .AllowMultiSelect = False
        If .Show = 0 Then GetFolderFPath = "NoPath": Exit Function
        
        GetFolderFPath = .SelectedItems(1) 'путь к каталогу
    End With
End Function
'##########################################################################
'процедура выполняет сортировку двумерного массива методом вставки
Sub SortMassBiv(ByRef arrToSort() As String, _
    Optional arrSortLevel As Integer = 0, Optional SortByNumeric As Boolean = False, Optional SortByDecrease = True)
    
    'по умолчанию сортировка идёт по первому уровню, текстовым значениям и по убыванию
    Dim K As Integer, i As Integer, j As Integer, iArrLEdge As Integer, iArrUEdge As Integer, iStep As Integer, jP As Integer
    ReDim arrTempSort(UBound(arrToSort())) 'создать дополнительный массив
    
    iArrLEdge = LBound(arrToSort, 2) + 1 'нижняя граница второго измерения массива
    iArrUEdge = UBound(arrToSort, 2) 'верхняя граница второго измерения массива
    iStep = 1 'прямой шаг
    
    If SortByDecrease = False Then  ' если выбрана опция сортировки по возрастанию
        iArrLEdge = UBound(arrToSort, 2) - 1: iArrUEdge = LBound(arrToSort, 2): iStep = -1 'обратный шаг
    End If
    
    For i = iArrLEdge To iArrUEdge Step iStep 'пройтись по всем элементам массива с шагом iStep
        j = i
        Do
            If j = iArrLEdge - iStep Then Exit Do 'если рассматривается самый первый элемент - выйти из цикла
            
            jP = j - 1 'по умолчанию сортировка по убыванию
            If SortByDecrease = False Then jP = j + 1
            
            If SortByNumeric = True Then 'сортировка идёт по числам
            
                If arrToSort(LBound(arrToSort), jP) <> 0 Then 'если предыдущий элемент не пустой
                    If arrToSort(LBound(arrToSort), j) = 0 Then Exit Do 'если текущий элемент - нулевой элемент
                    If CDbl(arrToSort(arrSortLevel, jP)) < CDbl(arrToSort(arrSortLevel, j)) Then Exit Do 'если предыдущий элемент меньше текущего
                    
                    If CDbl(arrToSort(arrSortLevel, jP)) = CDbl(arrToSort(arrSortLevel, j)) Then 'одинаковые наименования
                        If arrSortLevel + 1 <= UBound(arrToSort) Then 'существует такое измерение массива
                            If CDbl(arrToSort(arrSortLevel + 1, jP)) < CDbl(arrToSort(arrSortLevel + 1, j)) Then Exit Do
                        End If
                    End If
                End If
            Else 'сортировка по тексту
                
                If arrToSort(LBound(arrToSort), jP) <> "" Then 'если предыдущий элемент не пустой
                    If arrToSort(LBound(arrToSort), j) = "" Then Exit Do 'если текущий элемент - пустой элемент
                    If Format(arrToSort(arrSortLevel, jP), ">") < Format(arrToSort(arrSortLevel, j), ">") Then Exit Do 'если предыдущий элемент меньше текущего
                    
                    If LCase(arrToSort(arrSortLevel, jP)) = LCase(arrToSort(arrSortLevel, j)) Then 'одинаковые наименования
                        If arrSortLevel + 1 <= UBound(arrToSort) Then 'существует такое измерение массива
                            If LCase(arrToSort(arrSortLevel + 1, jP)) < LCase(arrToSort(arrSortLevel + 1, j)) Then Exit Do
                        End If
                    End If
                End If
            End If
            
            For K = LBound(arrToSort) To UBound(arrToSort) 'первое измерение сортируемого массива
                arrTempSort(K) = arrToSort(K, j) 'временный массив
                arrToSort(K, j) = arrToSort(K, jP): arrToSort(K, jP) = arrTempSort(K)
            Next K
            
            j = jP
        Loop
    Next i
End Sub
'##########################################################################
'сортировка одномерного массива методом вставки
Sub SortMassOne(arrToSort() As String, Optional SortByIncrease As Boolean, Optional PrimaryString As Boolean)
    'по умолчанию сортировка идёт по убыванию
    
    Dim iArrLEdge As Integer, iArrUEdge As Integer, iStep As Integer
    iArrLEdge = LBound(arrToSort) + 1 'нижняя граница второго измерения массива
    iArrUEdge = UBound(arrToSort) 'верхняя граница второго измерения массива
    iStep = 1 'прямой шаг
    
    If SortByIncrease Then 'сортировка по возрастанию
        iArrLEdge = UBound(arrToSort) - 1: iArrUEdge = LBound(arrToSort): iStep = -1 'обратный шаг
    End If


    Dim sTemp As String, i As Integer, j As Integer, jP As Integer
    
    For i = iArrLEdge To iArrUEdge Step iStep 'сортировка идёт не с нулевого, а с первого элемента
        j = i
        
        Do
            If j = iArrLEdge - iStep Then Exit Do 'если рассматривается самый первый элемент
            
            jP = j - 1 'по умолчанию сортировка по убыванию
            If SortByIncrease Then jP = j + 1 'сортировка по возрастанию
            If arrToSort(j) = "" Then Exit Do 'если текущий элемент - пустой элемент
            
            If arrToSort(jP) <> "" Then 'если предыдущий элемент непустой
            
                If PrimaryString And InStr(arrToSort(j), "ГЭТ") > 0 Then 'если в предыдущем есть ГЭТ
                    If InStr(arrToSort(jP), "ГЭТ") > 0 And _
                        LCase(arrToSort(jP)) < LCase(arrToSort(j)) Then Exit Do 'если в обоих элементах есть ГЭТ
                        
                    If InStr(arrToSort(jP), "ГЭТ") > 0 Then Exit Do    'если в текущем нет ГЭТ - выйти из цикла
                Else 'обычная сортировка
                    If LCase(arrToSort(jP)) < LCase(arrToSort(j)) Then Exit Do   'если предыдущий элемент меньше текущего
                End If
            End If
            
            sTemp = arrToSort(j)  'значение временного элемента
            arrToSort(j) = arrToSort(jP)
            arrToSort(jP) = sTemp
            j = jP
        Loop
    Next i
End Sub

'##########################################################################
'функция удаляет нечитаемые символы в начале и конце строки
Function DeleteSpaceStEnd(ByRef stringValue As String, Optional noDataForNull As Boolean) As String

    Dim i As Integer, sSym As String
    If stringValue <> "" Then
    
        Select Case Len(stringValue) 'количество символов в исходной строке
            Case 1
                If Asc(stringValue) <= 32 Then stringValue = "" 'если был передан один из нечитаемых символов
            Case Else
                sSym = ""
                Do '1) проход слева
                    sSym = Left(stringValue, 1)
                    If Asc(sSym) > 32 Then Exit Do
                    If Asc(sSym) <= 32 Then _
                        stringValue = Right(stringValue, Len(stringValue) - 1)
                    If Len(stringValue) < 1 Then Exit Do
                Loop
                
                sSym = ""
                Do '1) проход справа
                    If stringValue = "" Then Exit Do
                    sSym = Right(stringValue, 1)
                    If Asc(sSym) > 32 Then Exit Do
                    If Asc(sSym) <= 32 Then _
                        stringValue = Left(stringValue, Len(stringValue) - 1)
                    If Len(stringValue) < 1 Then Exit Do
                Loop
        End Select
        
    End If
    
    If noDataForNull Then
        If stringValue = "" Then stringValue = DEFAULT_VALUE
    End If
    
    DeleteSpaceStEnd = stringValue 'передать обработанное значение
End Function
'##########################################################################
'функция возвращает текстовую строку нужной длины в зависимости от установленной границы
Function ShortedString( _
    strData As String, _
    maxLength As Byte _
    ) As String
    
    ShortedString = strData ' по умолчанию возвращать всю строку 'название пути полностью умещается
    
    Dim leftPart As String, _
        rightPart As String, _
        iStrStart As Integer
        
    If Len(strData) > maxLength Then 'пусть не умещается полностью
        
        iStrStart = InStr(strData, "\")
        If iStrStart = 0 Then _
            Exit Function
        
        If iStrStart + 2 <= Len(strData) Then _
            leftPart = Left(strData, InStr(iStrStart + 2, strData, "\")) 'обязательная левая часть строки
        
        If maxLength - Len(leftPart) - 5 >= 0 Then _
            rightPart = Right(strData, maxLength - Len(leftPart) - 5) ' чтобы не было ошибки в случае короткой строки

        ShortedString = leftPart & " ... " & rightPart
    End If
End Function


'#######################################################
'функция предоставляет меню сохранения файла с именем пользователя
Function GetSaveAsFname( _
    DataBaseType As Byte, _
    Optional sBasePath As String _
    )
    
    Dim sStrAddition As String, sStrExt As String, sStrCaption As String
    Select Case DataBaseType
        Case 1, 2 'бд заказчиков
            sStrAddition = "_customers": sStrExt = ".cuDb": sStrCaption = "данных заказчиков:"
        Case 3, 4 'бд средств измерений
            sStrAddition = "_instruments": sStrExt = ".miDb": sStrCaption = "данных средств измерений:"
        Case 5, 6 'бд эталонов
            sStrAddition = "_etalons": sStrExt = ".etDb": sStrCaption = "данных эталонов:"
        Case 7, 8 'фамилии
            sStrAddition = "_employees": sStrExt = ".nmDb": sStrCaption = "данных фамилий и должностей:"
        Case 9 'экспорт конфигурации
            sStrAddition = "_pkrconfig": sStrExt = ".uCfg": sStrCaption = "конфигурации программы:"
    End Select
    
    GetSaveAsFname = InputBox("Введите имя для файла " & sStrCaption, "Сохранение нового файла базы данных", Environ("USERNAME") & sStrAddition)
    If GetSaveAsFname <> "" Then
        If sBasePath = "" Then
            GetSaveAsFname = setDir & GetSaveAsFname & sStrExt: Exit Function
        Else
            GetSaveAsFname = sBasePath & GetSaveAsFname & sStrExt: Exit Function
        End If
    End If
    GetSaveAsFname = "NoPath" 'по умолчанию
End Function
'#######################################################
'функция отделяет имя файла базы данных от пути
Function TrueName(fPath As String)
    If InStr(fPath, "\") = 0 Then Exit Function
    TrueName = Right(fPath, InStr(2, StrReverse(fPath), "\") - 1) 'имя без слеша
End Function
'############################################################
'функция корректирует падеж при выводе инфомации о количестве источников
Function RusPadejPozition(value As Integer)
    Select Case Right(value, 1)
        Case 1
            RusPadejPozition = value & " позиция.": If Len(CStr(value)) >= 2 Then _
                If Right(CStr(value), 2) = 11 Then RusPadejPozition = value & " позиций."
                
        Case 2 To 4
            RusPadejPozition = value & " позиции.": If Len(CStr(value)) >= 2 Then _
                If Right(CStr(value), 2) >= 12 And Right(CStr(value), 2) <= 14 Then RusPadejPozition = value & " позиций."
                
        Case Else
            RusPadejPozition = value & " позиций."
    End Select
End Function
'##########################################################################
'функция корректирует падеж при выводе информации о количестве результатов поиска
Function RusPadejCoincidence(number As Integer, objLabel As Object)
    objLabel.foreColor = &H80000012 ' - чёрный
    
    Select Case number
        Case 0
            RusPadejCoincidence = "нет совпадений"
        Case 1
            RusPadejCoincidence = "нажмите Enter"
            TrueElementForeColor objLabel
        Case 21, 31, 41, 51, 61, 71
            RusPadejCoincidence = "найдена " & number & " позиция"
        Case 2 To 4, 22 To 24, 32 To 34, 42 To 44, 52 To 54
            RusPadejCoincidence = "найдено " & number & " позиции"
        Case Else
            RusPadejCoincidence = "найдено " & number & " позиций"
    End Select
End Function
'#########################################################
'процедура заполняет сведения о существовании шаблонов
Sub CheckTempDir( _
    sStartDir As String, _
    ByRef sArrData() As String _
    )
    
    SortMassBiv sArrData, 1 'сортировать массив по типу СИ
    
    Dim i As Byte, sTempPath As String, sTempDir As String, sFifNum As String, sTempStr As String, sTypeName As String
    Dim bolMp As Boolean, bolOT As Boolean, bolTMP As Boolean, referenceRegFifNum As String, bolRef As Boolean, sRefDir As String, sRefPath As String
    
    For i = LBound(sArrData, 2) To UBound(sArrData, 2) 'для каждого блока
    
        bolMp = False
        bolOT = False
        bolTMP = False
        bolRef = False
        
        sTempStr = vbNullString
        
        sFifNum = sArrData(LBound(sArrData) + 2, i)
        sTypeName = sArrData(LBound(sArrData) + 1, i)
        sArrData(UBound(sArrData), i) = DEFAULT_VALUE 'по умолчанию каждый элемент недоступен
        referenceRegFifNum = sArrData(LBound(sArrData) + 4, i) 'перекрёстная ссылка
        
        sTempDir = Dir(sStartDir & "*" & sFifNum & "*", vbDirectory) 'каталог СИ
        sRefDir = Dir(sStartDir & "*" & referenceRegFifNum & "*", vbDirectory) 'каталог перекрёстной ссылки
        
        If sTempDir <> vbNullString Then 'каталог с номером в фиф обнаружен
            
            sTempPath = sStartDir & sTempDir  'каталог с номером в ФИФ
            sRefPath = sStartDir & sRefDir 'каталог перекрёстной ссылки
            
            If Dir(sTempPath & "\mp_" & sFifNum & "*") <> vbNullString Then bolMp = True    'наличие методики
            If Dir(sTempPath & "\ot_" & sFifNum & "*") <> vbNullString Then bolOT = True    'наличие описания типа
            
            If Dir(sTempPath & "\pr" & "*" & sFifNum & "*.xls*") <> vbNullString Or _
                    Dir(sTempPath & "\body" & "*" & sFifNum & "*.xls*") <> vbNullString Then bolTMP = True 'наличие шаблонов
            
            If Dir(sRefPath & "\pr" & "*" & referenceRegFifNum & "*.xls*") <> vbNullString Or _
                    Dir(sRefPath & "\body" & "*" & referenceRegFifNum & "*.xls*") <> vbNullString Then bolRef = True 'наличие перекрёстного шаблона

            If bolMp Then sTempStr = "МП" 'если опознано наличие методики поверки
            
            If bolOT Then 'если опознано наличие описания типа
                If sTempStr <> vbNullString Then sTempStr = sTempStr & "+"
                sTempStr = sTempStr & "ОТ"
            End If
            
            If bolRef Then 'опознано наличие перекрёстной ссылки - приоритет загрузки
                If sTempStr <> vbNullString Then sTempStr = sTempStr & "+"
                sTempStr = sTempStr & "ШБ*"
            Else
                If bolTMP Then 'опознано наличие шаблона
                    If sTempStr <> vbNullString Then sTempStr = sTempStr & "+"
                    sTempStr = sTempStr & "ШБ"
                End If
            End If

            If sTempStr <> vbNullString Then sArrData(UBound(sArrData), i) = sTempStr
        End If
    Next
End Sub
'#########################################################
'процедура уменьшает размерность одномерного массива
Sub ReduceArrOne(ByRef sArrToReduce() As String)
    Do While sArrToReduce(UBound(sArrToReduce)) = ""
        If UBound(sArrToReduce) = 0 Then Exit Do
        ReDim Preserve sArrToReduce(UBound(sArrToReduce) - 1)
    Loop
End Sub
'#########################################################
'функция устраняет в строке запрещённые в имени файла символы
Function ReplaceBadSymbols(sTempStr As String) As String
    
    sTempStr = Replace(sTempStr, "\", "_"): sTempStr = Replace(sTempStr, "/", "_")
    sTempStr = Replace(sTempStr, ":", "_"): sTempStr = Replace(sTempStr, "*", "_")
    sTempStr = Replace(sTempStr, "?", "_"): sTempStr = Replace(sTempStr, "<", "_")
    sTempStr = Replace(sTempStr, ">", "_"): sTempStr = Replace(sTempStr, "|", "_")
    sTempStr = Replace(sTempStr, """", "_")
    
    ReplaceBadSymbols = sTempStr
End Function
'#########################################################
'процедура убирает повторяющиеся значения в массиве
Sub ReplaceRepeateInArrOne(ByRef sArray() As String, Optional sReplaceText As String = DEFAULT_VALUE, Optional CompareWoExt As Boolean)
    
    Dim i As Integer, j As Integer, bRepeateCnt As Byte, sBaseStr As String, sCompareStr As String
    For i = LBound(sArray) To UBound(sArray)
        bRepeateCnt = 0
        
        For j = LBound(sArray) To UBound(sArray)
            If sArray(i) = "" Then Exit For
            
            sBaseStr = sArray(i): sCompareStr = sArray(j)
            If CompareWoExt Then sBaseStr = GetFileNameWithOutExt(sArray(i)): sCompareStr = GetFileNameWithOutExt(sArray(j))
            
            If sBaseStr = sCompareStr Then bRepeateCnt = bRepeateCnt + 1
        Next j
        
        If bRepeateCnt > 1 Then sArray(i) = sReplaceText
    Next i
End Sub
'#########################################################
'процедура заполняет двумерный массив, содержащий найденные в другом массиве значения
Sub FindInBivArr( _
    sArrWhereToFind() As String, _
    ByRef sArrWhereAddResults() As String, _
    sStrToSearch As String, _
    Optional sStrExc As String = DEFAULT_VALUE _
    )
        
    If sArrWhereToFind(LBound(sArrWhereToFind), UBound(sArrWhereToFind, 2)) = "" Then _
        Exit Sub 'массив не был получен
    
    Dim sArrSplit() As String, _
        sArrTemp() As String 'временные поисковые массивы
        
    sArrSplit = Split(sStrToSearch, " ")
    ReplaceRepeateInArrOne sArrSplit
    SortMassOne sArrSplit
    
    ReduceArrOne sArrSplit
    ReDim sArrTemp(UBound(sArrSplit), 1)
    
    Dim i As Integer, _
        j As Integer, _
        K As Integer, _
        bolIsFinded As Boolean, _
        iCoincedenceCnt As Integer, _
        bolExitFor As Boolean, _
        b As Integer

    For i = LBound(sArrWhereToFind, 2) To UBound(sArrWhereToFind, 2) 'пройтись по второму измерению поискового массива
        
        For K = LBound(sArrTemp) To UBound(sArrTemp) 'пройтись по массиву поисковых значений
            
            sArrTemp(K, 0) = sArrSplit(K)
            sArrTemp(K, 1) = False 'передать поисковые значения
            
        Next K
        
'        Dim bUbound As Byte
'        bUbound = UBound(sArrWhereToFind): If UMenu.typе = instrumentsOLD Then bUbound = LBound(sArrWhereToFind) + 3
        
        For j = LBound(sArrWhereToFind) To UBound(sArrWhereToFind) 'пройтись по первому измерению поискового массива
            
            For K = LBound(sArrTemp) To UBound(sArrTemp) 'пройтись по массиву поисковых значений
                If LCase(sArrTemp(K, 0)) <> sStrExc Then _
                    If InStr(LCase(sArrWhereToFind(j, i)), LCase(sArrTemp(K, 0))) > 0 Then sArrTemp(K, 1) = True 'найдено совпадение при поиске
            Next K
            
        Next j 'в результате будет опознано хотя бы одно совпадение по обоим элементам

        bolIsFinded = False
        iCoincedenceCnt = 0
        
        For K = LBound(sArrTemp) To UBound(sArrTemp) 'пройтись по массиву поисковых значений
            
            If sArrTemp(K, 1) Then _
                iCoincedenceCnt = iCoincedenceCnt + 1
                
            If iCoincedenceCnt = (UBound(sArrTemp) + 1) / 2 Then _
                bolIsFinded = True: _
                Exit For 'когда найдено минимальное количество совпадений
            
        Next K

        If bolIsFinded Then  'опознано полное совпадение поисковой строки
            
            For j = LBound(sArrWhereToFind) To UBound(sArrWhereToFind) 'пройтись по первому измерению поискового массива
                If sArrWhereToFind(j, i) = sStrExc Then Exit For
                
                For K = LBound(sArrWhereAddResults, 2) To UBound(sArrWhereAddResults, 2) 'бегло пройтись по массиву, в который добавляются данные
                
                    If sArrWhereToFind(j, i) = sArrWhereAddResults(j, K) Then 'опознано хоть одно совпадение по БД ранее
                        iCoincedenceCnt = 0
                        For b = LBound(sArrWhereAddResults) To UBound(sArrWhereAddResults) 'пройтись по всему измерению строки
                            If sArrWhereToFind(b, i) = sArrWhereAddResults(b, K) Then iCoincedenceCnt = iCoincedenceCnt + 1
                        Next b
                        If iCoincedenceCnt = UBound(sArrWhereAddResults) + 1 Then bolExitFor = True: Exit For 'все элементы присутствуют
                    End If
                Next K
                
                If bolExitFor Then bolExitFor = False: Exit For ' опознано полное совпадение строки

                If sArrWhereAddResults(LBound(sArrWhereAddResults), UBound(sArrWhereAddResults, 2)) <> "" Then _
                    ReDim Preserve sArrWhereAddResults(UBound(sArrWhereAddResults), UBound(sArrWhereAddResults, 2) + 1) 'если временный массив заполнен, расширить его
                
                For K = LBound(sArrWhereAddResults) To UBound(sArrWhereAddResults) 'пройтись по всему измерению дополнительного массива
                    sArrWhereAddResults(K, UBound(sArrWhereAddResults, 2)) = sArrWhereToFind(K, i) 'передать все параметры найденной строки во временный массив
                Next K
                Exit For 'выйти из цикла j
            Next j
        End If
    Next i
End Sub
'#########################################################
'заменить кавычки при вставке текста
Sub ReplaceInvCommas(ByRef sTempStr As String)
    
    Dim i As Integer
    For i = 0 To Len(sTempStr) - 1

        If Asc(Mid(sTempStr, i + 1, 1)) = 34 Then 'опознана кавычка
            Select Case i
                Case 0: sTempStr = Chr(171) & Right(sTempStr, Len(sTempStr) - 1) 'открывающаяся кавычка в начале
                Case Is < Len(sTempStr) - 1
                
                    If Asc(Mid(sTempStr, i, 1)) <= 32 Then
                        sTempStr = Left(sTempStr, i) & Chr(171) & Right(sTempStr, Len(sTempStr) - i - 1)
                    Else
                        sTempStr = Left(sTempStr, i) & Chr(187) & Right(sTempStr, Len(sTempStr) - i - 1)
                    End If
                    
                Case Len(sTempStr) - 1
                    sTempStr = Left(sTempStr, Len(sTempStr) - 1) & Chr(187) 'закрывающаяся в конце
            End Select
        End If
    Next i
End Sub
