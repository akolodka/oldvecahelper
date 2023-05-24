Attribute VB_Name = "z_old_LoadXLTemplate"
Option Explicit

Const REFERENCE_FIFNUM_FILENAME = "fifRegNum.ref"
Const DEFAULT_VALUE = "nodata"

Const FirstVerifyString = "первичная" 'составляющая имени файла xl о первичной поверке

Const FindCellInStrLimX = 10 'максимальное смещение при поиске ячеек ОХ
Const FindCellInStrLimY = 50 'максимальное смещение при поиске ячеек ОY
Const iHeadProrCnt = 23 'ячеек в голове шапки

Const DEFAULT_FILE_PREFIX = "prm_"

Const EtalonNameMaxLength = 75 'максимальная длина символов для первой строки эталонов

Private Type myCustomer
    sCustName As String: sCustINN As String: sCustAdress As String
End Type

Private Type EmpNameState
    empName As String: empState As String
End Type

Dim fso As New FileSystemObject
'###################################################################
'Процедура собирает шаблон протокола в рабочую директорию
Function LoadXlTemplate( _
    ByRef objZ_clsmSearch As Z_clsmSearch, _
    myMi As MeasInstrument) _
    As Boolean
    'todo:[-] LoadXlTemplate -- проработать загрузку шаблона из каталога новой архитектуры
    
    Dim sBaseDir As String, _
        sWorkDir As String, _
        meEmp As EmpDivision, _
        bolLoadHelp As Boolean, _
        templatesDir As String
    

    sBaseDir = Config.sourceDataPath
    sWorkDir = Config.sandboxPath
    templatesDir = Config.instrumentsPath

    
    Dim sFileName As String, _
        sShortFileName As String
        
    sFileName = TemplateFileName( _
                                dbinstrumentsPath:=templatesDir, _
                                myMi:=myMi, _
                                fileExt:="*.xls*") ' получить полное имя файла шаблона (как старого так и нового)
    
    If sFileName = vbNullString Then _
        Handler.Notify "Файл шаблона не обнаружен, выполнение прервано"
    
    If sFileName <> vbNullString Then 'файл был выбран
        
        If InStr(GetBuiltInProperty("Keywords"), myMi.sFif) > 0 Then  'опознана загрузка нового протокола того же СИ
            
            If MsgBox("Загрузить шаблон в текущую книгу Excel?", vbYesNo) = vbYes Then _
                myMultiSel.bolThisWBSheetLoad = True
            
        End If
        
        Application.ScreenUpdating = False
'###################################################################
        GetMeWorkFile sFileName, templatesDir, sWorkDir, meEmp, myMi 'распределительная операция загрузки
'###################################################################
        
        FillEtalonsAfterImport 'заполнить сведения об эталонном оборудовании и вставить фамилию исполнителя
        FillNormalCondition objZ_clsmSearch 'заполнить сведения о нормальных условиях поверки
        'If bolLoadHelp Then ActiveWorkbook.ActiveSheet.[l1] = "Подсказка " & LoadHelpString(sBaseDir) & "." 'заполнить сведения подсказки
        
        Check_Footer 'убрать лишние колонтитулы на первой странице
        
        ActiveWorkbook.Save
        
        LoadXlTemplate = True
        Application.ScreenUpdating = True
    End If
End Function
    Private Sub Check_Footer()
        
        With ActiveWorkbook.ActiveSheet
            
            If CBool(.HPageBreaks.count) Then 'количество разрывов страниц > 1
                .PageSetup.FirstPage.LeftFooter.text = vbNullString
                .PageSetup.FirstPage.RightFooter.text = vbNullString
                    
            End If
        End With
        
    End Sub


'###################################################################
'Процедура помещает базовые сведения в шаблон протокола перед загрузкой тела
Sub InsertBaseMIdata(objBaseWorkbook As Object, meEmp As EmpDivision, _
    myMi As MeasInstrument, bolFirstVer As Boolean, Optional bolDontChangeBuiltPror As Boolean)
          
    Dim rTempCell As Range
    If FindCellRight("НИО", rTempCell, , True) Then
        
        If Not rTempCell.address = "$A$2" Then
            rTempCell = Replace(rTempCell, "#НИО#", meEmp.sDepPref): rTempCell = Replace(rTempCell, "#НИЛ#", meEmp.sLabNum)
        
            Dim bStCh As Byte, sLenCh As Byte
            bStCh = InStr(CStr(rTempCell), "Московский")
            If bStCh = 0 Then bStCh = InStr(CStr(rTempCell), "Отдел")
                
            sLenCh = Len(rTempCell) - InStr(rTempCell, "Московский") + 1
            rTempCell.Characters(bStCh, sLenCh).Font.Size = 8
        End If
    End If
    
    If myMultiSel.bolThisWBSheetLoad = False Then
    
        If FindCellRight("Наименование,", rTempCell) Then
            
            If rTempCell = vbNullString Then
            
'                Dim sTempFif As String
'                sTempFif = Replace(myMi.sFif, "-", vbNullString) 'получить номер в фиф без тире
'
'                If IsNumeric(sTempFif) Then

                rTempCell = myMi.sName & " " & myMi.sType 'по умолчанию "Наименование Тип"
                
            End If
        End If
    End If
    
    If bolDontChangeBuiltPror = False Then 'не изменять свойство документа
    
'        With objBaseWorkbook
'            Dim sTempStr As String, sWBName As String
'            sWBName = .Name: sTempStr = Replace(myMi.sFif, "-", "") 'проверка номера фиф на поверку
'
'            SetBuiltInProperty "Keywords", myMi.sFif 'номер ФИФ
'
'            If .Sheets.Count = 1 Then
'
'                If InStr(.ActiveSheet.Name, myMi.sType) = 0 Then 'если имя модификации не включает в себя имя типа
'                    SetBuiltInProperty "Comments", myMi.sType & " -- " & .ActiveSheet.Name 'имя текущего листа в книгу
'                Else
'                    SetBuiltInProperty "Comments", .ActiveSheet.Name 'имя текущего листа в книгу
'                End If
'            Else
'
'                Dim serialNum As String
'                If FindCellRight("Заводской / серийный номер:", rTempCell) Then
'                    serialNum = CStr(rTempCell)
'                    DeleteSpaceStEnd serialNum
'                End If
'
'                Dim propertyText As String
'                propertyText = myMi.sType
'
'                If serialNum <> vbNullString Then _
'                    propertyText = propertyText & " -- " & serialNum
'
'                SetBuiltInProperty "Comments", propertyText
'            End If
'
'        End With
    End If
    
    If FindCellRight("ФИФ", rTempCell) Then
    
        Dim sTempFif As String
        sTempFif = Replace(myMi.sFif, "-", vbNullString) 'получить номер в фиф без тире

      
        If rTempCell.Column <> 2 Then 'в старых шаблонах ячейка могла быть не объединена
            rTempCell = Chr(150)
            
            If IsNumeric(sTempFif) Then _
                rTempCell = myMi.sFif
        End If
    End If
    
    If FindCellRight("Методика", rTempCell) Then rTempCell = myMi.sMetodic ': rTempCell.Font.Color = &H80000008: rTempCell.Font.Bold = False
    
    With objBaseWorkbook.ActiveSheet
        If myMi.bolEtal And InStr(.[l13], "эталон") > 0 Then .[l12] = "рабочего эталона 1-го разряда"
' ----------------------------------------------------------------
If Not myMi.bolEtal And InStr(objBaseWorkbook.name, "pr_") > 0 Then _
    .[e25:e26].EntireRow.Delete 'если загружается шаблон поверки СИ, а не эталона
' ----------------------------------------------------------------
        If FindCellRight("Вид поверки", rTempCell) Then
            rTempCell = "периодическая": rTempCell.Font.Italic = True
            
            If bolFirstVer Then rTempCell = "первичная": _
                rTempCell.Font.Italic = False: rTempCell.Font.Bold = True: .[j15] = "-": .[i16] = "-"
        End If
        
        If myMultiSel.bolThisWBSheetLoad = False Then
            If InStr(CStr(.[i8]), "Дата") > 0 Then .[j8] = Date  'передать текущую дату - старые протоколы
            If InStr(CStr(.[h8]), "Дата") > 0 Then .[i8] = Date 'передать текущую дату - предыдущие протоколы
            If InStr(CStr(.[i10]), "дата") > 0 Then .[i8] = Date 'новейшие протоколы
            If InStr(CStr(.[i4]), "дата") > 0 Then .[i2] = Date 'новейшие протоколы
        End If
    End With
    
    Set rTempCell = Nothing
End Sub

'#########################################################
'Функция производит поиск ячейки в блоке и возвращает адрес ячейки (смежной ячейки)
Function FindCellRight( _
    sSearchStr As String, _
    ByRef searchResultCell As Range, _
    Optional BottomCell As Boolean, _
    Optional ThisCell As Boolean, _
    Optional SelectAfterFind As Boolean, _
    Optional ByVal objWs As Worksheet) _
        As Boolean
    
    FindCellRight = False
    If objWs Is Nothing Then Set objWs = ActiveWorkbook.ActiveSheet
    
    With objWs

        Dim prArea As String, findedCell As Range
        If objWs.PageSetup.PrintArea <> "" Then
            Set findedCell = Range(objWs.PageSetup.PrintArea).Find(sSearchStr, LookAt:=xlPart) 'поиск в области печати
        Else
            Set findedCell = objWs.Cells.Find(sSearchStr) 'поиск по всему листу
        End If
        
        If Not findedCell Is Nothing Then  'что-то найдено
            
            FindCellRight = True
            
            Set searchResultCell = findedCell.Offset(0, 1) 'по умолчанию - ячейка справа
            If BottomCell Then Set searchResultCell = findedCell.Offset(1, 0) 'ячейка ниже
            If ThisCell Then Set searchResultCell = findedCell 'найденная ячейка
            
            If SelectAfterFind Then searchResultCell.Select
        End If

        Set findedCell = Nothing: Set objWs = Nothing
    End With
End Function
'###################################################################
'Процедура копирует шапку предыдущей страницы
Sub CopyPrevTempHead( _
    objBaseWorkbook As Object _
)
    With objBaseWorkbook
    
        Dim firstSheet As Worksheet, _
            startRow As Integer, _
            endRow As Integer
            
        Set firstSheet = .Worksheets(1)
        GetHeadRows firstSheet, startRow, endRow
        
        Application.DisplayAlerts = False
        firstSheet.Rows(startRow).Resize(endRow - startRow + 1).Copy
        
        Dim newSheet As Worksheet
        Set newSheet = .Worksheets(.Worksheets.count)
        
        With newSheet
            
            .Rows(startRow).Resize(endRow - startRow + 1).Select
            .Paste
            
            If InStr(.[l9], "номер") > 0 Then
                .[l8].numberFormat = "General"
                .[l8].FormulaR1C1 = "='" & firstSheet.name & "'!RC:RC[1]"
                
            End If
            
        End With
        Application.DisplayAlerts = True
        
        Set firstSheet = Nothing
        Set newSheet = Nothing
    End With
End Sub
    Private Sub GetHeadRows( _
        ws As Worksheet, _
        startRow As Integer, _
        endRow As Integer _
    )
        With ws
        
            Dim i As Byte
            Do While i < 30
            
                i = i + 1
                If InStr(.Cells(i, 1), "НИО") > 0 Then startRow = i - 1
                If InStr(.Cells(i, 1), "svid export") > 0 Then endRow = i - 1
                
                If Not startRow = 0 And Not endRow = 0 Then Exit Do
            Loop
            
            .Rows(startRow).Resize(endRow - startRow + 1).Copy
        End With
        
    End Sub
    
' ----------------------------------------------------------------
' Дата: 17.03.2023 09:58
' ----------------------------------------------------------------
Private Sub Protocol_BaseDir()
        
    Dim prBaseDir As String
    prBaseDir = fso.BuildPath(configDirNew, "protocol base")
    
    If Not fso.FolderExists(prBaseDir) Then _
        MkDir prBaseDir
        
    Set fso = Nothing
    Explorer.OpenFolder prBaseDir, True
    
End Sub

' ----------------------------------------------------------------
' Дата: 17.03.2023 09:58
' ----------------------------------------------------------------
Private Sub RenameProtocolBases()
    
    Dim fileDir As Object
    Set fileDir = fso.GetFolder(fso.BuildPath(configDirNew, "protocol base"))
    
    Dim oFile As Object
    For Each oFile In fileDir.Files
        
        Dim newName As String
        newName = Replace(oFile.name, "_temp", "_base")
        
        Name fso.BuildPath(fileDir, oFile.name) As fso.BuildPath(fileDir, newName)
        
    Next
    
    Set fileDir = Nothing
    Set fso = Nothing
    
    
End Sub
'###################################################################
'процедура производит загрузку непосредственно файла протокола по имеющимся данным
Private Sub GetMeWorkFile( _
    sFullName As String, _
    sBaseDir As String, _
    sWorkDir As String, _
    meEmp As EmpDivision, _
    myMi As MeasInstrument _
    )
    
    Dim sShortFileName As String, _
        bolFirstVerify As Boolean
        
    sShortFileName = TrueName(sFullName) 'получить имя файла без пути (отделить на \)
    
    If InStr(sShortFileName, FirstVerifyString) > 0 Then _
        bolFirstVerify = True
    
    Dim sOldFileName As String, _
        sFileExt As String, _
        sFileTemp As String, _
        sFilePref As String
        
    sOldFileName = sFullName 'полное имя открываемого для загрузки шаблона
    
    Dim sCheckFif As String, _
        sCheckRef As String
        
    sCheckFif = Replace(myMi.sFif, "-", "") 'получить номер в фиф без дефисов
    sCheckRef = Replace(myMi.sRef, "-", "") 'получить номер в фиф перекрёстной ссылки без дефисов
    
    sFilePref = DEFAULT_FILE_PREFIX 'prc или prm
    
    Dim bodyPath As String
    bodyPath = Left(sFullName, Len(sFullName) - Len(sShortFileName))
    
    If IsNumeric(sCheckFif) Then 'то, что поверяется
        sFilePref = "pr_"

        If Dir(bodyPath & "fif_" & "*") <> vbNullString Then _
            sFilePref = "jr_"
    Else 'то, что калибруется
    
        If Dir(bodyPath & "fif_" & "*") <> vbNullString Then _
            sFilePref = "jrc_" 'тип загружаемой шапки
    End If
        
    If InStr(sShortFileName, "body") > 0 Then 'новый тип шаблона

        Dim prBaseDir As String
        prBaseDir = Config.templatesBasePath
        
        sFileTemp = Dir(prBaseDir & Application.PathSeparator & sFilePref & "*base*")
        sOldFileName = vbNullString 'поиск файла
        
        If sFileTemp <> vbNullString Then _
            sOldFileName = fso.BuildPath(prBaseDir, sFileTemp)
            
        Set fso = Nothing
        
        If sOldFileName = vbNullString Then _
            MsgBox "Отсутствует шаблон «шапки» протокола в директории " & vbNewLine & prBaseDir: _
            Exit Sub
            
    End If

    Dim objBaseWorkbook As Object, _
        objBodyWb As Object
        
    If myMultiSel.bolThisWBSheetLoad Then 'загрузить в эту книгу
    
        Set objBaseWorkbook = ActiveWorkbook
        Set objBodyWb = Application.Workbooks.Open(sOldFileName, , True) 'открыть книгу в режиме чтения
        
        With objBodyWb
            .ActiveSheet.Copy _
                After:=objBaseWorkbook.Worksheets(objBaseWorkbook.Worksheets.count)
            
            .Close False  'вставить в конце
        End With
        
        CopyPrevTempHead objBaseWorkbook   'вставить шапку предыдущей страницы
    Else  'загрузить в новую отдельную книгу
               
        SetBuiltInProperty "Comments", , True 'очистить свойство
        SetBuiltInProperty "Category", , True 'очистить свойство
        
'        Dim bodyPath As String
'        bodyPath = Left(sFullName, Len(sFullName) - Len(sShortFileName))
'        If Dir(bodyPath & "fif_" & "*") <> vbNullString Then _
'        sFilePref = "jr_" 'тип загружаемой шапки
        
        Dim sNewFileName As String
        sNewFileName = sFilePref  'префикс файла для поиска нового шаблона

        sFileExt = GetExt(sFullName) 'получить расширение файла
        sNewFileName = sNewFileName & "TEMP_" & Format(Date, "dd-mm-yyyy") & sFileExt
'        sNewFileName = sNewFileName & meEmp.sDepPref & "_" & meEmp.sLabNum & "_TEMP_" & Right(Date, 2) & sFileExt
                
        sNewFileName = ReturnNotExistingName( _
                                            Load_Directory(sWorkDir), _
                                            sNewFileName)  'проверить имя файла и добавить индекс

        ActiveWorkbook.SaveAs sNewFileName, xlWorkbookDefault 'сохранить как обычная книга

        Set objBaseWorkbook = ActiveWorkbook
        Set objBodyWb = Application.Workbooks.Open(sOldFileName, , True) 'открыть книгу в режиме чтения
        
        With objBodyWb
            .ActiveSheet.Copy After:=objBaseWorkbook.Worksheets(objBaseWorkbook.Worksheets.count): .Close False  'вставить в конце
        End With
        
        Dim wsDelCount As Byte, i As Byte
        wsDelCount = ActiveWorkbook.Sheets.count - 1 'количество листов для удаления в книге

        Application.DisplayAlerts = False
        Do While wsDelCount > 0
            ActiveWorkbook.Sheets(1).Delete
            wsDelCount = wsDelCount - 1
        Loop
        Application.DisplayAlerts = True
                   
        ActiveWorkbook.Save
    End If
    
    InsertBaseMIdata objBaseWorkbook, meEmp, myMi, bolFirstVerify ' передать в конкретные ячейки базовые данные
    
    If myMultiSel.bolThisWBSheetLoad Then _
        CopyPrevTempHead objBaseWorkbook
        
    If InStr(sShortFileName, "body") > 0 Then _
        InsertBodyTemplate objBaseWorkbook, sFullName, myMi 'новый тип шаблона - объединить
    
    FillInstrumentType objBaseWorkbook, myMi
    'TrueConvertation myMi 'заменить "поверка" на "калибровка" и наоборот в тексте

    myMultiSel.bolThisWBSheetLoad = False
    Set objBodyWb = Nothing: Set objBaseWorkbook = Nothing
End Sub
    Private Function Load_Directory( _
        default_workDir As String _
        ) As String
        
        Load_Directory = default_workDir 'по умолчанию - каталог загрузки пустых шаблонов
        
        Dim aWbPath As String
        aWbPath = ActiveWorkbook.path
        
        If aWbPath = vbNullString Then _
            Exit Function 'если работа идёт с новосозданной книгой
            
        aWbPath = aWbPath & Application.PathSeparator
        If aWbPath <> default_workDir Then Load_Directory = aWbPath
        
    End Function

    Private Sub FillInstrumentType( _
        objBaseWorkbook As Workbook, _
        myMi As MeasInstrument _
    )
            
        With objBaseWorkbook
            Dim sTempStr As String, _
                sWBName As String
                
            sWBName = .name
            sTempStr = Replace(myMi.sFif, "-", "") 'проверка номера фиф на поверку
        
            SetBuiltInProperty "Keywords", myMi.sFif 'номер ФИФ
            
            If .Sheets.count = 1 Then
            
                If InStr(.ActiveSheet.name, myMi.sType) = 0 Then 'если имя модификации не включает в себя имя типа
                    SetBuiltInProperty "Comments", myMi.sType & " -- " & .ActiveSheet.name 'имя текущего листа в книгу
                Else
                    SetBuiltInProperty "Comments", .ActiveSheet.name 'имя текущего листа в книгу
                End If
            Else
            
                Dim serialNum As String, _
                    rTempCell As Range
                    
                If FindCellRight("Заводской / серийный номер:", rTempCell) Then
                    serialNum = CStr(rTempCell)
                    DeleteSpaceStEnd serialNum
                End If
                
                Dim propertyText As String
                propertyText = myMi.sType
                
                If serialNum <> vbNullString Then _
                    propertyText = propertyText & " -- " & serialNum
                    
                SetBuiltInProperty "Comments", propertyText
            End If
            
        End With
    
    End Sub


'###################################################################
'процедура заменяет значения по тексту при загрузке шаблона независимо от тела шаблона
Sub TrueConvertation(myMi As MeasInstrument)
    
    Dim sFifNum As String, sRefNum As String, bolFif As Boolean, bolRef As Boolean
    sFifNum = Replace(myMi.sFif, "-", ""): sRefNum = Replace(myMi.sRef, "-", "")
    
    If IsNumeric(sFifNum) Then bolFif = True 'номер принадлежит средству, подлежащему поверке
    If IsNumeric(sRefNum) Then bolRef = True 'номер принадлежит средству, подлежащему поверке
    
    'If bolFif And bolRef = False Then ReplaceData 1, 2 'заменить "калибровка" на "поверка"
    'If bolFif = False And bolRef Then ReplaceData 2, 1 'заменить "поверка" на "калибровка"
End Sub
'###################################################################
'процедура собирает шаблон протокола из компонентов база+тело
Sub InsertBodyTemplate( _
    wbTemplateBase As Object, _
    bodyTemplatePath As String, _
    currentMI As MeasInstrument _
    )
    
    Dim i As Byte, _
        bodyRow As Byte, _
        instrumentNameRow As Byte, _
        cellText As String
        
    i = 1
    Do While i < 255
        
        cellText = CStr(wbTemplateBase.ActiveSheet.Cells(i, 1))
        
        If InStr(cellText, "Наименование,") > 0 Then instrumentNameRow = i
        If InStr(cellText, "body") > 0 Then bodyRow = i: Exit Do
        i = i + 1
    Loop
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Dim wbBody As Object, _
        bodyStartRow As Byte, _
        bodyEndRow As Byte, _
        instrumentName As String, _
        etalonRank As String
        
    Set wbBody = Application.Workbooks.Open(bodyTemplatePath, , True) 'открыть книгу с телом протокола только для чтения
    
    With wbBody.ActiveSheet
    
        bodyEndRow = .Cells(.Rows.count, 1).End(xlUp).Row 'последняя заполненная строка снизу вверх
        i = bodyEndRow
        
        Do While i > 0
            
            cellText = CStr(.Cells(i, 1))
            
            If InStr(cellText, "end body") > 0 Then
                
                bodyEndRow = i 'строка конца тела протокола
                If InStr(cellText, "=") > 0 Then _
                    etalonRank = cellText: _
                    etalonRank = Right(etalonRank, Len(etalonRank) - InStr(etalonRank, "=")) 'тип эталона
            End If
            
            If InStr(cellText, "body") > 0 And InStr(cellText, "end body") = 0 Then
                bodyStartRow = i 'строка начала тела протокола
                
                If InStr(cellText, "=") > 0 Then _
                    instrumentName = cellText: _
                    instrumentName = Right(instrumentName, Len(instrumentName) - InStr(instrumentName, "=")) 'частное название СИ
            End If
            i = i - 1
        Loop
        
        .Rows(bodyStartRow + 1).Resize(bodyEndRow - bodyStartRow - 1).Copy 'копировать нужную часть тела протокола
    End With
    
    With wbTemplateBase.ActiveSheet
        
        If myMultiSel.bolThisWBSheetLoad = False Then
        
            If instrumentName <> vbNullString Then _
                .Cells(instrumentNameRow, 1).Offset(0, 1) = currentMI.sName & " " & currentMI.sType & ", модификация " & instrumentName ' передать наименование из файла тела
                
            If etalonRank <> vbNullString Then _
                .Cells(instrumentNameRow, 1).Offset(0, 10) = etalonRank ' передать разряд эталона
                
        End If
        
        If bodyRow <> 0 Then
            
            .Cells(bodyRow, 1).Insert
            .Cells(bodyRow + bodyEndRow - bodyStartRow - 1, 1).EntireRow.Delete 'удалить строчку body
            
        End If
        
    End With
    
    wbTemplateBase.ActiveSheet.name = NonTakenShName(wbTemplateBase, wbBody.ActiveSheet.name) 'присвоить правильное имя текущему листу
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    wbBody.Close False
    Set wbBody = Nothing
End Sub
    '###################################################################
    'процедура корректно переименовывает лист при вставке
    Private Function NonTakenShName( _
        objWorkBook As Object, _
        currentName As String _
        ) As String
        
        NonTakenShName = currentName
        
        Dim isNameTaken As Boolean
        isNameTaken = SheetExists(objWorkBook, currentName)
        
        If Not isNameTaken Then Exit Function
        
        Dim newName As String
        newName = currentName
        
        Dim j As Integer
        j = 2
        
        Do While isNameTaken
            
            newName = currentName & " (" & j & ")"
            isNameTaken = SheetExists(objWorkBook, newName)
            j = j + 1
            
        Loop
        
        NonTakenShName = newName
        
    End Function
        Private Function SheetExists( _
            Workbook As Object, _
            sheetName As String _
            ) As Boolean
            
            SheetExists = False
            
            Dim i As Integer
            For i = 1 To Workbook.Sheets.count
    
                If Workbook.Sheets(i).name = sheetName Then _
                    SheetExists = True: _
                    Exit Function 'поиск совпадений имён
            Next
        
        End Function
            
'#########################################################
'процедура заполняет данные об эталонах
Private Sub FillEtalonsAfterImport( _
    Optional noEtalData As Boolean, _
    Optional objWs As Worksheet _
    )
    
    Dim myBase2 As New Z_clsmBase, WorkClsm2 As New Z_clsmSearch, sArrDataBase2() As String
    
    Application.ScreenUpdating = False
    
    With myBase2 'работа с классом настроек
        .GetArrFF setDir, Environ("USERNAME") & ".uCfg" 'загрузить в класс файл настроек
        .AddP "startDir", "etalDB"
        .AddP "headNAME", "headSTATE"
        .AddP "empNAME", "empSTATE"
        .AddP "empSecNAME", "empSecSTATE"
        
        .FillValues 'обязательно: найти значения выходных параметров по ключам
    End With
    

    With WorkClsm2
        .FillConfiguration myBase2.Parameters, myBase2.values 'передать извлечённые из настроек параметры в специальный класс
        
        
        
        
        
        
        
        ' ------------------------------------------------------
        'todo: FillLastName -- отвязать workclsm -- имя, должность
        FillLastName .headNAME, .headSTATE, True, , objWs 'передать фамилию и должность руководителя - для протокола калибровки снизу
        FillLastName .empSecName, .empSecState, , True, objWs 'передать фамилию и должность второго исполнителя
        FillLastName .empName, .empState, , , objWs 'передать фамилию и должность исполнителя
        ' ------------------------------------------------------
        
        
        
        
        
        
        
        
'        If noEtalData = False And FileExist(.startDir, .DbName) Then 'если база данных обнаружена по указанному пути
'            sArrDataBase2 = .FillDataBase(myBase2.GetArrFF(.startDir, .DbName), True)  'преобразовать массив файла в массив базы данных
'
'            If UBound(sArrDataBase2) > 0 Then PasteEtalonsData sArrDataBase2 'если массив файла был получен
'        End If
        
    End With
    Application.ScreenUpdating = True
End Sub



'#########################################################
'Процедура ищет и заполняет сведения по эталонам
Sub PasteEtalonsData(myDataBase() As String)
    With ActiveWorkbook.ActiveSheet
    
        Dim i As Byte, bEtalWorkRow As Byte, sCellStr As String
        i = 1: bEtalWorkRow = 1
        
        Do While i < 40 'найти строку начала заполнения данными
            sCellStr = CStr(.Cells(i, 1))
            If InStr(LCase(sCellStr), "эталоны") > 0 Then bEtalWorkRow = i + 3: Exit Do
            i = i + 1
        Loop
        
        Application.ScreenUpdating = False
        Do While CStr(.Cells(bEtalWorkRow, 1)) <> vbNullString  'выполнять, пока не кончатся строки
            
            With .Cells(bEtalWorkRow, 1)
            
                Dim sSearchType As String, sSearchNumber As String, j As Byte
                sSearchType = CStr(.Offset(0, 1))
                sSearchNumber = CStr(.Offset(0, 2))  'получить значения для сопоставления
                
                DeleteSpaceStEnd sSearchType: DeleteSpaceStEnd sSearchNumber
                For j = LBound(myDataBase, 2) To UBound(myDataBase, 2)
                    
                    If InStr(sSearchType, myDataBase(1, j)) > 0 Then 'опознано совпадение по типу
                        
                        If sSearchNumber <> vbNullString Then
                            
                            If InStr(myDataBase(2, j), sSearchNumber) > 0 Then _
                                GiveTrueDate myDataBase(3, j), .Offset(0, 3): Exit For  'двойное совпадение
                        Else
                            GiveTrueDate myDataBase(3, j), .Offset(0, 3): Exit For  'простое совпадение
                        End If
                        
                    End If
                        
                Next j
            End With
            
            bEtalWorkRow = bEtalWorkRow + 2
        Loop
        Application.ScreenUpdating = True
        
        If bEtalWorkRow <> 1 Then .[a1].Select 'когда данные были переданы
     '   If bEtalWorkRow = 1 Then ActiveCell = "Блок эталонов не обнаружен."
    End With
End Sub
'#########################################################
'Процедура проверяет данные и вставляет их корректно в ячейку
Sub GiveTrueDate(sCommentToPaste As String, objCell As Range)

    Dim sCellText As String
    sCellText = objCell: If sCellText = "" Then sCellText = sCommentToPaste
    
    Dim iStPos As Integer, sLeftStr As String, sRightStr As String, dDateOfComm As Date
    iStPos = InStr(sCellText, "действител") 'поиск даты в тексте
    dDateOfComm = Date + 30 'по умолчанию дата действия документа о поверке удовлетворительна
    
    If iStPos > 0 Then 'найдена дата в данных примечания
        sLeftStr = Left(sCommentToPaste, iStPos - 2): DeleteSpaceStEnd sLeftStr
        sRightStr = Right(sCommentToPaste, Len(sCommentToPaste) - Len(sLeftStr)): DeleteSpaceStEnd sRightStr
        dDateOfComm = CDate(Right(sRightStr, 10)): sCellText = sLeftStr & vbLf & sRightStr
    End If
    
    objCell = sCellText
    
    With Range(Cells(objCell.Row, 1), objCell)     'покрасить строку в зависимости от даты
        .Interior.Pattern = xlNone: .Font.ColorIndex = xlAutomatic 'чёрный цвет шрифта
        
        Select Case dDateOfComm - Date 'разница дат до окончания срока поверки
            Case Is <= 0 'просрочено
                '.Font.Color = -16776961 'красный цвет шрифта
                .Font.color = 10921638 'серый цвет шрифта
            Case Is <= 7 'осталась неделя
                .Font.color = -16776961 'красный цвет шрифта
                '.Font.Color = 36799 'желтый цвет шрифта
                '.Font.Italic = True 'выделить курсивом
                '.Interior.Color = 65535 'жёлтый цвет
            Case Is <= 21 'осталось 3 недели
                '.Font.Color = -6279056 'фиолетовый цвет '.Font.Color = 24704 'тёмно-золотистый цвет
                .Font.color = 36799 'желтый цвет шрифта
        End Select
    End With
End Sub
'###################################################################
'процедура заполняет текущие свойства файла при формировании шаблона протокола
Sub FillTempProperties(myInstr As MeasInstrument, sDestDir As String)
    DeleteWrongSheet 'удалить левый лист из списка
    
    GetTrueModificationMi myInstr 'получить значение модификации
 
    Dim sSaveName As String, sAdditionalText As String
    sSaveName = "body_": sAdditionalText = myInstr.sMetodic
    If myInstr.sModification <> "" Then sAdditionalText = myInstr.sModification
    
    sSaveName = sSaveName & myInstr.sFif & " " & sAdditionalText  'маска имени нового файла
    Application.ScreenUpdating = False
    
    Dim rHeadCell As Range
    If SheetIsEmpty(ActiveWorkbook.ActiveSheet) Then 'пустой лист
        Dim objBaseWb As Object, objBodyWb As Object
        Set objBaseWb = ActiveWorkbook: Set objBodyWb = ThisWorkbook

        With objBodyWb 'книга надстройки
            .Sheets.item("protocol").Copy After:=objBaseWb.Worksheets(objBaseWb.Worksheets.count) 'вставить в конце
            Application.DisplayAlerts = False: objBaseWb.Worksheets(1).Delete: Application.DisplayAlerts = True

            Set rHeadCell = objBaseWb.ActiveSheet.[a1]: rHeadCell = GetHeadBodyText(myInstr.sType, myInstr.sMetodic, myInstr.sName)  'получить заголовок
        End With
        
        Set objBaseWb = Nothing: Set objBodyWb = Nothing
    Else 'непустой лист
        Dim rEndCell As Range
        
        With ActiveWorkbook.ActiveSheet
            
            If FindCellRight("эталоны", rHeadCell, , True) Then 'если найдена ячейка "эталоны"
                Set rHeadCell = rHeadCell.Offset(-1, 0)
                FillHeadCell rHeadCell, myInstr, sSaveName 'преобразовать ячейку верхней границы
            
                If rHeadCell.Row > 1 Then _
                    .Range(.[a1], rHeadCell.Offset(-1, 0)).EntireRow.Delete 'удалить все строки выше
                .Range(rHeadCell, Cells(rHeadCell.Row, 10)).Merge  'объединить ячейки
                
                If .[o1] <> ThisWorkbook.Worksheets("Лист1").[o1] Then 'если блок справа отсутствует
                    ThisWorkbook.Worksheets("Лист1").Columns("N:Q").Copy 'базовый блок экспорта
                    .Columns("N:Q").PasteSpecial: .[a1].Select
                End If
                
                FillEtalStringBlock rHeadCell 'процедура на заполнение ячейки эталонов справа
            End If
            
            If GetrEndCell(rEndCell) Then 'получить ячейку нижней границы
            
                FillEndCell rEndCell 'преобразовать ячейку нижней границы
                .Range(rEndCell, Cells(rEndCell.Row, 10)).Merge 'объединить ячейки
                
                Application.CutCopyMode = False
                .Cells.Replace What:="$", Replacement:="" 'убрать абсолютизацию всех ячеек
                
                Dim shShape As Shape
                For Each shShape In .Shapes
                    shShape.Delete 'удалить все рисунки, формулы листа
                Next
            End If
        End With
    End If
    
    With ActiveWorkbook.ActiveSheet
        .name = myInstr.sType: If myInstr.sType <> myInstr.sMetodic Then .name = myInstr.sMetodic 'передать имя активного листа
    End With
          
    SaveNewName sSaveName, True, False, True, sDestDir 'сохранить текущую книгу excel как копию исходного документа
    FillThisPathDocTempName myInstr 'добавить шаблон wd
    Application.ScreenUpdating = True
End Sub
'###################################################################
'процедура обрабатывает данные о модификации СИ
Sub GetTrueModificationMi(ByRef myInstr As MeasInstrument)
    
    With myInstr
        DeleteSpaceStEnd .sModification 'удалить пустые символы в начале и конце наименования
    
        Select Case .sModification
            Case "", "без модификации"
                .sModification = "": .sMetodic = .sType 'МКС-А строка, которая будет помещена в свойства
                
            Case Else 'получена модификация
                '.sMetodic = .sType & " " & .sModification
                .sMetodic = .sModification
                
                If InStr(.sModification, .sType) > 0 Then _
                    .sMetodic = .sModification 'МКС-А МКС-А03-1Н превратить в МКС-А03-1Н
        End Select
    End With
    
    With ActiveWorkbook
       .BuiltinDocumentProperties("Keywords") = myInstr.sFif 'номер в фиф
       .BuiltinDocumentProperties("Comments") = myInstr.sMetodic 'полное наименование типа СИ с модификацией
    End With
End Sub
'###################################################################
'процедура заполняет ячейки ЭТАЛ1 и ЭТАЛ2 справа
Sub FillEtalStringBlock(Optional rHeadCell As Range, Optional bolMultiProtExport As Boolean)
    Dim rEtalCell As Range, sArrExport() As String ' массив со сведениями об эталонах
    
    If bolMultiProtExport Then 'заполнение сведений для полистраничного протокола
        If GetEtalExportCell(rEtalCell, ActiveWorkbook.Worksheets(1)) = False Then Exit Sub  'если ячейка не найдена
        If FillEtalArrayMultiExport(sArrExport) = False Then Exit Sub 'если в массив не было передано ни одного значения
    Else 'заполнение сведений для текущей страницы
        If GetEtalExportCell(rEtalCell, ActiveWorkbook.ActiveSheet) = False Then Exit Sub 'если ячейка не найдена
    
        Dim rWorkCell As Range '- поиск ячейки в блоке эталонов протокола
        Set rWorkCell = rHeadCell.Offset(4, 0): Set rWorkCell = rWorkCell.Offset(0, 1) 'первая заполненная ячейка с типом эталона
        
        If FillEtalArrayExportCell(sArrExport, rWorkCell) = False Then Exit Sub 'если в массив не было передано ни одного значения
    End If
    
    '########################### - формирование двух строк эталонов
    Dim sTempStr As String, i As Byte, sEtalFirstStr As String, sEtalSecondStr As String
    
    SortMassOne sArrExport, , True
    
    For i = LBound(sArrExport) To UBound(sArrExport)
        sTempStr = sTempStr & sArrExport(i) & ", "
    Next

    sEtalFirstStr = TrueNameLength(sTempStr, EtalonNameMaxLength): DeleteSpaceStEnd sEtalFirstStr
    If Right(sEtalFirstStr, 1) = "," Then sEtalFirstStr = Left(sEtalFirstStr, Len(sEtalFirstStr) - 1) 'убрать запятую
    
    sEtalSecondStr = TrueNameLength(sTempStr, EtalonNameMaxLength, True)
    If sEtalSecondStr <> "" Then sEtalSecondStr = sEtalSecondStr & " "
    
    sTempStr = "наименование эталона приведено на оборотной стороне"
    If UBound(sArrExport) > 0 Then sTempStr = "наименование эталонов приведено на оборотной стороне"
    
    sEtalSecondStr = sEtalSecondStr & sTempStr
    '########################### - передать значения в ячейки
    Set rEtalCell = rEtalCell.Offset(0, 1)
    
    rEtalCell.value = sEtalFirstStr
    rEtalCell.Offset(1, 0).value = sEtalSecondStr
    
    Set rEtalCell = Nothing: Set rWorkCell = Nothing
End Sub
'###################################################################
'функция заполняет массив данных эталона по всем страницам листа для передачи в блок экспорта справа
Function FillEtalArrayMultiExport(ByRef sArrTemp() As String) As Boolean
    FillEtalArrayMultiExport = False
    
    Dim ws As Worksheet, wsCnt As Byte, sTempStr As String
    wsCnt = ActiveWorkbook.Sheets.count: ReDim sArrTemp(0)
    
    For Each ws In ActiveWorkbook.Worksheets 'пройтись по всем листам
    
        Dim iRowSvidExp As Integer
        iRowSvidExp = FindSvidExportRow(ws) ' 1 - передать номер строки с блоком экспорта
        
        If iRowSvidExp <> -1 Then 'опознано наличие блока экспорта
        
            Dim rEtalCell As Range, sArrWork() As String, i As Byte, j As Byte, sArrHelp() As String
            If GetEtalExportCell(rEtalCell, ws) Then  'если ячейка найдена в блоке экспорта справа
            
                sTempStr = rEtalCell.Offset(0, 1) & ", " & rEtalCell.Offset(1, 1) 'соединить строки эталонов 1 и 2
                
                sArrWork = Split(sTempStr, ", ") 'разбить строку
                ReDim Preserve sArrWork(UBound(sArrWork) - 1) 'уменьшить размерность массива на 1 - удалить элемент "наименование эталонов..."
                ReDim sArrHelp(1, UBound(sArrWork)) 'вспомогательный массив
                
                For i = LBound(sArrWork) To UBound(sArrWork) 'пройтись по всему массиву полученных строк эталонов
                    sArrHelp(LBound(sArrHelp), i) = sArrWork(i) 'наименование эталона
                    sArrHelp(UBound(sArrHelp), i) = False 'заполнение сведениями этого эталона
                    
                    For j = LBound(sArrTemp) To UBound(sArrTemp) 'пройтись по всем значениям имеющихся строк эталонов
                        If InStr(sArrTemp(j), sArrWork(i)) > 0 Then sArrHelp(UBound(sArrHelp), i) = True 'наименование эталона - элемент уже в массиве
                    Next
                Next
                
                For i = LBound(sArrHelp, 2) To UBound(sArrHelp, 2)
                    If sArrHelp(UBound(sArrHelp), i) = False Then
                        If sArrTemp(UBound(sArrTemp)) <> "" Then ReDim Preserve sArrTemp(UBound(sArrTemp) + 1)
                        sArrTemp(UBound(sArrTemp)) = sArrHelp(LBound(sArrHelp), i) 'добавить элемент
                    End If
                Next
            End If
        End If
    Next
    If sArrTemp(LBound(sArrTemp)) <> "" Then FillEtalArrayMultiExport = True 'если был передан хотя бы 1 элемент
End Function

'###################################################################
'функция находит ячейку эталонов справа в блоке экспорта
Function GetEtalExportCell(ByRef rEtalCell As Range, ws As Worksheet) As Boolean
    GetEtalExportCell = False
   
    Set rEtalCell = ws.Cells(6, 11).End(xlToRight) 'первая заполненная ячейка справа
    If rEtalCell.Column > 20 Then Exit Function 'защита от отсуствтия блока
    
    Set rEtalCell = ws.Cells(1, rEtalCell.Column).End(xlDown) 'первая заполненная строка сверху, считая с 1 строки
    
    Dim K As Byte
    Do While K < 2 'выполнять, пока не будет 2 пустых ячейки подряд
        If InStr(CStr(rEtalCell), "#ЭТАЛ") > 0 Then Exit Do
        If CStr(rEtalCell) = "" Then K = K + 1 'счётчик пустых ячеек
        If CStr(rEtalCell) <> "" Then K = 0 'счётчик пустых ячеек
        Set rEtalCell = rEtalCell.Offset(1, 0) 'сместиться на ячейку ниже
    Loop
    
    If CStr(rEtalCell) = "" Then Exit Function 'если ячейка не найдена
    GetEtalExportCell = True
End Function
'###################################################################
'функция заполняет массив данных эталона для передачи в блок экспорта справа
Function FillEtalArrayExportCell(ByRef sArrTemp() As String, rWorkCell As Range) As Boolean
    FillEtalArrayExportCell = False
    
    Dim sTempName As String, sTempNum As String
    ReDim sArrTemp(0)
    
    Do While CStr(rWorkCell) = "-" 'выполнять, пока обрабатывается эталон без типа СИ
        sTempName = CStr(rWorkCell.Offset(0, -4)) 'получить наименование эталона
        sTempNum = CStr(rWorkCell.Offset(0, 1)) 'получить номер эталона
        
        If InStr(sTempName, "ГЭТ") > 0 Or InStr(sTempName, "ГВЭТ") > 0 Or InStr(sTempName, "эталон") > 0 Then   'строка этлона
            If sArrTemp(UBound(sArrTemp)) <> "" Then _
                ReDim Preserve sArrTemp(UBound(sArrTemp) + 1) 'расширить массив, если он полон
            
            If InStr(sTempName, " ГЭТ") > 0 Then sTempName = Right(sTempName, Len(sTempName) - InStr(sTempName, " ГЭТ")) 'если в наименовании присутствует пояснение перед типом эталона
            If InStr(sTempName, " ГВЭТ") > 0 Then sTempName = Right(sTempName, Len(sTempName) - InStr(sTempName, " ГВЭТ")) 'если в наименовании присутствует пояснение перед типом эталона
            
            sArrTemp(UBound(sArrTemp)) = sTempName & ", " 'по умолчанию передать наименование эталона (для перавичных)
            If InStr(sTempNum, "ZZB") > 0 Then sArrTemp(UBound(sArrTemp)) = "рег. № " & sTempNum & ", " 'передать номер для вторичных и рабочих
            
        End If
        Set rWorkCell = rWorkCell.Offset(1, 0) 'сместиться на 2 ячейки ниже
    Loop
    If sArrTemp(LBound(sArrTemp)) <> "" Then FillEtalArrayExportCell = True 'если найден хоть один эталон
End Function
'###################################################################
'процедура заполняет ячейку body
Sub FillHeadCell(rHeadCell As Range, myInstr As MeasInstrument, sSaveName As String)
    With rHeadCell
        .Font.color = 0: .HorizontalAlignment = xlCenter: .Interior.color = 65535 'жёлтый цвет ячейки
        rHeadCell = GetHeadBodyText(myInstr.sType, myInstr.sMetodic, myInstr.sName)  'получить заголовок

        .Offset(1, 0) = "Средства поверки: эталоны и вспомогательное оборудование"
        If InStr(sSaveName, "bodyc_") > 0 Then .Offset(1, 0) = "Средства калибровки: эталоны и вспомогательное оборудование"
    End With
End Sub
'###################################################################
'получить позицию ячейки rEndCell
Function GetrEndCell(ByRef rSearchCell As Range) As Boolean
    Dim rClearCell As Range, bolClear As Boolean
    
    GetrEndCell = False
    With ActiveWorkbook.ActiveSheet
        Set rSearchCell = .Cells(Rows.count, 1).End(xlUp) 'последняя заполненная строка снизу
        If rSearchCell.text = "svid export" Then Set rClearCell = rSearchCell: bolClear = True  'убрать метку старого типа свидетельства
    
        Do While InStr(CStr(rSearchCell), "svid export") > 0 _
            Or InStr(CStr(rSearchCell), "Выполнил") > 0 Or InStr(CStr(rSearchCell), "произвёл") > 0 _
                Or InStr(CStr(rSearchCell), "Error") > 0
                
            Set rSearchCell = .Cells(rSearchCell.Row, 1).End(xlUp) 'последняя заполненная строка снизу
        Loop
    End With
    
    If bolClear Then rClearCell.value = ""
    If CStr(rSearchCell) <> "end body" Then Set rSearchCell = rSearchCell.Offset(1, 0)
    
    GetrEndCell = True
End Function
'###################################################################
'процедура заполняет ячейку end body
Sub FillEndCell(rEndCell As Range)

    With rEndCell
        .UnMerge
        .HorizontalAlignment = xlCenter: .Font.color = 0: .Font.Size = 8: .Interior.color = 65535        'жёлтый цвет ячейки
        rEndCell = "end body": .Offset(0, 11) = "'= разряд эталона в родительном падеже"
        .Offset(1, 0).Resize(8, 10).EntireRow.Delete 'удалить 8 строк, считая с текущей ячейки
        .Offset(1, 0).Resize(40, 10).EntireRow.UnMerge 'убрать обхединение ячеек
        
        Dim i As Byte
        For i = 2 To 20
            If .Offset(i, 3) = "" Then Exit For 'чтобы писать на пустой строке
        Next
        
        With .Offset(i, 1)
            .Font.Size = 14: .Font.color = -16776961 'красный цвет шрифта
            .value = "АБСОЛЮТИЗАЦИЯ ССЫЛОК В ЯЧЕЙКАХ УДАЛЕНА СО ВСЕГО ЛИСТА": .VerticalAlignment = xlCenter
        End With
        
    End With
End Sub
'###################################################################
'функция возвращает текст для заполнения шапки body =
Function GetHeadBodyText(sTypeMiBase As String, sInputTypeMi As String, sNameMiBase As String) As String
    GetHeadBodyText = "body" 'значение по умолчанию
    
    If sTypeMiBase <> sInputTypeMi Then _
        GetHeadBodyText = GetHeadBodyText & "=" & sInputTypeMi 'модификация СИ
         'GetHeadBodyText = GetHeadBodyText & "=" & sNameMiBase & " " & sTypeMiBase & ", модификация " & sInputTypeMi
End Function
'#################################################
'процедура удаляет левый лист из книги
Sub DeleteWrongSheet()
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name = "svidDSpec" Then ws.Delete: Exit For
    Next
    
    Application.DisplayAlerts = True
End Sub
'#########################################################
'функция определяет слово для вставки в InputBox модификации
Function GetDefaultInput(workInstr As MeasInstrument)
    GetDefaultInput = "без модификации" 'значение по умолчанию
    
    Dim sTempStr As String, sArrTemp() As String
    sTempStr = GetFileNameWithOutExt(ActiveWorkbook.name)  'имя текущей книги Excel без расширения pr_17406-17 МКС-А
    
    If InStr(sTempStr, workInstr.sFif) > 0 Then 'правится протокол ранее выбранного типа СИ
        sArrTemp = Split(sTempStr, workInstr.sFif) 'разбить строку на номере в фиф
        
        If sArrTemp(UBound(sArrTemp)) = "" Then Exit Function
        If sArrTemp(UBound(sArrTemp)) <> workInstr.sFif Then _
            GetDefaultInput = DeleteSpaceStEnd(sArrTemp(UBound(sArrTemp))) 'получить исходную модификацию
    End If
End Function
'#################################################
'функция проверяет, заполнен ли лист excel
Function SheetIsEmpty(ws As Worksheet) As Boolean
    SheetIsEmpty = True 'по умолчанию лист пустой

    If GetExt(ActiveWorkbook.name) = "" Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Range(ws.[a1], ws.[z50])) > 0 Then SheetIsEmpty = False
End Function


'###################################################################
'функция производит загрузку непосредственно файла протокола по имеющимся данным
Function LoadNewTempHead(sNewProtType As String, sBaseDir As String, _
    myCurrMI As MeasInstrument, myCurrEmp As EmpDivision, myCurrCust As myCustomer, objWs As Worksheet) As Boolean
    
    Dim rTempCell As Range
    If FindCellRight("svid", rTempCell, , , , objWs) = False Then Set rTempCell = Nothing: Exit Function 'защита от конвертации других страниц
    
    Dim iCurrBeardCnt As Integer, iNewBeardCnt As Integer, iBeardCntCopy As Integer, iCurrWsBeardStart As Integer
    iCurrBeardCnt = GetProtTypeBeardCount(sBaseDir) 'количество строк в текущем "подбородке"
    iNewBeardCnt = GetProtTypeBeardCount(sBaseDir, Right(sNewProtType, 1)) 'строк в "подбородке" нового типа протокола
    iCurrWsBeardStart = objWs.Cells(Rows.count, 1).End(xlUp).Row + 1 - iCurrBeardCnt + 1 'текущее начало подбородка
    
    iBeardCntCopy = iCurrBeardCnt: If iNewBeardCnt > iCurrBeardCnt Then iBeardCntCopy = iNewBeardCnt 'количество строк для копирования
     
    Dim sTempInstrSym As String, sTempFileName As String 'новый тип протокола
    sTempInstrSym = "pr_": If Right(sNewProtType, 1) = 2 Then sTempInstrSym = "prc_"
    If Right(sNewProtType, 1) = 3 Then sTempInstrSym = "prm_"
    
    sTempFileName = Dir(sBaseDir & sTempInstrSym & "*" & "temp" & "*"): If sTempFileName = "" Then Exit Function 'если не найден шаблон
    sTempFileName = sBaseDir & sTempFileName 'полный путь к файлу-шапке шаблона
     
    Dim objWorkBook As Object
    Set objWorkBook = Application.Workbooks.Open(sTempFileName, , True) 'открыть книгу в режиме чтения
    
    Application.DisplayAlerts = False
    
    With objWorkBook
        .ActiveSheet.Rows("1:1").Resize(iHeadProrCnt).Copy: objWs.Rows("1:1").Resize(iHeadProrCnt).PasteSpecial
        .ActiveSheet.Rows("25:25").Resize(iBeardCntCopy).Copy 'верхняя часть шапки
        
        objWs.Rows(iCurrWsBeardStart).PasteSpecial: .Close
    End With
    
    Application.CutCopyMode = False: Application.DisplayAlerts = True
    
    InsertBaseMIdata ActiveWorkbook, myCurrEmp, myCurrMI, False, True ' передать в конкретные ячейки базовые данные
    
    If FindCellRight("Заводской", rTempCell, , , , objWs) Then rTempCell = myCurrMI.sRef 'заводской номер текущего листа
    If FindCellRight("ФИФ", rTempCell, , , , objWs) Then rTempCell = myCurrMI.sFif  'номер в ФИФ текущего листа
    If FindCellRight("Наименование", rTempCell, , , , objWs) Then rTempCell = myCurrMI.sName 'наименование
    If FindCellRight("Методика", rTempCell, , , , objWs) Then rTempCell = myCurrMI.sMetodic  'методика
    
    If FindCellRight("Масса наполнителя", rTempCell, , , , objWs) Then _
                rTempCell = myCurrMI.sModification: rTempCell.Offset(0, 2) = myCurrMI.sType   'масса и объём
    
    With myCurrCust
        If FindCellRight("Заказчик", rTempCell, , , , objWs) Then rTempCell = .sCustName 'наименование заказчика
        If FindCellRight("ИНН", rTempCell, , , , objWs) Then rTempCell = .sCustINN 'ИНН заказчика
        If FindCellRight("Адрес", rTempCell, , , , objWs) Then rTempCell = .sCustAdress 'адрес заказчика
    End With
    
    LoadNewTempHead = True: Set rTempCell = Nothing: Set objWorkBook = Nothing
End Function
'###################################################################
'функция возвращает количество строк "подбородка" текущего типа протокола
Function GetProtTypeBeardCount(sBaseDir As String, Optional iProtType As Integer) As Integer
    GetProtTypeBeardCount = -1 'по умолчанию
    If iProtType = 0 Then iProtType = GetCurrProtType 'получить текущий тип файла протокола
    
    Dim sTempInstrSym As String, sTempFileName As String 'текущий тип протокола
    sTempInstrSym = "pr_": If iProtType = 2 Then sTempInstrSym = "prc_"
    If iProtType = 3 Then sTempInstrSym = "prm_"
    
    sTempFileName = Dir(sBaseDir & sTempInstrSym & "*" & "temp" & "*"): If sTempFileName = "" Then Exit Function 'если не найден шаблон
    sTempFileName = sBaseDir & sTempFileName 'полный путь к файлу-шапке шаблона
    
    Dim objWorkBook As Object, iRowEnd As Integer
    Set objWorkBook = Application.Workbooks.Open(sTempFileName, , True) 'открыть книгу в режиме чтения
    
    iRowEnd = objWorkBook.ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1 'последняя заполненная строка снизу + строка ниже - должность
    GetProtTypeBeardCount = iRowEnd - iHeadProrCnt - 1 'количество строк "подбородка"
    
    objWorkBook.Close: Set objWorkBook = Nothing
End Function
'###################################################################
'функция возвращает текущий тип протокола (поверка, калибровка или измерение)
Function GetCurrProtType() As Integer
    GetCurrProtType = -1 'по умолчанию
    
    With ActiveWorkbook
        If InStr(.name, "pr_") > 0 Then GetCurrProtType = 1 'поверка
        If InStr(.name, "prc_") > 0 Then GetCurrProtType = 2 'калибровка
        If InStr(.name, "prm_") > 0 Then GetCurrProtType = 3 'измерение
    End With
End Function

'###################################################################
'процедура заменяет на текущем листе значения
Sub ReplaceData(Optional iConverType As Integer = -1, Optional iSourceType As Integer = -1)

    If iConverType = -1 Or iSourceType = -1 Then Exit Sub 'защита от некорректного изменения
    
    Dim sCurrFileType As String, sNewFileType As String, bRowFrom As Byte, bRowTo As Byte
    sCurrFileType = "поверки": If iSourceType = 2 Then sCurrFileType = "калибровки"
    If iSourceType = 3 Then sCurrFileType = "измерений"
    
    sNewFileType = "поверки": bRowFrom = 1: bRowTo = 2
    If iConverType = 2 Then sNewFileType = "калибровки": bRowFrom = 2: bRowTo = 1 'протокол калибровки
    If iConverType = 3 Then sNewFileType = "измерений": bRowFrom = 2: bRowTo = 1 'протокол измерений

    ActiveSheet.Cells.Replace sCurrFileType, sNewFileType 'заменить все значения
    
    Dim rTempCell As Range, i As Byte, sSearchStr As String, sReplaceStr As String
    For i = 1 To 4 'заменить специальные символы
        sSearchStr = ThisWorkbook.Worksheets(1).Cells(bRowFrom, i + 1) 'значение для поиска на листе
        sReplaceStr = ThisWorkbook.Worksheets(1).Cells(bRowTo, i + 1) 'значение для замены
        
        If i < 3 Then
            ActiveSheet.Cells.Replace sSearchStr, sReplaceStr
            
        Else ' заменить ячейку с индексом
            Do While FindCellRight(sSearchStr, rTempCell, , True) = True
                ThisWorkbook.Worksheets(1).Cells(bRowTo, i + 1).Copy
                rTempCell.PasteSpecial xlPasteAllExceptBorders
                
                If rTempCell.Column > 1 Then rTempCell.HorizontalAlignment = xlCenter
            Loop
        End If
    Next i
    
    Application.CutCopyMode = False: Set rTempCell = Nothing
End Sub


    Private Function FifNumByReference( _
        fifNum As String _
        ) As String
        
        FifNumByReference = fifNum
         
         
    
    End Function
'###################################################################
'функция возвращает имя файла шаблона / тела шаблона для дальнейшей загрузки
Function TemplateFileName( _
    dbinstrumentsPath As String, _
    myMi As MeasInstrument, _
    fileExt As String _
    ) As String
    
    TemplateFileName = vbNullString
    
    Dim currentTemplateFileName As String, _
        currentInstrumentDir As String, _
        fifRegNum As String
        
    fifRegNum = myMi.sFif
    If myMi.sRef <> "nodata" Then _
        fifRegNum = myMi.sRef
    
    If InStr(fifRegNum, "-") = 0 Then 'заполнение номера в фиф из АРШИНА
        
        If Len(fifRegNum) >= 8 Then
            fifRegNum = Left(fifRegNum, 8)
            fifRegNum = Replace(fifRegNum, ".", "-")
        End If
    End If
    
    currentInstrumentDir = Dir(dbinstrumentsPath & Application.PathSeparator & "*" & fifRegNum & "*", vbDirectory)  'поиск каталога по номеру ФИФ в БД
    
    Dim refPath As String
    refPath = fso.BuildPath(dbinstrumentsPath, currentInstrumentDir)
    refPath = fso.BuildPath(refPath, REFERENCE_FIFNUM_FILENAME)
     
    If fso.FileExists(refPath) Then _
        fifRegNum = Base.ContentFromFile(refPath)
    
    currentInstrumentDir = Dir(dbinstrumentsPath & Application.PathSeparator & "*" & fifRegNum & "*", vbDirectory)   'повторный поиск каталога по номеру ФИФ в БД
    currentTemplateFileName = Dir(dbinstrumentsPath & Application.PathSeparator & currentInstrumentDir & "\*" & fifRegNum & fileExt)  'поиск любого протокола в каталоге, найденном выше
    
    Dim isSeveralTemplates As Boolean
    Do While currentTemplateFileName <> vbNullString
    
        If Not currentTemplateFileName = vbNullString And _
            Not currentTemplateFileName Like "fif_" & "*" Then
            
            isSeveralTemplates = True
            Exit Do
'            currentTemplateFileName = vbNullString
            
        End If
    
        currentTemplateFileName = Dir
    
    Loop
    currentTemplateFileName = Dir(dbinstrumentsPath & Application.PathSeparator & currentInstrumentDir & "\*" & fifRegNum & fileExt)  'поиск любого протокола
    
    If currentTemplateFileName <> vbNullString Then _
        currentTemplateFileName = dbinstrumentsPath & Application.PathSeparator & currentInstrumentDir & "\" & currentTemplateFileName
    
    If isSeveralTemplates Then    'вызов меню выбора файла, если файлов несколько
        
        With myMultiSel
            .sMSBaseDir = dbinstrumentsPath & Application.PathSeparator
            .sMSfifNum = fifRegNum
            .sMSFullFileName = vbNullString
            .sMSType = myMi.sType
        End With
        
        Z_UF_MultiSelect_Load.Show 'форма выбора файла по имени

        currentTemplateFileName = myMultiSel.sMSFullFileName
    End If
    
    TemplateFileName = currentTemplateFileName
End Function
'###################################################################
'процедура проверяет наличие в каталоге текущей книги дополнительного файла WD и дополняет его свойства, если он имеется
Sub FillThisPathDocTempName(myInstr As MeasInstrument)

    Dim objWdApp As Object, objWdDoc As Object, sTempDocExt As String
    Set objWdApp = CreateObject("Word.Application"): Set objWdDoc = objWdApp.Documents.Add(, , , True) 'создать новый документ
    
    sTempDocExt = Dir(ActiveWorkbook.path & "\*.doc*") 'проверить документ на ДИР, если есть - выделить расширение
    sTempDocExt = GetExt(sTempDocExt): If sTempDocExt = "" Then sTempDocExt = ".docx"
    
    With objWdDoc
        .BuiltinDocumentProperties("Keywords") = myInstr.sFif
        .BuiltinDocumentProperties("Comments") = myInstr.sMetodic 'полное наименование типа СИ
        
        Dim sDocNewName As String, sTempStr As String
        With myInstr
            sDocNewName = "body_": sTempStr = .sType: If .sModification <> "" Then sTempStr = sTempStr & " #!" & .sModification
            sDocNewName = sDocNewName & .sFif & " " & sTempStr
            sDocNewName = ReturnNotExistingName(ActiveWorkbook.path & "\", sDocNewName, sTempDocExt)
        End With
        
        .SaveAs sDocNewName
    End With
    
    objWdApp.Quit: Set objWdApp = Nothing: Set objWdDoc = Nothing
End Sub
Sub Commit_factory_number( _
    ) 'зафиксировать заводской номер СИ в текущем протоколе
    
    Dim serialNum As String
    serialNum = SerialNumFromSheet
    
    If serialNum = vbNullString Then _
        Exit Sub
    
    Dim currentComment() As String
    currentComment = MiComment
        
    Dim arrayIndex As Integer
    arrayIndex = LBound(currentComment)
    
    Dim currentMiName As String
    currentMiName = ActiveWorkbook.ActiveSheet.name
    
    Do While arrayIndex <= UBound(currentComment)
    
        If currentMiName = currentComment(arrayIndex) Then _
            Exit Do
            
        arrayIndex = arrayIndex + 1
    Loop
    
    Dim NewComment As String
    If arrayIndex = UBound(currentComment) Then
        NewComment = Join(currentComment, " -- ") & " -- " & serialNum
        
    Else
        currentComment(UBound(currentComment)) = serialNum
        NewComment = Join(currentComment, " -- ")
    End If
               
    SetBuiltInProperty "Comments", NewComment  ': If InStr(sSaveName, "prm_") > 0 Then sTempStr = serialNum
End Sub
    Private Function SerialNumFromSheet( _
        ) As String

        Dim serialNumCell As Range
        If FindCellRight("Заводской / серийный номер:", serialNumCell) Then
                
            Dim serialNum As String
            serialNum = CStr(serialNumCell)
            
            serialNum = Replace(serialNum, "№", vbNullString)
            serialNum = Trim(serialNum)
            
            Set serialNumCell = Nothing
            
        End If
        
        SerialNumFromSheet = serialNum
    End Function
    Private Function MiComment( _
        ) As String()
        
        Dim currentComment As String
        currentComment = GetBuiltInProperty("Comments")
        DeleteSpaceStEnd currentComment
        
        MiComment = Split(currentComment, " -- ")
    End Function
        
'###################################################################
'процедура загружает случайную строку подсказки из файла
Function LoadHelpString(sStartDir As String) As String
    
    Dim sArrHelp() As String, myBase As New Z_clsmBase, iHelpInd As Integer, sStr As String
    ReDim sArrHelp(0): If FileExist(sStartDir, "info.hpDb") = False Then Exit Function 'если файл производственного календаря опознан
    sArrHelp = myBase.GetArrFF(sStartDir, "info.hpDb") 'получить массив
    
    Randomize: iHelpInd = Rnd * UBound(sArrHelp): sStr = "#" & iHelpInd + 1 & ". " & sArrHelp(iHelpInd)
    LoadHelpString = DeleteSpaceStEnd(sStr): Set myBase = Nothing
End Function

'###################################################################
'Процедура заполняет блок условий поверки
Private Sub FillNormalCondition( _
    ByRef objZ_clsmSearch As Z_clsmSearch _
    )
    
    Dim sCurrCond As String, sgTemp As Single, sgHum As Single, sgPress As Single, sgBkg As Single
    
    
    
    'todo: [+] FillNormalCondition -- ОТВЯЗАТЬ z_clsmSearch, сделать Cache
    sCurrCond = objZ_clsmSearch.normalCondition
    
    If sCurrCond <> "недоступно" Then 'извлечь сохранённые свойства из настроек
        Dim sArrTemp() As String, lb As Byte
        sArrTemp = Split(sCurrCond, InStrDelimiter): lb = LBound(sArrTemp)
        
        If Now - CDate(sArrTemp(lb)) < 0.21 Then _
            sgTemp = sArrTemp(lb + 1): sgHum = sArrTemp(lb + 2): _
            sgPress = sArrTemp(lb + 3): sgBkg = sArrTemp(lb + 4) 'если давность значений меньше 5 часов
    End If
    
    If sgTemp = 0 Then 'свойства не были извлечены из настроек
        sgTemp = Format(GaussRnd(22.7, 0.4), "0.0") 'получить значение температуры
        
        Select Case Month(Date) 'получить значение влажности от месяца в году
            Case 1, 2, 3, 12: sgHum = Format(GaussRnd(43.5, 0.7), "0.0") 'январь, февраль, март, декабрь
            Case 4, 5, 6, 10, 11: sgHum = Format(GaussRnd(46, 1.2), "0.0") ' апрель, май, июнь, октябрь, ноябрь
            Case Else: sgHum = Format(GaussRnd(55, 2), "0.0") ' июль, август, сентябрь
        End Select
        
        sgPress = Format(GaussRnd(101, 0.5), "0.0") 'получить значение атмосферного давления
        sgBkg = Format(GaussRnd(0.14, 0.007), "0.00") 'получить значение фона
        
        sCurrCond = Now & InStrDelimiter & sgTemp & InStrDelimiter & sgHum & _
            InStrDelimiter & sgPress & InStrDelimiter & sgBkg
            
        'todo: [+] FillNormalCondition -- ОТВЯЗАТЬ z_clsmSearch, сделать Cache
        objZ_clsmSearch.normalCondition = sCurrCond 'передать свойство в класс
    End If
    
    Dim rMyCell As Range 'передать извлечённые свойства в ячейки
    If FindCellRight("температура", rMyCell) Then _
        rMyCell = sgTemp: rMyCell.Offset(1, 0) = sgHum: rMyCell.Offset(2, 0) = sgPress: rMyCell.Offset(3, 0) = sgBkg
    
    Set rMyCell = Nothing
End Sub
'###################################################################
'Фнукция возвращает случайное значение нормального распределения
Function GaussRnd( _
    sgMathExp As Double, _
    sgSigma As Double _
    ) As Double
    
    Dim r1 As Double, _
        r2 As Double, _
        pi As Double
        
    pi = WorksheetFunction.pi
    
    Randomize
    r1 = Rnd
    r2 = Rnd
    
    GaussRnd = sgMathExp + sgSigma * Cos(2 * pi * r1) * Sqr(-2 * Log(r2))
End Function

Private Sub testgauss()
    
    Dim myVal As Double, _
        mySig As Double
        
    myVal = 7
    
    Dim i As Byte
    For i = 0 To 200
        Debug.Print GaussRnd(myVal, 0.1 * myVal) + GaussRnd(myVal, 0.2 * myVal) + _
                    GaussRnd(myVal, 0.25 * myVal)
        
    Next
    
End Sub


'###################################################################
'функция функция корректно отделяет количество символов в наименовании
Private Function TrueNameLength(sStr As String, iStrMaxLen As Integer, _
    Optional bolRightPart As Boolean, Optional bolComma As Boolean)

    Dim sLeftStr As String, sRightStr As String
    If Len(sStr) <= iStrMaxLen Then 'если наименование полностью умещается
        TrueNameLength = sStr: If bolRightPart Then TrueNameLength = sRightStr
        Exit Function 'правая часть - пустая строка
    End If

    sLeftStr = Left(sStr, iStrMaxLen): sRightStr = Right(sStr, Len(sStr) - iStrMaxLen) 'разбиение строки на максимальное количество символов

    Dim sTypeInStr As String, bInputType As Byte, sNameInStr As String
    bInputType = InStr(sStr, "; тип")
    If bInputType > 0 Then
        sTypeInStr = Right(sStr, Len(sStr) - bInputType - 5)     'получить чистый тип СИ из наименования
        sNameInStr = Left(sStr, Len(sStr) - Len(sTypeInStr) - 6)
    End If

    If Right(sLeftStr, 1) <> "" And Left(sRightStr, 1) <> "" Then 'строка была разбита надвое

        Do While Right(sLeftStr, 1) <> " " 'выполнять, пока не будет получено корректное урезание
            sLeftStr = Left(sLeftStr, Len(sLeftStr) - 1)
        Loop
        DeleteSpaceStEnd sLeftStr

        Select Case Right(sLeftStr, 2)
            Case " и", " в", " а", "ПО"
                sLeftStr = Left(sLeftStr, Len(sLeftStr) - 2) 'чтобы не обрывать фразу на союзе
            Case " №"
                sLeftStr = Left(sLeftStr, Len(sLeftStr) - 6) 'чтобы не обрывать фразу на рег.№
        End Select

        If InStr(sLeftStr, "; тип ") > 0 And InStr(sLeftStr, sTypeInStr) = 0 Then
            sLeftStr = sNameInStr: sRightStr = "тип " & sTypeInStr
        Else
            If Right(sLeftStr, 5) = "; тип" Then sLeftStr = Left(sLeftStr, Len(sLeftStr) - 5) 'чтобы не обрывать фразу на типе

            If Right(sLeftStr, 11) = "модификация" Then sLeftStr = Left(sLeftStr, Len(sLeftStr) - 11)
            DeleteSpaceStEnd sLeftStr

            sRightStr = Right(sStr, Len(sStr) - Len(sLeftStr))
        End If

        DeleteSpaceStEnd sRightStr
        If bolComma And sRightStr <> "" Then sRightStr = sRightStr & ";               "
    End If
    DeleteSpaceStEnd sLeftStr: TrueNameLength = sLeftStr: If bolRightPart Then TrueNameLength = sRightStr
End Function


'1-------###################################################-------1
'функция возвращает строку начала блока экспорта данных
Private Function FindSvidExportRow(myWs As Worksheet) As Integer
    FindSvidExportRow = -1 'по умолчанию блок не найден

    With myWs
        Dim sTempStr As String, i As Byte
        sTempStr = .Cells(.Rows.count, 1).End(xlUp).text 'содержимое последней заполненной снизу ячейки

        If sTempStr = "svid export" Then _
            FindSvidExportRow = .Cells(.Rows.count, 1).End(xlUp).Row 'передать номер строки

        If FindSvidExportRow = -1 Then 'наличие блока старого файла протокола не было опознано
            i = 5 'начальная строка для поиска

            Do While i < 30
                sTempStr = .Cells(i, 1).text
                If sTempStr = "svid export" Then _
                    FindSvidExportRow = .Cells(i, 1).Row: Exit Function 'передать номер строки
                i = i + 1
            Loop

        End If
    End With
End Function


