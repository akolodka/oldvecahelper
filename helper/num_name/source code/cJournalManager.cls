VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJournalManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cJournalManager (Class Module)
' Назначение: Управление записями журнала регистрации поверок
' Автор: akolodka
' Date: 02.02.2023 22:43
' ------------------------------------------------------
Option Explicit
Public Enum ERegistrationTypes
    
    newItem
    fgisNumber
    
End Enum

Const DELIMITER = "   "

Private fso As New FileSystemObject, _
        keys As New cJournalColumnKeys, _
        wbJournal As Workbook

Const MSG_MAXLENGTH = 18
Const FRIS_LINK_PREFIX = "https://fgis.gost.ru/fundmetrology/cm/results/1-"

Private err As New cErrorMessages

' ----------------------------------------------------------------
' Наименование: RegisterNewItem (Friend Function)
' Назначение: добавить новый элемент в журнал
'    параметр data: блок данных для регистрации
' Возвращаемый тип: cJournalItem
' Дата: 04.02.2023 21:48
' ----------------------------------------------------------------
Friend Function RegisterNewItem( _
    data As cJournalItem, _
    ByRef status As Boolean _
    ) As cJournalItem
    
    status = False
    Set RegisterNewItem = data
    ' ----------------------------------------------------------------
    If Not isExists Then _
        Handler.Notify (err.journalNotFinded): _
        Exit Function
    ' ----------------------------------------------------------------
    Application.ScreenUpdating = False
    Set wbJournal = Application.Workbooks.Open(Config.journalPath)
    
    'Set wbJournal = GetObject(Config.journalPath)
    ' ----------------------------------------------------------------
    If Not isAvailable Then
        
        Dim assholeName As String: assholeName = FileOwner
        InterruptLoad
        
        Handler.Notify err.journalBusy(assholeName)
        Exit Function
        
    End If
    ' ----------------------------------------------------------------
    If Not IsColumnsCorrect(registrationType:=newItem) Then _
        InterruptLoad: _
        Exit Function
    ' ----------------------------------------------------------------
    If Not IsDataCorrect(data, registrationType:=newItem) Then _
        InterruptLoad: _
        Exit Function
    ' ----------------------------------------------------------------
    If IsRegisteredYet(data) Then
    
        If Not IsRegisterAnyway(data) Then _
            UnloadSafety: Exit Function
    
    End If
    ' ----------------------------------------------------------------
    Dim result As cJournalItem
    Set result = RegisterNewItemInJournal(data)
    
    Set RegisterNewItem = result
    UnloadSafety
    
    status = True
    
    'Debug.Print result.protocolFullNumber
    
End Function
    ' ----------------------------------------------------------------
    ' Дата: 10.04.2023 10:18
    ' ----------------------------------------------------------------
    Private Function FileOwner( _
        ) As String
        
        Dim workPath As String: workPath = wbJournal.path
        Dim mask As String: mask = wbJournal.name
        ' ----------------------------------------------------------------
        Dim targetFileName As String
        targetFileName = Dir(workPath & Application.PathSeparator & "*~*" & mask & "*", vbHidden)
        ' ----------------------------------------------------------------
        Dim sourcePath As String: sourcePath = fso.BuildPath(workPath, targetFileName)
        Dim destinationPath As String: destinationPath = fso.BuildPath(Base.configDataPath, targetFileName)
        
        If targetFileName <> vbNullString Then _
            fso.CopyFile sourcePath, destinationPath
        
        Dim fileData As String: fileData = Base.ContentFromFile(destinationPath)
        FileOwner = OwnerFromContent(fileData)
        ' ----------------------------------------------------------------
        fso.DeleteFile destinationPath
        
    End Function
        ' ----------------------------------------------------------------
        ' Дата: 10.04.2023 11:14
        ' ----------------------------------------------------------------
        Private Function OwnerFromContent( _
            fileData As String _
            ) As String
            
            OwnerFromContent = Base.defaultValuePrinted
            If fileData = Base.defaultValue Then _
                Exit Function
            
            Dim newArr() As String
            If InStr(fileData, ",") Then 'файл открыт в LibreOffice
                
                newArr = Split(fileData, ",")
                OwnerFromContent = newArr(LBound(newArr) + 2)
                
            Else 'файл открыт в Microsoft Office

                newArr = Split(fileData, " ")
                OwnerFromContent = Mid(newArr(LBound(newArr)), 2)
                
            End If

        End Function

    
    Private Function isExists( _
        ) As Boolean
                
        isExists = fso.FileExists(Config.journalPath)

    End Function
    Private Function isAvailable( _
        ) As Boolean
        
        isAvailable = wbJournal.ReadOnly = False
        
    End Function
    ' ----------------------------------------------------------------
    ' Наименование: IsColumnsCorrect (Private Function)
    ' Назначение: достаточность журнала для записи в него данных
    '    параметр registrationType:
    ' Возвращаемый тип: Boolean
    ' Дата: 05.02.2023 12:16
    ' ----------------------------------------------------------------
    Private Function IsColumnsCorrect( _
        registrationType As ERegistrationTypes _
        ) As Boolean
        
        IsColumnsCorrect = False
        
        If wbJournal Is Nothing Then _
            Exit Function
        
        Select Case True
                
            Case registrationType = newItem
                IsColumnsCorrect = IsCorrectForNewItem(JournalHeaderRange)
                
            Case registrationType = fgisNumber
                IsColumnsCorrect = IsCorrectForfgisNumber(JournalHeaderRange)
                
        End Select
        
        IsColumnsCorrect = True
        
    End Function
        Private Function JournalHeaderRange( _
            ) As Range
            
            Set JournalHeaderRange = wbJournal.ActiveSheet.Range("1:1")
        
        End Function
        Private Function IsCorrectForNewItem( _
            searchRange As Range _
            ) As Boolean
            
            IsCorrectForNewItem = False
    
            If Not IsFinded(searchRange, keys.typeInstrument) Then _
                Handler.Notify (err.columnTypeNotFinded): _
                Exit Function
    
            If Not IsFinded(searchRange, keys.serial) Then _
                Handler.Notify (err.columnSerialNotFinded): _
                Exit Function
                
            If Not IsFinded(searchRange, keys.ProtocolNumber) Then _
                Handler.Notify (err.columnProtocolNumberNotFinded): _
                Exit Function
            
            IsCorrectForNewItem = True
            
        End Function
            Private Function IsFinded( _
                searchRange As Range, _
                findKey As String _
                ) As Boolean
                
                IsFinded = Not FindCellByKey(searchRange, findKey) Is Nothing
                
            End Function
                Private Function FindCellByKey( _
                    searchRange As Range, _
                    findKey As String _
                    ) As Range
                    
                    Dim result As Range
                    Set result = searchRange.Find(findKey, LookAt:=xlPart, LookIn:=xlValues)
                    
                    Set FindCellByKey = result

                End Function
        Private Function IsCorrectForfgisNumber( _
            searchRange As Range _
            ) As Boolean
            
            IsCorrectForfgisNumber = False
        
            If Not IsFinded(searchRange, keys.ProtocolNumber) Then _
                Handler.Notify (err.columnProtocolNumberNotFinded): _
                Exit Function
                
            If Not IsFinded(searchRange, keys.fgisDocumentNumber) Then _
                Handler.Notify (err.columnFgisNumberNotFinded): _
                Exit Function
            
            If Not IsFinded(searchRange, keys.fgisLink) Then _
                Handler.Notify (err.columnFgisLinkNotFinded): _
                Exit Function
            
            IsCorrectForfgisNumber = True
            
        End Function
    ' ----------------------------------------------------------------
    Private Sub InterruptLoad()
    
        wbJournal.Close False
        Set wbJournal = Nothing
        Application.ScreenUpdating = True
    
    End Sub
    ' ----------------------------------------------------------------
    ' Наименование: IsDataCorrect (Private Function)
    ' Назначение: Достаточность входныех данных для записи в журнал
    '    параметр data:
    '    параметр registrationType:
    ' Возвращаемый тип: Boolean
    ' Дата: 05.02.2023 12:15
    ' ----------------------------------------------------------------
    Private Function IsDataCorrect( _
        data As cJournalItem, _
        registrationType As ERegistrationTypes _
        ) As Boolean

        Select Case True
            
            Case registrationType = newItem
                IsDataCorrect = IsDataForNewItem(data)
                
            Case registrationType = fgisNumber
                IsDataCorrect = IsDataForFgis(data)
        
        End Select
         
    End Function
        Private Function IsDataForNewItem( _
            data As cJournalItem _
            ) As Boolean
            
            IsDataForNewItem = False
                        
            If data.typeInstrument = Base.defaultValue Then _
                Handler.Notify (err.incorrectTypeInstrument): _
                Exit Function
            
            If data.factoryNumber = Base.defaultValue Then _
                Handler.Notify (err.incorrectFactoryNumber): _
                Exit Function
                
            IsDataForNewItem = True
            
        End Function
        Private Function IsDataForFgis( _
            data As cJournalItem _
            ) As Boolean
            
            IsDataForFgis = False
            
            If data.protocolFullNumber = Base.defaultValue Then _
                Handler.Notify (err.incorrectProtocolNumber): _
                Exit Function
                
            If data.fgisDocNumber = Base.defaultValue Then _
                Handler.Notify (err.incorrectFgisDocumentNumber): _
                Exit Function
            
            If data.fgisLinkNumber = Base.defaultValue Then _
                Handler.Notify (err.incorrectFgisLinkNumber): _
                Exit Function
            
            If data.dateVerifying = Base.defaultValue Then _
                Handler.Notify (err.incorrectDateVerifying): _
                Exit Function
                
            IsDataForFgis = True
            
        End Function
    ' ----------------------------------------------------------------
    ' Наименование: IsRegisteredYet (Private Function)
    ' Назначение: Проверка на предмет наличия записи в журнале
    '    параметр data:
    ' Возвращаемый тип: Boolean
    ' Дата: 05.02.2023 12:14
    ' ----------------------------------------------------------------
    Private Function IsRegisteredYet( _
        data As cJournalItem _
        ) As Boolean
        
        Dim targetColumn As Integer
        targetColumn = ColumnNumber(keys.serial)
        
        Dim searchRange As Range
        Set searchRange = wbJournal.ActiveSheet.Columns(targetColumn)
        
        Dim result As Range
        Set result = searchRange.Find(data.factoryNumber, LookAt:=xlWhole, LookIn:=xlValues)
        
        IsRegisteredYet = Not result Is Nothing
        
    End Function
        Private Function ColumnNumber( _
            key As String _
            ) As Integer
            
            Dim result As Range
            Set result = FindCellByKey( _
                searchRange:=JournalHeaderRange, _
                findKey:=key _
                )
            
            If Not result Is Nothing Then _
                ColumnNumber = result.Column
            
        End Function
    ' ----------------------------------------------------------------
    ' Наименование: IsRegisterAnyway (Private Function)
    ' Назначение: подтверждение пользователем необходимости регистрации
    '    параметр data:
    ' Возвращаемый тип: Boolean
    ' Дата: 05.02.2023 12:17
    ' ----------------------------------------------------------------
    Private Function IsRegisterAnyway( _
        data As cJournalItem _
        ) As Boolean
        
        IsRegisterAnyway = False
        
        Dim existingData As New Collection
        Set existingData = exData(data)
        
        If existingData.count > vbEmpty Then
                
            If Not IsUserConfirm(data, existingData) Then _
                Handler.Notify "Прервано пользователем", msgType:=msgboxMessage: _
                Exit Function

        End If
        
        IsRegisterAnyway = True
        
    End Function
        Private Function exData( _
            data As cJournalItem _
            ) As Collection
            
            Dim targetColumn As Integer
            targetColumn = ColumnNumber(keys.serial)
            
            Dim searchRange As Range
            Set searchRange = wbJournal.ActiveSheet.Columns(targetColumn)
            
            Dim searchResult As Range
            Set searchResult = searchRange.Find(data.factoryNumber, LookAt:=xlWhole, LookIn:=xlValues)
            
            If searchResult Is Nothing Then: _
                Exit Function 'на случай чуда
                
            Dim firstResult As Range
            Set firstResult = searchResult
            
            Dim resultData As New Collection, _
                matchingData As cJournalItem
                
            Set matchingData = MatchData(searchResult)
            
            If IsMatched(data, matchingData) Then _
                    resultData.Add matchingData
            
            Do While Not searchResult Is Nothing
            
                Set searchResult = searchRange.Find(data.factoryNumber, After:=searchResult, LookAt:=xlWhole, LookIn:=xlValues)
                If searchResult.address = firstResult.address Then _
                    Exit Do
                
                Set matchingData = MatchData(searchResult)
                
                If IsMatched(data, matchingData) Then _
                    resultData.Add matchingData
            Loop

            Set exData = resultData
            
        End Function
            Private Function MatchData( _
                searchResult As Range _
                ) As cJournalItem
                                
                Dim result As New cJournalItem
                
                result.typeInstrument = JournalItemData(searchResult, keys.typeInstrument)
                result.factoryNumber = JournalItemData(searchResult, keys.serial)
                result.protocolFullNumber = JournalItemData(searchResult, keys.ProtocolNumber)
                result.fgisDocNumber = JournalItemData(searchResult, keys.fgisDocumentNumber)
                result.organisation = JournalItemData(searchResult, keys.organisation)
                result.employees = JournalItemData(searchResult, keys.employees)
                
                Set MatchData = result
                
            End Function
                Private Function JournalItemData( _
                    searchResult As Range, _
                    key As String _
                    ) As String
                    
                    Dim targetColumn As Integer
                    targetColumn = ColumnNumber(key)
    
                    Dim jrCellData As Range
                    Set jrCellData = wbJournal.ActiveSheet.Cells(searchResult.Row, targetColumn)
                    
                    JournalItemData = Trim(jrCellData)
                    
                End Function
                
            ' ----------------------------------------------------------------
            ' Дата: 18.02.2023 16:01
            ' Назначение: дополнительное сравнение по типу СИ, зав. № уже совпадает
            '    параметр data:
            '    параметр matchingData:
            ' Возвращаемый тип: Boolean
            ' ----------------------------------------------------------------
            Private Function IsMatched( _
                data As cJournalItem, _
                matchingData As cJournalItem _
                ) As Boolean
                
                IsMatched = matchingData.typeInstrument = data.typeInstrument
                
            End Function
'        Public Sub TESTTE()
'
'            Dim test As New cJournalItem
'            Dim ex As New Collection
'            ex.Add test
'            ex.Add test
'
'            IsUserConfirm test, ex
'        End Sub

        ' ----------------------------------------------------------------
        ' Наименование: IsUserConfirm (Private Function)
        ' Назначение: взаимодействие с пользователем для подтверждения регистрации
        '    параметр exData:
        ' Возвращаемый тип: Boolean
        ' Дата: 05.02.2023 13:04
        ' ----------------------------------------------------------------
        Private Function IsUserConfirm( _
            data As cJournalItem, _
            exData As Collection _
            ) As Boolean
            
            IsUserConfirm = False
        
            Dim msg As String '                                                                   " зарегистрировано ранее:"
            msg = "Средство измерений " & data.typeInstrument & " зав. № " & data.factoryNumber & " уже есть в журнале:" & vbNewLine & vbNewLine
            msg = msg & "№" & DELIMITER & "Протокол" & DELIMITER & DELIMITER & DELIMITER & DELIMITER & "Документ Аршин               " & DELIMITER & DELIMITER & "Исполнители " '& DELIMITER & DELIMITER & "Заказчик" & vbNewLine
          '  msg = msg & "-----------------------------------------------------------------------------------------------------"
            msg = msg & "--------------------------------------------------------------------------------------"
            
            Dim item As cJournalItem, _
                i As Integer
                
            For Each item In exData
            
                i = i + 1
                msg = msg & vbNewLine & SpacedIndex(i) & DELIMITER & item.protocolFullNumber & DELIMITER & item.fgisDocNumber & DELIMITER & DELIMITER & CuttedEmployees(item.employees) ' & DELIMITER & DELIMITER & item.organisation
                
            Next
            
            msg = msg & vbNewLine & "--------------------------------------------------------------------------------------"
            msg = msg & vbNewLine & vbNewLine & "Зарегистрировать новую запись?"
            
            Dim ans As Integer
            ans = Handler.ask(msg)
            
            IsUserConfirm = ans = vbYes
            
        End Function
            ' ----------------------------------------------------------------
            ' Наименование: SpacedIndex (Private Function)
            ' Назначение: Добавление пробелов, чтобы красиво выглядело
            '    параметр index:
            ' Возвращаемый тип: String
            ' Дата: 15.02.2023 22:28
            ' ----------------------------------------------------------------
            Private Function SpacedIndex( _
                index As Integer _
                ) As String
                
                Select Case True
                
                    Case Len(index) = 1
                        SpacedIndex = index & "   "
                    
                    Case Len(index) = 2
                        SpacedIndex = index & "  "
                        
                    Case Len(index) = 3
                        SpacedIndex = index & " "
                        
                    Case Len(index) = 0
                        SpacedIndex = "Fkn slave"
                        
                End Select
            
            End Function
            ' ----------------------------------------------------------------
            ' Наименование: CuttedEmployees (Private Function)
            ' Назначение: Обрезка строки, чтобы не было переноса текста
            '    параметр inputData:
            ' Возвращаемый тип: String
            ' Дата: 15.02.2023 22:28
            ' ----------------------------------------------------------------
            Private Function CuttedEmployees( _
                inputData As String _
                ) As String
                
                Dim result As String
                result = Left(inputData, MSG_MAXLENGTH)
                
                If Len(result) < Len(inputData) Then _
                    result = result & "..."
                
                CuttedEmployees = result
            
            End Function
        Private Function RegisterNewItemInJournal( _
            data As cJournalItem _
            ) As cJournalItem
            
            Set RegisterNewItemInJournal = data
            
            Dim newItemRow As Integer
            newItemRow = FreeJournalRow 'работает
            
            Dim result As New cJournalItem
            Set result = FillNewItemData(newItemRow, data)
            
            result.protocolFullNumber = wbJournal.ActiveSheet.Cells(newItemRow, ColumnNumber(keys.ProtocolNumber))
            
            Set RegisterNewItemInJournal = result
        End Function
            Private Function FillNewItemData( _
                rwTarget As Integer, _
                data As cJournalItem _
                ) As cJournalItem
                
                FillByColumn rwTarget, ColumnNumber(keys.organisation), data.organisation
                FillByColumn rwTarget, ColumnNumber(keys.typeInstrument), data.typeInstrument
                FillByColumn rwTarget, ColumnNumber(keys.serial), data.factoryNumber
                FillByColumn rwTarget, ColumnNumber(keys.employees), data.employees
 
            End Function
                Private Sub FillByColumn( _
                    rwTarget As Integer, _
                    colTarget As Integer, _
                    value As String _
                    )
                    
                    wbJournal.ActiveSheet.Cells(rwTarget, colTarget) = Trim(value)
                    
                End Sub
            ' ----------------------------------------------------------------
            ' Наименование: FreeJournalRow (Private Function)
            ' Назначение: Последняя заполненная строка в журнале
            ' Возвращаемый тип: Integer
            ' Дата: 16.02.2023 22:44
            ' ----------------------------------------------------------------
            Private Function FreeJournalRow( _
                ) As Integer
                
                Dim rwOrgranisation As Integer
                rwOrgranisation = LastRowByColumn(keys.organisation)
                
                Dim rwType As Integer
                rwType = LastRowByColumn(keys.typeInstrument)
                
                Dim rwSerial As Integer
                rwSerial = LastRowByColumn(keys.serial)
                
                FreeJournalRow = WorksheetFunction.Max(rwOrgranisation, rwType, rwSerial) + 1

'                Debug.Print rwOrgranisation, rwType, rwSerial, , FreeJournalRow
                
            End Function
                Private Function LastRowByColumn( _
                    findKey As String _
                    ) As Integer
                    
                    LastRowByColumn = -1
                    
                    Dim colTarget As Integer
                    colTarget = ColumnNumber(findKey)
                                           
                    Dim resultCell As Range
                    Set resultCell = wbJournal.ActiveSheet.Cells(wbJournal.ActiveSheet.Rows.count, colTarget).End(xlUp)
                    
                    Dim result As Integer
                    result = RowFromAddress(resultCell.MergeArea.address)
                    
                    LastRowByColumn = result
                    
                End Function
                    Private Function RowFromAddress( _
                        address As String _
                        ) As Integer
            
                        RowFromAddress = Right(address, InStr(StrReverse(address), "$") - 1)
            
                    End Function
    Private Sub UnloadSafety()
        
        If Not wbJournal Is Nothing Then
            
           ' wbJournal.Windows(1).Visible = True
            wbJournal.Close SaveChanges:=True
            
            Set wbJournal = Nothing
            Application.ScreenUpdating = True
            
        End If
        
    End Sub
' ----------------------------------------------------------------
' Наименование: RegisterfgisNumber (Friend Function)
' Назначение: добавить недостающие сведения в журнал (ФГИС)
'    параметр data: блок данных
' Возвращаемый тип: cJournalItem
' Дата: 04.02.2023 21:49
' ----------------------------------------------------------------
Friend Sub RegisterFgisNumber( _
    data As cJournalItem, _
    Optional ByRef status As Boolean _
    )
    
    status = False
    ' ----------------------------------------------------------------
    If Not isExists Then _
        Handler.Notify (err.journalNotFinded): _
        Exit Sub
    ' ----------------------------------------------------------------
    Application.ScreenUpdating = False
    Set wbJournal = Application.Workbooks.Open(Config.journalPath)
    ' ----------------------------------------------------------------
    If Not isAvailable Then
        
        Dim assholeName As String: assholeName = FileOwner
        InterruptLoad
        
        Handler.Notify err.journalBusy(assholeName)
        Exit Sub
        
    End If
    ' ----------------------------------------------------------------
    If Not IsColumnsCorrect(registrationType:=fgisNumber) Then _
        InterruptLoad: _
        Exit Sub
    ' ----------------------------------------------------------------
    If Not IsDataCorrect(data, registrationType:=fgisNumber) Then _
        InterruptLoad: _
        Exit Sub

    status = RegisterFgisInJournal(data)
    UnloadSafety

End Sub
    Private Function RegisterFgisInJournal( _
        data As cJournalItem _
        ) As Boolean
        
        RegisterFgisInJournal = False
        
        Dim rwItem As Integer
        rwItem = ItemRow(data)
        
        If rwItem < vbEmpty Then _
            Exit Function 'если ничего не найдено
        ' ----------------------------------------------------------------
        FillItemfgisNumber rwItem, data
        
        RegisterFgisInJournal = True
    
    End Function
        ' ----------------------------------------------------------------
        ' Дата: 19.02.2023 16:45
        ' Назначение:
        '    параметр data:
        ' Возвращаемый тип: Integer
        ' ----------------------------------------------------------------
        Private Function ItemRow( _
            data As cJournalItem _
            ) As Integer
            
            ItemRow = -1
            
            Dim targetColumn As Integer
            targetColumn = ColumnNumber(keys.ProtocolNumber)
            
            Dim searchRange As Range
            Set searchRange = wbJournal.ActiveSheet.Columns(targetColumn)
            
            Dim searchResult As Range
            Set searchResult = searchRange.Find(data.protocolFullNumber, LookAt:=xlWhole, LookIn:=xlValues)
    
            If searchResult Is Nothing Then: _
                Exit Function 'на случай чуда
                
            ItemRow = searchResult.Row
            
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 19.02.2023 16:56
        ' Назначение:
        '    параметр rwTarget:
        '    параметр data:
        ' ----------------------------------------------------------------
        ' ----------------------------------------------------------------
        ' Дата: 16.04.2023 19:17
        ' ----------------------------------------------------------------
        Private Sub FillItemfgisNumber( _
            rwTarget As Integer, _
            data As cJournalItem _
            )
            
            FillByColumn rwTarget, ColumnNumber(keys.dateVerifying), data.dateVerifying
            
            FillByColumn rwTarget, ColumnNumber(keys.fgisDocumentNumber), data.fgisDocNumber
            FillByColumn rwTarget, ColumnNumber(keys.fgisEntryNumber), data.fgisLinkNumber
            FillByColumn rwTarget, ColumnNumber(keys.fgisLink), FRIS_LINK_PREFIX & data.fgisLinkNumber

        End Sub
