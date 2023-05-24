VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabaseManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: IDatabase (Class Module)
' Назначение:
' Автор: akolodka
' Date: 24.02.2023 20:01
' ------------------------------------------------------
Option Explicit
Option Compare Text

Public Enum EDadaBaseTypes

    organisations
    instruments
    etalons
    persons
    archiveLocal
    
End Enum

Const CACHE_FILEEXT = "dbc"
Const CACHE_DELIMITER = " \\ "
Const RECACHE_INTERVAL = 30 'раз в сколько дней перекешировать базы данны
Const SUBKEY_DELIMITER = "#"

Const KEY_VNIIM = "вниим"
Const RESULT_MAXCOUNT = 4

Const BODY_KEY = "body_"
Const REPLACEMENT_SAFETY = "_"


Private Type TFilterResult
    
    count As Integer
    data() As New cFilterResults
    targetKey As String
    defaultSuggestions As Collection
    keysExistingItems As String
    
End Type

Private Type TDataBase

    type As EDadaBaseTypes
    pathDb As String
    
    cacheData As String
    cacheFileName As String
    cacheFilePath As String
    
    cacheItemCount As Integer
    
End Type

Private ds As New cDatabaseSource, _
        fso As New FileSystemObject, _
        this As TDataBase, _
        ct As New cCharTable
        
Private uv As New cUserFormVisualisator, _
        fr As TFilterResult
        
Private Sub Class_Initialize()
    
    this.pathDb = Base.defaultValue
    this.cacheData = Base.defaultValue
    ReDim fr.data(vbEmpty)
    
End Sub
' ----------------------------------------------------------------
' Дата: 24.03.2023 07:58
' ----------------------------------------------------------------
Public Property Get SearchBackgroundText() As String
    SearchBackgroundText = uv.LabelBackgroundText(this.type)
End Property
' ----------------------------------------------------------------
' Дата: 24.03.2023 08:12
' ----------------------------------------------------------------
Public Property Get LabelInfoText() As String
    LabelInfoText = " Всего " & uv.ItemTextSuffix(CStr(this.cacheItemCount)) & "." '& uv.ShortedString(Config.organisationsPath)
End Property

' ----------------------------------------------------------------
' Дата: 27.03.2023 20:28
' ----------------------------------------------------------------
Public Sub SpecialAction()
    Select Case True
    
        Case this.type = organisations
            SpecialActionOrganisation
        
        'todo: SpecialAction -- для других баз данных
        
    End Select
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 27.03.2023 20:28
    ' ----------------------------------------------------------------
    Private Sub SpecialActionOrganisation( _
        )
        Dim item As cItemOrganisation
        Set item = GetItem(fr.targetKey)
                
        Dim cm As New cCellManager
        cm.FillManufacturer fr.targetKey, item
        ' ----------------------------------------------------------------
        UpdateSuggestions fr.targetKey
        
    End Sub
    
' ----------------------------------------------------------------
' Дата: 24.03.2023 07:55
' ----------------------------------------------------------------
Public Sub MainAction()
    
    Select Case True
    
        Case this.type = organisations
            MainActionOrganisation
        
        'todo: MainAction -- для других баз данных
        
    End Select
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 23.03.2023 22:35, 26.03.2023 13:08
    ' ----------------------------------------------------------------
    Private Sub MainActionOrganisation( _
        )
        Dim item As cItemOrganisation
        Set item = GetItem(fr.targetKey)
                
        Dim cm As New cCellManager
        cm.FillOrganisation fr.targetKey, item
        ' ----------------------------------------------------------------
        UpdateSuggestions fr.targetKey
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 26.03.2023 13:08
        ' ----------------------------------------------------------------
        Private Sub UpdateSuggestions( _
            newKey As String _
            )

            Dim result As New Collection
            Set result = MergedSearchCache(newKey)
            ' ----------------------------------------------------------------
            Dim newSearchKeys As String
            newSearchKeys = PreparedSuggestionKeys(result)
            ' ----------------------------------------------------------------
            FillNewSearchKeys newSearchKeys
            
        End Sub
            ' ----------------------------------------------------------------
            ' Дата: 26.03.2023 13:12
            ' ----------------------------------------------------------------
            Private Function MergedSearchCache( _
                newKey As String _
                ) As Collection
                
                If fr.defaultSuggestions Is Nothing Then _
                    Exit Function
                
                Dim result As New Collection
                result.Add newKey
            
                Dim i As Integer
                For i = 1 To fr.defaultSuggestions.count
                    
                    If result.count = RESULT_MAXCOUNT Then _
                        Exit For
                    ' ----------------------------------------------------------------
                    If fr.defaultSuggestions(i) <> newKey Then _
                        result.Add fr.defaultSuggestions(i)
                    
                Next i

                Set fr.defaultSuggestions = result
                Set MergedSearchCache = result
                
            End Function
            ' ----------------------------------------------------------------
            ' Дата: 26.03.2023 13:19
            ' ----------------------------------------------------------------
            Private Function PreparedSuggestionKeys( _
                source As Collection _
                ) As String
                
                Dim i As Integer, result As String
                For i = 1 To source.count
                    result = CheckAndDelimite(result) & source(i)
                Next i
                
                PreparedSuggestionKeys = result
                
            End Function
    
' ----------------------------------------------------------------
' Дата: 12.03.2023 10:00
' ----------------------------------------------------------------
Friend Sub Initialize( _
    typeDb As EDadaBaseTypes, _
    Optional checkIfCacheActual As Boolean = True _
    )
    
    this.type = typeDb
    ' ----------------------------------------------------------------
    If Config.sourceDataPath = Base.defaultValue Then _
        Exit Sub
    ' ----------------------------------------------------------------
    this.pathDb = PathFromType

    this.cacheFileName = CacheFileNameFromType
    this.cacheFilePath = fso.BuildPath(this.pathDb, this.cacheFileName)
    this.cacheItemCount = Base.defaultValue(intDefault)
    
    LoadCache checkIfCacheActual
    FillDefaultSuggestions
    FillKeysExistingItems

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.03.2023 15:16
    ' ----------------------------------------------------------------
    Private Sub FillDefaultSuggestions()
        
        Dim cacheString As String: cacheString = LastSearchFromType
        Dim cacheColl As New Collection: Set cacheColl = GetStoredCache(cacheString)
        
        Dim isRecache As Boolean: isRecache = IsNeedRecache(cacheString)
            ' ----------------------------------------------------------------
        If isRecache Then
        
            Dim newSearchKeys As String: newSearchKeys = PreparedSuggestionKeys(cacheColl)
            FillNewSearchKeys newSearchKeys
            
        End If
        
        FillSuggestions cacheColl
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:03
        ' ----------------------------------------------------------------
        Private Function LastSearchFromType( _
            ) As String
            
            Select Case True
                
                Case this.type = organisations
                    LastSearchFromType = Cache.organisationLastSearch
                    
                Case this.type = instruments
                    LastSearchFromType = Cache.instrumentLastSearch
                    
                Case this.type = persons
                    LastSearchFromType = Cache.personLastSearch
                
                Case this.type = etalons
                    LastSearchFromType = Cache.etalonLastSearch
                
                Case Else
                    LastSearchFromType = Base.defaultValue
                
            End Select
        
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:05
        ' ----------------------------------------------------------------
        Private Sub FillNewSearchKeys( _
            newSearchKeys As String _
            )
            
            Select Case True
                
                Case this.type = organisations
                    Cache.organisationLastSearch = newSearchKeys
                    
                Case this.type = instruments
                    Cache.instrumentLastSearch = newSearchKeys
                    
                Case this.type = persons
                    Cache.personLastSearch = newSearchKeys
                    
                Case this.type = etalons
                    Cache.etalonLastSearch = newSearchKeys
                
                Case Else
                    Exit Sub
                
            End Select
            Cache.Save
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 26.03.2023 12:52, 19.04.2023 20:00
        ' ----------------------------------------------------------------
        Private Function GetStoredCache( _
            cacheString As String _
            ) As Collection
            
            Set GetStoredCache = New Collection
            If cacheString = Base.defaultValue Then _
                Exit Function
            ' ----------------------------------------------------------------
            Dim tmpArr() As String
            tmpArr = Split(cacheString, CACHE_DELIMITER)
            
            Dim i As Integer, result As New Collection
            For i = LBound(tmpArr) To UBound(tmpArr)
            
                Dim tmp As IDatabaseItem
                Set tmp = GetItem(tmpArr(i)) 'проверка на существование элемента в БД
                
                If Not tmp Is Nothing Then _
                    result.Add tmpArr(i)
            Next
            
            Set GetStoredCache = result

        End Function
        ' ----------------------------------------------------------------
        ' Дата: 26.03.2023 12:52
        ' ----------------------------------------------------------------
        Private Function IsNeedRecache( _
            cacheString As String _
            ) As Boolean
            
            If cacheString = Base.defaultValue Then _
                Exit Function
            ' ----------------------------------------------------------------
            Dim tmpArr() As String
            tmpArr = Split(cacheString, CACHE_DELIMITER)
            
            Dim i As Integer
            For i = LBound(tmpArr) To UBound(tmpArr)
            
                Dim tmp As IDatabaseItem
                Set tmp = GetItem(tmpArr(i))
                
                If tmp Is Nothing Then _
                    IsNeedRecache = True: _
                    Exit Function
            Next
            
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 25.03.2023 15:22
        ' ----------------------------------------------------------------
        Private Sub FillSuggestions( _
            cacheColl As Collection _
            )
            
            If cacheColl.count = RESULT_MAXCOUNT Then _
                Set fr.defaultSuggestions = cacheColl: _
                Exit Sub
            ' ----------------------------------------------------------------
            Dim allc As New Collection: Set allc = AllCache
            Dim keyString As String: keyString = PreparedSuggestionKeys(cacheColl)
            
            Dim i As Integer
            For i = 1 To allc.count
            
                If cacheColl.count = RESULT_MAXCOUNT Then _
                    Exit For
                ' ----------------------------------------------------------------
                Dim keyFromAllCache As String: keyFromAllCache = KeyFromCache(allc(i))
                Dim isKeyExist As Boolean: isKeyExist = CBool(InStr(keyString, keyFromAllCache))
                ' ----------------------------------------------------------------
                If Not isKeyExist Then _
                    cacheColl.Add keyFromAllCache
                
            Next i
            ' ----------------------------------------------------------------
            Set fr.defaultSuggestions = cacheColl
            
        End Sub
            ' ----------------------------------------------------------------
            ' Дата: 15.03.2023 09:15
            ' ----------------------------------------------------------------
            Private Function AllCache( _
                ) As Collection
                
                Dim splitData() As String
                splitData = Split(this.cacheData, vbNewLine)
                
                Dim i As Integer, result As New Collection
                For i = LBound(splitData) To UBound(splitData)
                    result.Add splitData(i)
                Next
                
                Set AllCache = result
                
            End Function
    ' ----------------------------------------------------------------
    ' Дата: 11.03.2023 21:03
    ' ----------------------------------------------------------------
    Private Function PathFromType( _
        ) As String
        
        Select Case True
    
            Case this.type = organisations
                PathFromType = Config.organisationsPath
    
            Case this.type = instruments
                PathFromType = Config.instrumentsPath
    
            Case this.type = etalons
                PathFromType = Config.etalonsPath
    
            Case this.type = persons
                PathFromType = Config.personsPath
            
            Case this.type = archiveLocal
                PathFromType = Config.archiveLocalPath
    
        End Select
    
    End Function
    Private Function CacheFileNameFromType( _
        ) As String
        
        Dim ck As New cConfigKeysDatabase
        Select Case True
    
            Case this.type = organisations
                CacheFileNameFromType = ck.organisationsDir & "." & CACHE_FILEEXT
    
            Case this.type = instruments
                CacheFileNameFromType = ck.instrumentsDir & "." & CACHE_FILEEXT
    
            Case this.type = etalons
                CacheFileNameFromType = ck.etalonsDir & "." & CACHE_FILEEXT
    
            Case this.type = persons
                CacheFileNameFromType = ck.personsDir & "." & CACHE_FILEEXT
            
            Case this.type = archiveLocal
                CacheFileNameFromType = ck.archiveLocalDir & "." & CACHE_FILEEXT
    
        End Select
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 11.03.2023 20:04, 02.04.2023 09:47
    ' Назначение: загрузка в память кеша базы данных
    ' ----------------------------------------------------------------
    Private Sub LoadCache( _
        Optional isCheckActual As Boolean = True _
        )
        
        Dim isReacheNeeded As Boolean
        ' ----------------------------------------------------------------
        If Not fso.FileExists(this.cacheFilePath) Then _
            isReacheNeeded = True
        ' ----------------------------------------------------------------
        If isCheckActual Then
             
            If Not IsCacheActual Then _
                isReacheNeeded = True 'перехешировать раз в RECACHE_INTERVAL дней
        End If
        ' ----------------------------------------------------------------
        If isReacheNeeded Then _
            ReCacheData msgFirst:=msgboxMessage, msgSecond:=msgboxMessage
        ' ----------------------------------------------------------------
        this.cacheData = Base.ContentFromFile(this.cacheFilePath)
        this.cacheItemCount = GetCacheItemCount
        ' ----------------------------------------------------------------
        If Not IsCacheCountCorrect Then _
            ReCacheData msgFirst:=msgboxMessage, msgSecond:=msgboxMessage: _
            this.cacheItemCount = GetCacheItemCount
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 02.04.2023 09:48
        ' ----------------------------------------------------------------
        Private Function IsCacheCountCorrect( _
            ) As Boolean
            
            IsCacheCountCorrect = False
            
            Select Case this.type
            
                Case organisations, persons, etalons
                    IsCacheCountCorrect = IsDirCountCorrect
                
                Case instruments, archiveLocal
                    'todo: [-] IsCacheCountCorrect придумать, как связать количество каталогов
                    IsCacheCountCorrect = True
                
            End Select
            
        End Function
            ' ----------------------------------------------------------------
            ' Дата: 02.04.2023 09:49
            ' ----------------------------------------------------------------
            Private Function IsDirCountCorrect( _
                ) As Boolean
                
                Dim dirCount As Integer: dirCount = fso.GetFolder(this.pathDb).SubFolders.count
                IsDirCountCorrect = dirCount = this.cacheItemCount
                
            End Function
            Private Function IsInstrumentsCountCorrect( _
                ) As Boolean
    
                Dim dirCount As Integer: dirCount = fso.GetFolder(this.pathDb).SubFolders.count
                IsInstrumentsCountCorrect = dirCount = this.cacheItemCount + 1 ' для СИ дополнительный каталог templates base
    
            End Function
                
        ' ----------------------------------------------------------------
        ' Дата: 21.03.2023 14:03
        ' ----------------------------------------------------------------
        Private Function GetCacheItemCount( _
            ) As Integer
            
            Dim arrCache() As String
            arrCache = Split(this.cacheData, vbNewLine)
            
            Dim result As Integer
            result = UBound(arrCache) + 1
            
            GetCacheItemCount = result
            
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 13.03.2023 19:33
        ' ----------------------------------------------------------------
        Private Function IsCacheActual( _
            ) As Boolean
            
            IsCacheActual = False
            ' ----------------------------------------------------------------
            If Not fso.FileExists(this.cacheFilePath) Then _
                Exit Function
            ' ----------------------------------------------------------------
            Dim cacheDate As Date
            cacheDate = CDate(fso.GetFile(this.cacheFilePath).DateLastModified)

            Dim cacheDateDiff As Single
            cacheDateDiff = Now - cacheDate
            
            If cacheDateDiff > RECACHE_INTERVAL Then _
                Exit Function
            
            IsCacheActual = True
        
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 12.03.2023 10:52
        ' Назначение:
        ' ----------------------------------------------------------------
        Private Sub InitiateCaching( _
            )
                
            Dim db As Scripting.folder
            Set db = fso.GetFolder(this.pathDb)
            
            If Not CBool(db.SubFolders.count) Then _
                Exit Sub
            
            this.cacheData = vbNullString
            
            Dim pb As New cProgressBar
            pb.Initialize db.SubFolders.count

            Dim dbItem As Scripting.folder, result As String
            For Each dbItem In db.SubFolders
            
                DoEvents
                pb.ShowProgress
                
                Dim itemCache As String
                itemCache = CacheFromItem(dbItem.name)
                
                If itemCache <> Base.defaultValue Then
                    ' ----------------------------------------------------------------
                    If result <> vbNullString Then _
                        result = result & vbNewLine
    
                    result = result & itemCache
                    ' ----------------------------------------------------------------
                End If
            Next

            pb.Terminate
            
            Dim sortedResult As String
            sortedResult = SortedString(result)
            
            this.cacheData = sortedResult
            
        End Sub
            ' ----------------------------------------------------------------
            ' Дата: 05.04.2023 10:52
            ' ----------------------------------------------------------------
            Private Function SortedString( _
                source As String _
                ) As String

                Dim tmpArr() As String
                tmpArr = Split(source, vbNewLine)
                
                Dim newArr() As String
                newArr = Base.SortArray(tmpArr)
                
                Dim result As String
                result = Join(newArr, vbNewLine)
                
                SortedString = result
                
            End Function
            
            ' ----------------------------------------------------------------
            ' Дата: 12.03.2023 10:11
            ' Назначение:
            ' ----------------------------------------------------------------
            Private Function CacheFromItem( _
                key As String _
                ) As String
                
                CacheFromItem = Base.defaultValue
                
                ds.Initialize _
                    sourcePath:=this.pathDb, _
                    key:=key
                ' ----------------------------------------------------------------
                If ds.isHidden Then _
                    Exit Function
                ' ----------------------------------------------------------------
                Select Case True
                
                    Case this.type = organisations
                        CacheFromItem = OrgaisationCache(key)
                    
                    Case this.type = instruments
                        CacheFromItem = InstrumentCache(key)
                    
                    Case this.type = etalons
                        CacheFromItem = EtalonCache(key)
                    
                    Case this.type = persons
                        CacheFromItem = PersonsCache(key)
                    
                    Case this.type = archiveLocal
                        CacheFromItem = ArchiveLocalCache(key)
                        
                End Select
                
            End Function
                ' ----------------------------------------------------------------
                ' Дата: 19.03.2023 09:39
                ' ----------------------------------------------------------------
                Private Function ArchiveLocalCache( _
                    key As String _
                    ) As String

                    ArchiveLocalCache = Base.defaultValue

                    If CBool(InStr(key, "===")) Then _
                        Exit Function 'месячный разделитель

                    If CBool(InStr(key, "+++")) Then _
                        Exit Function 'годовой разделитель

                    If Not key Like "????.??*" Then _
                        Exit Function

                    ArchiveLocalCache = key

                End Function
                
                ' ----------------------------------------------------------------
                ' Дата: 12.03.2023 10:27
                ' Назначение: строка кеша файла
                ' ----------------------------------------------------------------
                Private Function OrgaisationCache( _
                    key As String _
                    ) As String
                    
                    OrgaisationCache = Base.defaultValue
                    
                    Dim it As cItemOrganisation
                    Set it = OrganisationFromItem
                    ' ----------------------------------------------------------------
                    'первый блок гарантированно будет в кеше
                    Dim result As String
                    If it.shortName <> Base.defaultValue Then _
                        result = it.shortName
                        
                    If it.legalAddress <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & it.legalAddress
                    
                    If it.taxNumber <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & it.taxNumber
                    ' ----------------------------------------------------------------
                    If it.country <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & it.country
                    
                    If it.fullName <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & it.fullName
                    ' ----------------------------------------------------------------
                    If result = vbNullString Then _
                        Exit Function
                    ' ----------------------------------------------------------------
                    result = result & Base.delimiterKeyValue & key
                    OrgaisationCache = result
                    
                End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 18.03.2023 12:28
                    ' ----------------------------------------------------------------
                    Private Function CheckAndDelimite( _
                        source As String _
                        ) As String
                        
                        CheckAndDelimite = source
                        
                        If source <> vbNullString Then _
                            CheckAndDelimite = source & CACHE_DELIMITER
                        
                    End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 15.03.2023 13:56
                    ' ----------------------------------------------------------------
                    Private Function OrganisationFromItem( _
                        ) As cItemOrganisation

                        Dim org As New cItemOrganisation, _
                            keys As New cItemKeysOrganisation
            
                        org.shortName = ds.FindValue(keys.shortName)
                        org.country = ds.FindValue(keys.country)
                        org.legalAddress = ds.FindValue(keys.legalAddress)
                        org.fullName = ds.FindValue(keys.fullName)
                        org.taxNumber = ds.FindValue(keys.taxNumber)
                        
                        Set OrganisationFromItem = org
                        
                    End Function
                ' ----------------------------------------------------------------
                ' Дата: 18.03.2023 15:49
                ' ----------------------------------------------------------------
                Private Function InstrumentCache( _
                    key As String _
                    ) As String
                    
                    InstrumentCache = Base.defaultValue
                    
                    Dim groupItem As New Collection
                    Set groupItem = ds.GetInstrumentGroup
                    
                    Dim i As Integer, result As String, mi As cItemInstrument
                    For i = 1 To groupItem.count
                    
                        Set mi = InstrumentFromItem(groupItem(i))
                        
                        If result <> vbNullString Then _
                            result = result & vbNewLine 'для второй и последующих строк

                        If mi.regFgisNum <> Base.defaultValue Then _
                            result = result & mi.regFgisNum
                        
                        If mi.typeDevice <> Base.defaultValue Then _
                            result = CheckAndDelimite(result) & mi.typeDevice
                        
                        If mi.methodic <> Base.defaultValue Then _
                            result = CheckAndDelimite(result) & mi.methodic
                        
                        If mi.manufacturer <> Base.defaultValue Then _
                            result = CheckAndDelimite(result) & mi.manufacturer
                        ' ----------------------------------------------------------------
                        If result <> vbNullString Then _
                            result = result & Base.delimiterKeyValue & key & SUBKEY_DELIMITER & mi.regFgisNum
                        ' ----------------------------------------------------------------
                    Next
                    
                    If CBool(groupItem.count) Then _
                        InstrumentCache = result 'чтобы не было пустых строк
                    
                End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 18.03.2023 15:52
                    ' ----------------------------------------------------------------
                    Private Function InstrumentFromItem( _
                        subKey As String _
                        ) As cItemInstrument
                        
                        If subKey = Base.defaultValue Then _
                            Exit Function
                        
                        Dim mInst As New cItemInstrument, _
                            keys As New cItemKeysInstrument
                        
                        mInst.regFgisNum = ds.FindValue( _
                            searchKey:=keys.regFgisNum, _
                            subKey:=subKey _
                            ) 'первый поиск в новом файле
                            
                        mInst.name = ds.FindValue(keys.name)
                        mInst.typeDevice = ds.FindValue(keys.typeDevice)
                        mInst.methodic = ds.FindValue(keys.methodic)
                        
                        mInst.manufacturer = ds.FindValue(keys.manufacturer)
                        mInst.dateExpire = ds.FindValue(keys.dateExpire)
                        
                        Set InstrumentFromItem = mInst
                        
                    End Function
                ' ----------------------------------------------------------------
                ' Дата: 18.03.2023 12:19
                ' ----------------------------------------------------------------
                Private Function EtalonCache( _
                    key As String _
                    ) As String
                    
                    EtalonCache = Base.defaultValue
                    
                    Dim et As cItemEtalon
                    Set et = EtalonFromItem
                    
                    Dim result As String
                    
                    If et.typeDevice <> Base.defaultValue Then _
                        result = et.typeDevice
                    
                    If et.factoryNumber <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & et.factoryNumber
                    
                    If et.regFgisNum <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & et.regFgisNum
                        
                    If et.numberDocument <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & et.numberDocument
                    
                    If et.dateExpire <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & ExpireStatus(et)
                    
                    If et.comment <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & et.comment
                    ' ----------------------------------------------------------------
                    If result = vbNullString Then _
                        Exit Function
                    ' ----------------------------------------------------------------
                    result = result & Base.delimiterKeyValue & key
                    EtalonCache = result
                    
                End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 18.03.2023 12:22
                    ' ----------------------------------------------------------------
                    Private Function EtalonFromItem( _
                        ) As cItemEtalon

                        Dim etal As New cItemEtalon, _
                            keys As New cItemKeysEtalon
                        
                        etal.typeDevice = ds.FindValue(keys.typeDevice)
                        etal.factoryNumber = ds.FindValue(keys.factoryNumber)
                        
                        etal.regFgisNum = ds.FindValue(keys.regFgisNum)
                        etal.numberDocument = ds.FindValue(keys.numberDocument)
                        etal.dateExpire = ds.FindValue(keys.dateExpire)
                        
                        etal.comment = ds.FindValue(keys.comment)
                        
                        Set EtalonFromItem = etal
                        
                    End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 18.03.2023 13:03
                    ' ----------------------------------------------------------------
                    Private Function ExpireStatus( _
                        etal As cItemEtalon _
                        ) As String
                        
                        If etal.status = illegal Then _
                            ExpireStatus = "истёк " & etal.dateExpire: _
                            Exit Function
                        
                        If CBool(InStr(etal.numberDocument, "С-В")) Then _
                            ExpireStatus = "поверен до " & etal.dateExpire: _
                            Exit Function
                        
                        ExpireStatus = "аттестован до " & etal.dateExpire
                        
                    End Function
                ' ----------------------------------------------------------------
                ' Дата: 13.03.2023 19:57
                ' ----------------------------------------------------------------
                Private Function PersonsCache( _
                    key As String _
                    ) As String
                    
                    PersonsCache = Base.defaultValue
                    
                    Dim pe As cItemPerson
                    Set pe = PersonFromItem
                    
                    Dim result As String
                    result = pe.lastName & " " & pe.firstName & " " & pe.middleName
                    
                    If pe.state <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & pe.state
                        
                    If pe.organisation <> Base.defaultValue Then _
                        result = CheckAndDelimite(result) & pe.organisation
                    
                    If pe.birthDay <> Base.defaultValue Then
                        
                        result = CheckAndDelimite(result) & pe.birthDay
                        result = CheckAndDelimite(result) & Format(CDate(pe.birthDay), "mmmm")
                        
                    End If
                    ' ----------------------------------------------------------------
                    result = result & Base.delimiterKeyValue & key
                    ' ----------------------------------------------------------------
                    PersonsCache = result
                    
                End Function
                    ' ----------------------------------------------------------------
                    ' Дата: 15.03.2023 13:56
                    ' ----------------------------------------------------------------
                    Private Function PersonFromItem( _
                        ) As cItemPerson
                        
                        Dim per As New cItemPerson, _
                            keys As New cItemKeysPerson
                        
                        per.lastName = ds.FindValue(keys.lastName)
                        per.firstName = ds.FindValue(keys.firstName)
                        per.middleName = ds.FindValue(keys.middleName)
                        
                        per.state = ds.FindValue(keys.state)
                        per.organisation = ds.FindValue(keys.organisation)
                        
                        Dim value As String: value = ds.FindValue(keys.phoneNumber)
                        If Not value = Base.defaultValue Then _
                            per.phoneNumber = value
                        
                        value = ds.FindValue(keys.email)
                        If Not value = Base.defaultValue Then _
                            per.email = value
                        
                        per.birthDay = ds.FindValue(keys.birthDay)
                        
                        Set PersonFromItem = per

                    End Function
' ----------------------------------------------------------------
' Дата: 26.03.2023 21:34
' ----------------------------------------------------------------
Private Sub FillKeysExistingItems()
    
    Dim splitData() As String
    splitData = Split(this.cacheData, vbNewLine)
    
    Dim i As Integer
    For i = LBound(splitData) To UBound(splitData)
        splitData(i) = KeyFromCache(splitData(i))
    Next
    
    fr.keysExistingItems = Join(splitData, CACHE_DELIMITER)
    
End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 21:38, 28.03.2023 21:01
' ----------------------------------------------------------------
Public Function IsKeyExists( _
    key As String _
    )
        
    Dim keyArr(vbEmpty) As String
    keyArr(vbEmpty) = key
    
    Dim values As New Collection
    Set values = FindInCache(keyArr, returnKeys:=True) 'результат случайного поиска по индивидуальному ключу по всему кешу

    Dim i As Integer, result As Boolean
    For i = 1 To values.count
        
        If key = values(i) Then _
            result = True: _
            Exit For
        
    Next i
    
    IsKeyExists = result

End Function
' ----------------------------------------------------------------
' Дата: 15.03.2023 13:53
' ----------------------------------------------------------------
'работает, но принимающий объект должен иметь конечный тип
Public Function GetItem( _
    key As String _
    ) As IDatabaseItem
    
    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=key
    
    If Not fso.FolderExists(ds.itemFolderPath) Then
        
        Handler.Notify _
            mes:="Не удалось найти элемент по ключу «" & key & "»." & vbNewLine & "Выполнение прервано.", _
            msgType:=debugMessage
            
        Exit Function
        
    End If
    
    Select Case True

        Case this.type = organisations
            Set GetItem = OrganisationFromItem
            
        Case this.type = instruments
            Dim subKey As String: subKey = InstrumentSubKey(key)
            Set GetItem = InstrumentFromItem(subKey)

        Case this.type = etalons
            Set GetItem = EtalonFromItem

        Case this.type = persons
            Set GetItem = PersonFromItem
        
        Case this.type = archiveLocal
            Set GetItem = ArchiveDataFromItem
            
        Case Else
            Handler.Notify "Не работает для этой конфигурации"

    End Select
    
End Function
    ' ----------------------------------------------------------------
    ' Дата: 19.03.2023 10:10
    ' ----------------------------------------------------------------
    Private Function ArchiveDataFromItem( _
        ) As cItemArchive

        Dim arc As New cItemArchive, _
            keys As New cItemKeysArchive
            
        arc.beginDate = ds.FindValue(keys.beginDate)
        arc.endDate = ds.FindValue(keys.endDate)
                    
        arc.typeSerial = ds.FindValue(keys.typeSerial)
        arc.organisation = ds.FindValue(keys.organisation)
        arc.comment = ds.FindValue(keys.comment)

        Set ArchiveDataFromItem = arc
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 18.03.2023 18:20
    ' ----------------------------------------------------------------
    Private Function InstrumentSubKey( _
        key As String _
        ) As String
        
        InstrumentSubKey = Base.defaultValue
        
        Dim instrKey As Integer
        instrKey = InStr(key, SUBKEY_DELIMITER)
        
        If CBool(instrKey) Then _
            InstrumentSubKey = Mid(key, instrKey + 1)
        
    End Function
    
' ----------------------------------------------------------------
' Дата: 12.03.2023 11:27
' Назначение: принудительное получение файла кеширования базы данных
' ----------------------------------------------------------------
Public Sub ReCacheData( _
    Optional msgFirst As ENotifyTypes = nomessage, _
    Optional msgSecond As ENotifyTypes = msgboxMessage _
    )

    InitiateCaching
    
    Base.WriteContent _
        path:=this.cacheFilePath, _
        content:=this.cacheData
    
End Sub
' ----------------------------------------------------------------
' Дата: 24.03.2023 10:31
' ----------------------------------------------------------------
Public Sub OpenItemDir( _
    )
    
    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=fr.targetKey

    ds.OpenItemDirectory
    
    UpdateSuggestions fr.targetKey
    
End Sub
' ----------------------------------------------------------------
' Дата: 24.03.2023 10:28
' ----------------------------------------------------------------
Public Sub OpenDatabaseDir( _
    )
    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=vbNullString
        
    ds.OpenItemDirectory
    
End Sub

Public Property Get IsDirAvailable( _
    key As String _
    ) As Boolean

    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=key
        
    Dim result As String
    result = fso.FolderExists(ds.itemFolderPath)
    
    IsDirAvailable = result
    
End Property
' ----------------------------------------------------------------
' Дата: 12.03.2023 09:07
' Назначение: добавить элемент в БД
' ----------------------------------------------------------------
Friend Function AddItem( _
    key As String, _
    Optional itemData As IDatabaseItem _
    ) As String
    
    Dim checkedKey As String
    checkedKey = CheckedName(key)
    
    AddItem = checkedKey
    
    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=checkedKey
    ' ----------------------------------------------------------------
    ds.CreateItemDir
    ' ----------------------------------------------------------------
    
    Select Case True
    
        Case this.type = organisations
            AddOrganisation itemData
        
        Case this.type = instruments
            AddInstrument itemData
            Exit Function 'это не ошибка, процедура записи в файл отличается от остальных
        
        Case this.type = etalons
            AddEtalon itemData
        
        Case this.type = persons
            AddPerson itemData
        
        Case this.type = archiveLocal
            AddToLocalArchive itemData
        
    End Select
    
    ds.WriteToFile
    UpdateSuggestions checkedKey
    
End Function
    ' ----------------------------------------------------------------
    ' Дата: 19.03.2023 10:08
    ' ----------------------------------------------------------------
    Private Sub AddToLocalArchive( _
        itemData As IDatabaseItem _
        )
        
        Dim item As New cItemArchive
        If Not itemData Is Nothing Then _
            Set item = itemData
        
        Dim keys As New cItemKeysArchive
        ds.PrepareToWrite keys.beginDate, item.beginDate
        ds.PrepareToWrite keys.endDate, item.endDate
        
        ds.PrepareToWrite keys.typeSerial, item.typeSerial
        ds.PrepareToWrite keys.organisation, item.organisation
        ds.PrepareToWrite keys.comment, item.comment
        
    End Sub
    Private Sub AddOrganisation( _
        itemData As IDatabaseItem _
        )
        
        Dim item As New cItemOrganisation
        If Not itemData Is Nothing Then _
            Set item = itemData
        
        Dim keys As New cItemKeysOrganisation
        ds.PrepareToWrite keys.shortName, item.shortName
        ds.PrepareToWrite keys.fullName, item.fullName
        
        ds.PrepareToWrite keys.country, item.country
        ds.PrepareToWrite keys.legalAddress, item.legalAddress
        ds.PrepareToWrite keys.taxNumber, item.taxNumber
    
    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 18.03.2023 14:34
    ' ----------------------------------------------------------------
    Private Sub AddInstrument( _
        itemData As IDatabaseItem _
        )
        
        Dim mi As New cItemInstrument
        If Not itemData Is Nothing Then _
            Set mi = itemData

        Dim keys As New cItemKeysInstrument
        ds.PrepareToWrite keys.regFgisNum, mi.regFgisNum
        ds.PrepareToWrite keys.name, mi.name

        ds.PrepareToWrite keys.typeDevice, mi.typeDevice
        ds.PrepareToWrite keys.methodic, mi.methodic
        ds.PrepareToWrite keys.dateExpire, mi.dateExpire
        ds.PrepareToWrite keys.manufacturer, mi.manufacturer
        ' ----------------------------------------------------------------
        ds.WriteToFile key:=mi.regFgisNum 'это не ошибка, процедура записи в файл отличается от остальных БД

    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 18.03.2023 12:09
    ' ----------------------------------------------------------------
    Private Sub AddEtalon( _
        itemData As IDatabaseItem _
        )
        
        Dim item As New cItemEtalon
        If Not itemData Is Nothing Then _
            Set item = itemData
        
        Dim keys As New cItemKeysEtalon
        ds.PrepareToWrite keys.typeDevice, item.typeDevice
        ds.PrepareToWrite keys.factoryNumber, item.factoryNumber
        
        ds.PrepareToWrite keys.regFgisNum, item.regFgisNum
        ds.PrepareToWrite keys.numberDocument, item.numberDocument
        ds.PrepareToWrite keys.dateExpire, item.dateExpire
        
        ds.PrepareToWrite keys.comment, item.comment
        
    End Sub
    Private Sub AddPerson( _
        itemData As IDatabaseItem _
        )
        
        Dim person As New cItemPerson
        If Not itemData Is Nothing Then _
            Set person = itemData
        
        Dim keys As New cItemKeysPerson
        ds.PrepareToWrite keys.lastName, person.lastName
        ds.PrepareToWrite keys.firstName, person.firstName
        ds.PrepareToWrite keys.middleName, person.middleName
        
        ds.PrepareToWrite keys.state, person.state
        ds.PrepareToWrite keys.organisation, person.organisation
        
        ds.PrepareToWrite keys.email, person.email
        ds.PrepareToWrite keys.phoneNumber, person.phoneNumber
        
        ds.PrepareToWrite keys.birthDay, person.birthDay
    
    End Sub
' ----------------------------------------------------------------
' Дата: 06.04.2023 22:02
' ----------------------------------------------------------------
Public Function UpdateKey( _
    sourceKey As String, _
    newKey As String _
    ) As Boolean
    
    UpdateKey = False
    ' ----------------------------------------------------------------
    ds.Initialize this.pathDb, sourceKey
    
    Dim sourceItemPath As String: sourceItemPath = ds.itemFilePath
    Dim sourceFolderPath As String: sourceFolderPath = ds.itemFolderPath
    Dim sourceItemCache As String: sourceItemCache = CacheFromItem(sourceKey)
    ' ----------------------------------------------------------------
    ds.Initialize this.pathDb, newKey
    
    Dim newItemPath As String: newItemPath = fso.BuildPath(sourceFolderPath, fso.GetFileName(ds.itemFilePath))
    Dim newFolderPath As String: newFolderPath = ds.itemFolderPath
    
    Dim newItemCache As String: newItemCache = Replace(sourceItemCache, _
        Base.delimiterKeyValue & sourceKey, _
        Base.delimiterKeyValue & newKey _
        )
    ' ----------------------------------------------------------------
    fso.MoveFile sourceItemPath, newItemPath
    fso.MoveFolder sourceFolderPath, newFolderPath
    ' ----------------------------------------------------------------
    
    this.cacheData = Replace(this.cacheData, sourceItemCache, newItemCache)
    
    Base.WriteContent _
        path:=this.cacheFilePath, _
        content:=this.cacheData
        
    FillDefaultSuggestions
    FillKeysExistingItems
    
    UpdateSuggestions newKey

    UpdateKey = True
    
End Function
' ----------------------------------------------------------------
' Дата: 12.03.2023 12:08, 08.04.2023 13:18
' ----------------------------------------------------------------
Public Function Merge( _
    resultKey As String, _
    wipeKey As String, _
    Optional silentMode As Boolean _
    ) As Boolean

    Merge = False
    
    If Not IsUserApprove(wipeKey, resultKey, silentMode) Then _
        Exit Function
    ' ----------------------------------------------------------------
    ds.Initialize this.pathDb, wipeKey
    Dim wipeFolderPath As String: wipeFolderPath = ds.itemFolderPath
    Dim wipeItemCache As String: wipeItemCache = CacheFromItem(wipeKey)
    ' ----------------------------------------------------------------
    ds.Initialize this.pathDb, resultKey
    Dim resultFolderPath As String: resultFolderPath = ds.itemFolderPath
    ' ----------------------------------------------------------------
    MergeSubFiles wipeFolderPath, resultFolderPath
    MergeSubFolders wipeFolderPath, resultFolderPath
    ' ----------------------------------------------------------------
    Base.DeletePath wipeFolderPath
    If this.cacheData <> Base.defaultValue And wipeItemCache <> Base.defaultValue Then _
        this.cacheData = Replace(this.cacheData, wipeItemCache, vbNullString)
    
    Base.WriteContent _
        path:=this.cacheFilePath, _
        content:=this.cacheData
        
    If Not silentMode Then _
        ds.OpenItemDirectory
    
    If silentMode Then _
        Merge = True: _
        Exit Function
    
    FillDefaultSuggestions
    FillKeysExistingItems

    UpdateSuggestions resultKey
        
End Function
    ' ----------------------------------------------------------------
    ' Дата: 08.04.2023 13:23
    ' ----------------------------------------------------------------
    Private Function IsUserApprove( _
        wipeKey As String, _
        resultKey As String, _
        silentMode As Boolean _
        ) As Boolean
        
        IsUserApprove = False
        
        If silentMode Then _
            IsUserApprove = True: _
            Exit Function
        
        Dim msg As String
        msg = "Вы собираетесь выполнить слияние данных двух элементов. " & vbNewLine & vbNewLine & _
            "Все файлы и подкаталоги «" & wipeKey & "» будут перенесены в каталог «" & resultKey & "». " & _
            "Доступ к данным «" & wipeKey & "» будет осуществляться через «" & resultKey & "»." & vbNewLine & vbNewLine & _
            "Продолжить?"
        
        Dim ans As Integer: ans = Handler.ask(msg)
        IsUserApprove = ans = vbYes
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 12.03.2023 13:16, 18.03.2023 19:24
    ' ----------------------------------------------------------------
    Private Sub MergeSubFiles( _
        sourcePath As String, _
        destinationPath As String _
        )
        
        Dim db As Scripting.folder
        Set db = fso.GetFolder(sourcePath)
        
        If Not CBool(db.Files.count) Then _
            Exit Sub
        
        Dim subFile As Scripting.file
        For Each subFile In db.Files
        
            Dim destinationFilePath As String
            destinationFilePath = CheckedFilePath( _
                sourcePath:=subFile.path, _
                destinationPath:=fso.BuildPath(destinationPath, fso.GetFile(subFile).name))
            
            fso.CopyFile _
                subFile, _
                destinationFilePath, False

        Next
    
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 12.03.2023 12:53, 18.03.2023 19:24, 19.03.2023 08:44
        ' ----------------------------------------------------------------
        Private Function CheckedFilePath( _
            sourcePath As String, _
            destinationPath As String _
            ) As String
            
            Dim fileName As String
            fileName = fso.GetFileName(sourcePath)
            
            Dim parentFolderPath As String
            parentFolderPath = fso.GetParentFolderName(destinationPath)
            
            Dim result As String
            result = fso.BuildPath(parentFolderPath, fileName)
            ' ----------------------------------------------------------------
            If IsBodyFile(sourcePath) Then
                
                Dim key As String
                key = KeyFromPath(parentFolderPath)
                
             '   Debug.Assert Not CBool(InStr(key, "ГО.2.86"))
                
                If Not key = Base.defaultValue Then _
                    ds.Initialize this.pathDb, key: _
                    result = fso.BuildPath(ds.bodyTemplatesPath, fileName)
                
            End If
            ' ----------------------------------------------------------------
            Dim i As Integer: i = 1
            Dim fileMask As String: fileMask = fso.GetBaseName(result)
            Dim fileExt As String: fileExt = fso.GetExtensionName(result)
            
            Do
            
                If Not fso.FileExists(result) Then _
                    Exit Do
                
                result = fso.BuildPath(fso.GetParentFolderName(result), fileMask & " (" & i & ")." & fileExt)
                i = i + 1
                
            Loop
            
            CheckedFilePath = result
            
        End Function
            ' ----------------------------------------------------------------
            ' Дата: 19.03.2023 08:36
            ' ----------------------------------------------------------------
            Private Function IsBodyFile( _
                sourcePath As String _
                ) As Boolean
                
                Dim fileName As String
                fileName = fso.GetFileName(sourcePath)
                
                IsBodyFile = CBool(InStr(fileName, BODY_KEY))
                
            End Function
            ' ----------------------------------------------------------------
            ' Дата: 19.03.2023 08:59
            ' ----------------------------------------------------------------
            Private Function KeyFromPath( _
                path As String _
                ) As String
                
                KeyFromPath = Base.defaultValue
                
                Dim instrSeparator As Integer
                instrSeparator = InStr(StrReverse(path), Application.PathSeparator)
                
                KeyFromPath = Right(path, instrSeparator - 1)
                
            End Function
            
    ' ----------------------------------------------------------------
    ' Дата: 12.03.2023 13:20
    ' Назначение:
    ' ----------------------------------------------------------------
    Private Sub MergeSubFolders( _
        sourcePath As String, _
        destinationPath As String _
        )
        
        Dim db As Scripting.folder
        Set db = fso.GetFolder(sourcePath)
        
        If Not CBool(db.SubFolders.count) Then _
            Exit Sub
        
        Dim subFolder As Scripting.folder
        For Each subFolder In db.SubFolders
            
            Dim destinationFolderPath As String
            destinationFolderPath = CheckedFolderPath( _
                sourcePath:=subFolder.path, _
                destinationPath:=fso.BuildPath(destinationPath, subFolder.name))
            
            fso.CopyFolder _
                subFolder, _
                destinationFolderPath, False

        Next

    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 12.03.2023 12:53
        ' Назначение:
        ' ----------------------------------------------------------------
        ' ----------------------------------------------------------------
        ' Дата: 18.03.2023 19:24
        ' ----------------------------------------------------------------
        Private Function CheckedFolderPath( _
            sourcePath As String, _
            destinationPath As String _
            ) As String
            
            Dim result As String
            result = destinationPath
            
            Dim i As Integer: i = 1
            Dim folderMask As String: folderMask = fso.GetBaseName(result)
            
            Do
            
                If Not fso.FolderExists(result) Then _
                    Exit Do
                
                result = fso.BuildPath(fso.GetParentFolderName(destinationPath), folderMask & " (" & i & ")")
                i = i + 1
                
            Loop
            
            CheckedFolderPath = result
            
        End Function
'' ----------------------------------------------------------------
'' Дата: 15.03.2023 12:53
'' ----------------------------------------------------------------
'Friend Function Hide( _
'    key As String _
'    ) As Boolean
'
'    Hide = False
'
'    ds.Initialize _
'        sourcePath:=this.pathDb, _
'        key:=key
'
'    ds.HideFromCaching
'    ReCacheData msgSecond:=msgboxMessage
'
'    Hide = True
'
'End Function
' ----------------------------------------------------------------
' Дата: 26.03.2023 19:31
' ----------------------------------------------------------------
Public Sub Edit( _
    key As String _
    )
    
    ds.Initialize _
        sourcePath:=this.pathDb, _
        key:=key
    
    ds.OpenForEdit
    
    UpdateSuggestions key
    
End Sub
' ----------------------------------------------------------------
' Дата: 24.03.2023 09:05
' ----------------------------------------------------------------
Public Property Get ResultFilter( _
    index As Integer _
    ) As cFilterResults
    
    Set ResultFilter = Nothing
    
    If index <= fr.count Then _
        Set ResultFilter = fr.data(index - 1)

End Property
' ----------------------------------------------------------------
' Дата: 24.03.2023 21:47
' ----------------------------------------------------------------
Public Property Get ResultFilterCount() As Integer
    ResultFilterCount = fr.count
End Property
' ----------------------------------------------------------------
' Дата: 24.03.2023 21:48
' ----------------------------------------------------------------
Public Property Get TargetItemKey() As String
    TargetItemKey = fr.targetKey
End Property
Public Property Let TargetItemKey(ByVal NewValue As String)
    fr.targetKey = NewValue
End Property
' ----------------------------------------------------------------
' Дата: 25.03.2023 15:03, 16.04.2023 17:12
' ----------------------------------------------------------------
Public Function VniimPersons( _
    Optional filterKey As String _
    ) As Collection
    
    If this.type <> persons Then _
        Exit Function
        
    Dim FilterResult As Collection
    Set FilterResult = FilterByString(KEY_VNIIM)
    
    Dim i As Integer, result As New Collection
    For i = 1 To FilterResult.count
        
        Dim itemCorrect As Boolean: itemCorrect = True 'по умолчанию добавлять все элементы
        If filterKey <> vbNullString Then _
            itemCorrect = InStr(FilterResult(i), filterKey) > vbEmpty
        
        If itemCorrect Then
            
            Dim stringData As String: stringData = Left(FilterResult(i), InStr(FilterResult(i), CACHE_DELIMITER) - 1)
            result.Add stringData
            
        End If
        
    Next
    
    Set VniimPersons = result
    
End Function
' ----------------------------------------------------------------
' Дата: 27.03.2023 11:08
' ----------------------------------------------------------------
'todo: [-] FilterArchive Отказаться от этой функции в пользу Filter Cache
Public Function FilterArchive( _
    Optional searchKey As String _
    ) As Collection
    
    If Not this.type = archiveLocal Then _
        Exit Function
    
    If Not CBool(Len(searchKey)) Then _
        Set FilterArchive = AllCache: _
        Exit Function 'при вводе пустой строки выведется вся БД
    ' ----------------------------------------------------------------
    Dim result As Collection
    Set result = FilterByString(searchKey)
    ' ----------------------------------------------------------------
    Set FilterArchive = result 'полный перечень результатов поиска
    
End Function
' ----------------------------------------------------------------
' Дата: 16.03.2023 10:30
' ----------------------------------------------------------------
Friend Function FilterCache( _
    Optional searchKey As String _
    ) As Collection
    
    If Not CBool(Len(searchKey)) Then _
        Set FilterCache = fr.defaultSuggestions: _
        FillResultsFilter FilterCache: _
        Exit Function 'при вводе пустой строки выведется вся БД
    ' ----------------------------------------------------------------
    Dim result As Collection
    Set result = FilterByString(searchKey)
    ' ----------------------------------------------------------------
    Set FilterCache = result 'полный перечень результатов поиска
    
    FillResultsFilter result
    
End Function
    ' ----------------------------------------------------------------
    ' Дата: 24.03.2023 20:40
    ' Назначение: получить коллекцию данных для передачи на форму
    ' ----------------------------------------------------------------
    Private Sub FillResultsFilter( _
        fullSearchData As Collection _
        )
        
        ReDim fr.data(vbEmpty)
        fr.count = fullSearchData.count
        ' ----------------------------------------------------------------
        If Not CBool(fr.count) Then _
            Exit Sub
        ' ----------------------------------------------------------------
        Dim i As Integer: i = IndexMaxCount(fullSearchData)
                
        ReDim fr.data(i - 1)
        For i = LBound(fr.data) To UBound(fr.data)
            ' ----------------------------------------------------------------
            fr.data(i).key = KeyFromCache(fullSearchData(i + 1))
            ' ----------------------------------------------------------------
            Select Case True
            
                Case this.type = organisations
                    FillFilterOrganisations i
                
                Case this.type = instruments
                    FillFilterInstruments i
                
                Case this.type = etalons
                    FillFilterEtalons i
                
                Case this.type = persons
                    FillFilterPersons i

            End Select
        Next
    
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:17
        ' ----------------------------------------------------------------
        Private Sub FillFilterOrganisations( _
            index As Integer _
            )
            
            Dim item As New cItemOrganisation: Set item = GetItem(fr.data(index).key)
            If Not item Is Nothing Then
            
                fr.data(index).LabelLargeCaption = " " & item.shortName & ", " & item.country
                fr.data(index).LabelSmallCaption = " " & item.legalAddress
                fr.data(index).TipText = item.fullName
                
            End If
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:20
        ' ----------------------------------------------------------------
        Private Sub FillFilterInstruments( _
            index As Integer _
            )
            'todo: FillFilterInstruments -- проработать
            Dim item As New cItemInstrument: Set item = GetItem(fr.data(index).key)
            If Not item Is Nothing Then
            
                fr.data(index).LabelLargeCaption = " " & 1 'item.lastName & " " & item.firstName & " " & item.middleName & ", " & item.organisation
                fr.data(index).LabelSmallCaption = " " & 2 ' item.state & ", " & item.email & ", " & item.phoneNumber
                fr.data(index).TipText = 3 ' "День рождения: " & item.birthDay
                
            End If
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:20
        ' ----------------------------------------------------------------
        Private Sub FillFilterEtalons( _
            index As Integer _
            )
            'todo: FillFilterEtalons -- проработать
            Dim item As New cItemEtalon: Set item = GetItem(fr.data(index).key)
            If Not item Is Nothing Then
            
                fr.data(index).LabelLargeCaption = " " & 4 'item.lastName & " " & item.firstName & " " & item.middleName & ", " & item.organisation
                fr.data(index).LabelSmallCaption = " " & 5 ' item.state & ", " & item.email & ", " & item.phoneNumber
                fr.data(index).TipText = 6 ' "День рождения: " & item.birthDay
                
            End If
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.04.2023 20:20
        ' ----------------------------------------------------------------
        Private Sub FillFilterPersons( _
            index As Integer _
            )
            
            Dim item As New cItemPerson: Set item = GetItem(fr.data(index).key)
            If Not item Is Nothing Then
            
                fr.data(index).LabelLargeCaption = " " & item.lastName & " " & item.firstName & " " & item.middleName & ", " & item.organisation
                fr.data(index).LabelSmallCaption = " " & item.state & ", " & item.email & ", " & item.phoneNumber
                fr.data(index).TipText = "День рождения: " & Replace(item.birthDay, Base.defaultValue, Base.defaultValuePrinted)
                
            End If
            
        End Sub        ' ----------------------------------------------------------------
        ' Дата: 24.03.2023 21:21
        ' ----------------------------------------------------------------
        Private Function IndexMaxCount( _
            fullSearchData As Collection _
            ) As Integer
            
            Dim result As Integer: result = RESULT_MAXCOUNT
            If fullSearchData.count < RESULT_MAXCOUNT Then _
                result = fullSearchData.count
            
            IndexMaxCount = result
            
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 19.03.2023 11:21
        ' ----------------------------------------------------------------
        Private Function KeyFromCache( _
            cachedString As String _
            ) As String

            Dim result As String
            result = Mid(cachedString, InStr(cachedString, Base.delimiterKeyValue) + 1)
            
            KeyFromCache = result
            
        End Function
    ' ----------------------------------------------------------------
    ' Дата: 16.03.2023 16:56
    ' ----------------------------------------------------------------
    Private Function FilterByString( _
        Optional key As String _
        ) As Collection
        'ключ — непустое значение
        Set FilterByString = New Collection
        ' ----------------------------------------------------------------
        Dim keys() As String  'набор ключей для поиска в кеше
        keys = KeysForFilter(key)
        ' ----------------------------------------------------------------
        Dim values As New Collection
        Set values = FindInCache(keys) 'результат случайного поиска по индивидуальному ключу по всему кешу
        ' ----------------------------------------------------------------
        If Not CBool(values.count) Then _
            Exit Function 'совсем ничего не найдено
        ' ----------------------------------------------------------------
        Dim result As New Collection
        Set result = FilterByCoincedence(keys, values) 'фильтровать по совпадению ключей в строке
        ' ----------------------------------------------------------------
        Set FilterByString = result
        
    End Function
        ' ----------------------------------------------------------------
        ' Дата: 16.03.2023 11:43, 21.03.2023 21:05
        ' ----------------------------------------------------------------
        Private Function KeysForFilter( _
            key As String _
            ) As String()
            
            Dim mixedKey As String
            mixedKey = MixedLayoutKey(key) 'поисковый запрос с учётом русской и английской раскладки
    
            Dim keyArr() As String
            keyArr = Split(mixedKey, " ")
            
            KeysForFilter = keyArr
            
        End Function
            ' ----------------------------------------------------------------
            ' Дата: 16.03.2023 10:45
            ' ----------------------------------------------------------------
            Private Function MixedLayoutKey( _
                key As String _
                ) As String
                
                Dim result As String
                result = key
            
                If ct.isAvaiable Then _
                    result = Trim(key) & " " & ct.Reverse(Trim(key))
                
                MixedLayoutKey = result
                
            End Function
        ' ----------------------------------------------------------------
        ' Дата: 16.03.2023 12:10, 21.03.2023 21:05, 28.03.2023 20:56
        ' ----------------------------------------------------------------
        Private Function FindInCache( _
            keys() As String, _
            Optional returnKeys As Boolean = False _
            ) As Collection
            
            Dim i As Integer, findResult As String
            For i = LBound(keys) To UBound(keys)
                
                Dim keyResult As String 'результат посика всех совпадений по ключу
                keyResult = FindAllByKey(keys(i), findResult)
                
                If returnKeys Then _
                    keyResult = GetKeysFromCache(keyResult)
                
                If keyResult <> Base.defaultValue Then
                    
                    If findResult <> vbNullString Then _
                        findResult = findResult & vbNewLine
                    
                    findResult = findResult & keyResult
                    
                End If
                   
            Next
               
            Dim result As Collection
            Set result = CollFromString(findResult)
                
            Set FindInCache = result
                 
        End Function
            ' ----------------------------------------------------------------
            ' Дата: 08.04.2023 14:16
            ' ----------------------------------------------------------------
            Private Function GetKeysFromCache( _
                keyAllResults As String _
                ) As String
                
                Dim arrCache() As String
                arrCache = Split(keyAllResults, vbNewLine)
                
                Dim i As Integer
                For i = LBound(arrCache) To UBound(arrCache)
                    arrCache(i) = KeyFromCache(arrCache(i))
                Next i
                
                Dim result As String
                result = Join(arrCache, vbNewLine)
                
                GetKeysFromCache = result
                
            End Function
            ' ----------------------------------------------------------------
            ' Дата: 16.03.2023 17:10
            ' ----------------------------------------------------------------
            Private Function FindAllByKey( _
                key As String, _
                Optional exclude As String _
                ) As String
                
                FindAllByKey = Base.defaultValue
                
                Dim instrKey As Long
                instrKey = InStr(this.cacheData, key)
                
                Dim result As String
                Do While CBool(instrKey)
                    
                    Dim cacheString As String
                    cacheString = StringFromCache(instrKey)
                    
                    Dim isExist As Boolean
                    isExist = InStr(exclude, cacheString) > 0
                    ' ----------------------------------------------------------------
                    If Not isExist Then
                        
                        If result <> vbNullString Then _
                            result = result & vbNewLine
                        
                        result = result & cacheString
                        
                    End If
                    
                    ' ----------------------------------------------------------------
                    Dim startPosition As Long
                    startPosition = InStr(this.cacheData, cacheString) + Len(cacheString)
                    
                    If startPosition >= Len(this.cacheData) Then _
                        Exit Do
                    ' ----------------------------------------------------------------
                    instrKey = InStr(startPosition, this.cacheData, key)
                
                Loop
                
                If result <> vbNullString Then _
                    FindAllByKey = result
        
            End Function
                 ' ----------------------------------------------------------------
                ' Дата: 15.03.2023 17:25
                ' ----------------------------------------------------------------
                Private Function StringFromCache( _
                    instrKey As Long _
                    ) As String
                    
                    Dim leftEdge As Long
                    leftEdge = InStrRev( _
                                    start:=instrKey, _
                                    StringCheck:=this.cacheData, _
                                    StringMatch:=vbNewLine _
                                    )
                    ' ----------------------------------------------------------------
                    If Not CBool(leftEdge) Then
                        leftEdge = 1 'поиск в начале кеша
                    Else
                        leftEdge = leftEdge + 2 'иначе будут пустые строки
                    End If
                    ' ----------------------------------------------------------------
                    Dim rightEdge As Long
                    rightEdge = InStr(instrKey, this.cacheData, vbNewLine)
                    
                    If Not CBool(rightEdge) Then _
                        rightEdge = Len(this.cacheData) 'поиск в конце кеша
                
                    StringFromCache = Mid(this.cacheData, leftEdge, rightEdge - leftEdge)
                    
                End Function
        ' ----------------------------------------------------------------
        ' Дата: 16.03.2023 17:47
        ' ----------------------------------------------------------------
        Private Function CollFromString( _
            values As String _
            ) As Collection
            
            Dim arr() As String
            arr = Split(values, vbNewLine)
            
            Dim i As Integer, result As New Collection
            For i = LBound(arr) To UBound(arr)
            
                result.Add arr(i)
                
            Next
            
            Set CollFromString = result
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 16.03.2023 16:59
        ' ----------------------------------------------------------------
        Private Function FilterByCoincedence( _
            keys() As String, _
            values As Collection _
            ) As Collection
            
            Dim i As Integer, result As New Collection
            For i = 1 To values.count
                ' ----------------------------------------------------------------
                Dim isCoincedence As Boolean
                isCoincedence = CheckCoincedence(keys, values(i))
                ' ----------------------------------------------------------------
                If isCoincedence Then _
                    result.Add values(i)
                ' ----------------------------------------------------------------
            Next
            
            Set FilterByCoincedence = result
                        
        End Function
            ' ----------------------------------------------------------------
            ' Дата: 16.03.2023 16:50
            ' ----------------------------------------------------------------
            Private Function CheckCoincedence( _
                keys() As String, _
                value As String _
                ) As Boolean
                
                Dim i As Integer, result As Integer
                For i = LBound(keys) To UBound(keys)
                    
                    If InStr(value, keys(i)) Then _
                        result = result + 1
                
                Next
                
                Dim tcCount As Integer
                tcCount = TargetCoincedenceCount(keys)
                
                CheckCoincedence = result = tcCount
                 
            End Function
                ' ----------------------------------------------------------------
                ' Дата: 16.03.2023 10:49, 21.03.2023 21:08
                ' ----------------------------------------------------------------
                Private Function TargetCoincedenceCount( _
                    keys() As String _
                    ) As Integer
                    
                    Dim result As Integer
                    result = UBound(keys) + 1
                
                    If ct.isAvaiable Then _
                        result = result / 2
                    
                    TargetCoincedenceCount = result
                    
                End Function
' ----------------------------------------------------------------
' Дата: 24.02.2023 23:29
' ----------------------------------------------------------------
Friend Function Refactor( _
    oldDataBaseArray() As String _
    ) As Boolean
    
    Refactor = False

    Dim ans As Integer
    ans = Handler.ask("Преобразовать базу " & DataBaseName(this.type) & "?")
    
    If ans = vbNo Then _
        Exit Function
        
    Select Case True
    
        Case this.type = organisations
            RefactorOrganisations oldDataBaseArray
        
        Case this.type = instruments
            RefactorInstruments oldDataBaseArray
        
        Case this.type = persons
            RefactorPersons oldDataBaseArray
        
        Case Else
            Handler.Notify ("Преобразование недоступно для этого типа базы данных.")
            Exit Function
        
    End Select

    
    Refactor = True
    
End Function
    Private Function DataBaseName( _
        typeDb As EDadaBaseTypes _
        ) As String
        
        Select Case True
    
            Case typeDb = organisations
                DataBaseName = "организаций"
    
            Case typeDb = instruments
                DataBaseName = "средств измерений"
    
            Case typeDb = etalons
                DataBaseName = "эталонов"
    
            Case typeDb = persons
                DataBaseName = "сотрудников"
    
        End Select
    
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 12.03.2023 14:04
    ' ----------------------------------------------------------------
    Private Sub RefactorOrganisations( _
        oldDataBaseArray() As String _
        )

        Dim i As Integer, _
            item As cItemOrganisation
            
        Dim pb As New cProgressBar
        pb.Initialize UBound(oldDataBaseArray, 2) + 1, "Рефакторинг базы контрагентов"
            
        For i = LBound(oldDataBaseArray, 2) To UBound(oldDataBaseArray, 2)
            
            DoEvents
            pb.ShowProgress
            
            Set item = New cItemOrganisation
            item.shortName = CommasCorrected(oldDataBaseArray(0, i))
            item.taxNumber = oldDataBaseArray(1, i)
            item.legalAddress = oldDataBaseArray(3, i)
            
            If IsRussian(item.legalAddress) Then _
                item.country = "Россия"
            
            Dim itemKey As String
            itemKey = CommasCorrected(oldDataBaseArray(2, i))
            
            AddItem itemKey, item
            
        Next
        
        pb.Terminate
        ReCacheData
        
        Explorer.OpenFolder _
            path:=Config.organisationsPath, _
            isFocusOnWindow:=True
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 26.03.2023 18:46
        ' ----------------------------------------------------------------
        Private Function IsRussian( _
            address As String _
            ) As Boolean
            
            IsRussian = False
            
            If CBool(InStr(address, "USA")) Then _
                Exit Function
                
            If CBool(InStr(address, "Finland")) Then _
                Exit Function
            
            If CBool(InStr(address, "Беларусь")) Then _
                Exit Function
            
            IsRussian = True
            
        End Function
        Private Function CommasCorrected( _
            sTempStr As String _
            ) As String
            
            Dim i As Integer
            For i = 0 To Len(sTempStr) - 1
        
                If Asc(Mid(sTempStr, i + 1, 1)) = 34 Then 'опознана кавычка
                
                    Select Case i
                    
                        Case 0
                            sTempStr = Chr(171) & Right(sTempStr, Len(sTempStr) - 1) 'открывающаяся кавычка в начале
                            
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
             
            CommasCorrected = sTempStr
            
        End Function
    ' ----------------------------------------------------------------
    ' Дата: 18.03.2023 14:02
    ' ----------------------------------------------------------------
    Private Sub RefactorInstruments( _
        oldDataBaseArray() As String _
        )
        
        Dim pb As New cProgressBar
        pb.Initialize UBound(oldDataBaseArray, 2) + 1, "Рефакторинг базы СИ"
        
        Dim i As Integer, mi As cItemInstrument
        For i = LBound(oldDataBaseArray, 2) To UBound(oldDataBaseArray, 2)
            
            DoEvents
            pb.ShowProgress
            
            Set mi = New cItemInstrument
            
            mi.name = oldDataBaseArray(0, i)
            If mi.name = "-" Then _
                mi.name = Base.defaultValue
            
            mi.typeDevice = CommasCorrected(oldDataBaseArray(1, i))
            mi.typeDevice = CommasCorrected(mi.typeDevice)

            mi.regFgisNum = oldDataBaseArray(2, i)
            mi.methodic = CommasCorrected(oldDataBaseArray(3, i))
            
            AddItem CheckedName(mi.typeDevice), mi
            
            Dim wipeDirKey As String
            wipeDirKey = WipeName(CheckedName(mi.typeDevice), mi.regFgisNum)
            
            If Not wipeDirKey = Base.defaultValue Then
                ' ----------------------------------------------------------------
                Merge _
                    resultKey:=CheckedName(mi.typeDevice), _
                    wipeKey:=wipeDirKey, _
                    silentMode:=True
                ' ----------------------------------------------------------------
            End If
            
            
        Next
        
        pb.Terminate
        
        ReCacheData
        
        Explorer.OpenFolder _
            path:=Config.instrumentsPath, _
            isFocusOnWindow:=True
            
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 18.03.2023 18:54
        ' ----------------------------------------------------------------
        Private Function WipeName( _
            key As String, _
            subKey As String _
            ) As String
            
            WipeName = Base.defaultValue
            
            Dim result As String
            result = Dir(this.pathDb & "\*" & subKey & "*", vbDirectory)
            'result = Dir(this.pathDb & "\*" & key & "*" & subKey & "*", vbDirectory)
            
            If result <> vbNullString Then _
                WipeName = result
            
        End Function
        ' ----------------------------------------------------------------
        ' Дата: 18.03.2023 18:27
        ' ----------------------------------------------------------------
        Private Function CheckedName( _
            name As String _
            ) As String 'функция устраняет в строке запрещённые в имени файла символы
        
            name = Replace(name, "\", REPLACEMENT_SAFETY): name = Replace(name, "/", REPLACEMENT_SAFETY)
            name = Replace(name, ":", REPLACEMENT_SAFETY): name = Replace(name, "*", REPLACEMENT_SAFETY)
            name = Replace(name, "?", REPLACEMENT_SAFETY): name = Replace(name, "<", REPLACEMENT_SAFETY)
            name = Replace(name, ">", REPLACEMENT_SAFETY): name = Replace(name, "|", REPLACEMENT_SAFETY)
            name = Replace(name, """", REPLACEMENT_SAFETY)
        
            CheckedName = name
        
        End Function
    ' ----------------------------------------------------------------
    ' Дата: 12.03.2023 21:43
    ' ----------------------------------------------------------------
    Private Sub RefactorPersons( _
        oldDataBaseArray() As String _
        )
        
        Dim i As Integer, _
            person As cItemPerson
            
        Dim pb As New cProgressBar
        pb.Initialize UBound(oldDataBaseArray, 2) + 1, "Рефакторинг базы контактов"
        
        For i = LBound(oldDataBaseArray, 2) To UBound(oldDataBaseArray, 2)
            
            DoEvents
            pb.ShowProgress
            
            Dim fullName As String
            fullName = oldDataBaseArray(0, i)
            
            Dim splitName() As String
            splitName = Split(fullName, " ")
            
            Dim key As String
            key = splitName(0) & " " & Left(splitName(1), 1) & Left(splitName(2), 1)
            
            Set person = New cItemPerson
            person.lastName = splitName(0)
            person.firstName = splitName(1)
            person.middleName = splitName(2)
            person.state = oldDataBaseArray(1, i)
            
            person.organisation = "ВНИИМ"

            AddItem key, person
            
        Next
        
        pb.Terminate
        ReCacheData
        
        Explorer.OpenFolder _
            path:=Config.personsPath, _
            isFocusOnWindow:=True
            
    End Sub

