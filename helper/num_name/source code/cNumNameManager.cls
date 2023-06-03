VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNumNameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DELIMITER_PRENUM = "/"
Const DELIMITER_POSTNUM = "–"

Const NUMBER_FORMAT = "0000"
Const NUMBER_FORMAT_MEASUREMENT = "000"

Const SUFFIX_MEASUREMENT = "–И"
Const DELIMITER_FILENAME = "_"
Const FILENAME_NULLNUM = "TEMP"

Const EMP_DELITIMER = ", "

Private m_sourcePrefix As String
Private m_numberCore As String
Private m_numberSuffix As String

Private m_numberPrefix As String


Private Sub Class_Initialize()
    
    m_sourcePrefix = Config.numDepartment

    If Protocol.typeDoc = verifying Then _
        m_sourcePrefix = Config.numLaboratory
    
    If Protocol.typeDoc = calibration Then _
        m_sourcePrefix = Config.numLaboratory
    
    m_numberPrefix = m_sourcePrefix
    '=====================================================
    m_numberCore = CheckedProtocolNumber
    m_numberSuffix = Year(Date)
    
End Sub
     Private Function NumFormatFromType() As String
        
        NumFormatFromType = NUMBER_FORMAT

        If Protocol.typeDoc = measurement Then _
            NumFormatFromType = NUMBER_FORMAT_MEASUREMENT
        
    End Function
    Private Function CheckedProtocolNumber( _
        ) As String
        
        CheckedProtocolNumber = vbNullString
        
        Dim sourceNumber As String
        sourceNumber = Cache.cachedNumber
        
        If sourceNumber = Base.defaultValue Then _
            Exit Function 'по умолчанию пустое значение
        
        CheckedProtocolNumber = Format(sourceNumber, NumFormatFromType)
    
    End Function
'=====================================================
Friend Property Get numberCore() As String
    numberCore = CheckedProtocolNumber
End Property
Friend Property Let numberCore( _
    ByVal NewValue As String _
    )

    m_numberCore = Format(NewValue, NumFormatFromType) 'может быть vbnullstring
    Cache.cachedNumber = Replace(NewValue, vbNullString, Base.defaultValue)
    
End Property
'=====================================================
Friend Property Get numberSuffix( _
    ) As String
    
    numberSuffix = m_numberSuffix
        
End Property
Friend Property Let numberSuffix( _
    ByVal NewValue As String _
    )
    
    m_numberSuffix = Year(Date)
    
    If NewValue <> vbNullString Then _
        m_numberSuffix = Format(NewValue, NUMBER_FORMAT)
        
End Property
' ----------------------------------------------------------------
' Наименование: ProtocolFileMask (Friend Function)
' Назначение: Новое имя файла
' Возвращаемый тип: String
' Дата: 17.02.2023 23:05
' ----------------------------------------------------------------
Friend Function ProtocolFileMask( _
    ) As String

    ProtocolFileMask = vbNullString
    
    If Protocol.typeDoc = unknown Then _
        Exit Function
        
    Dim checkedNum As String
    checkedNum = FILENAME_NULLNUM
    
    If m_numberCore <> vbNullString Then _
        checkedNum = CheckedProtocolNumber
        
    Dim cuttedSuffix As String
    cuttedSuffix = Right(m_numberSuffix, 2)
        
    Dim result As String
    result = Protocol.filePrefix & _
             m_numberPrefix & DELIMITER_FILENAME & _
             checkedNum & DELIMITER_FILENAME & _
             cuttedSuffix
    
'    If Protocol.typeDoc = measurement Then _
'        result = result & SUFFIX_MEASUREMENT
        
    ProtocolFileMask = result
    
End Function
'=====================================================
Friend Property Get numberPrefix() As String
    numberPrefix = m_numberPrefix
End Property

Friend Property Let numberPrefix(ByVal NewValue As String)
    
    m_numberPrefix = m_sourcePrefix
    If NewValue <> vbNullString Then _
        m_numberPrefix = NewValue
        
End Property
' ----------------------------------------------------------------
' Дата: 18.02.2023 19:16
' Назначение: Полный номер для передачи в протокол 2104/0000-2023
' Возвращаемый тип: String
' ----------------------------------------------------------------
Friend Function FullNumber( _
    ) As String
    
    Dim result As String
    result = m_numberPrefix & DELIMITER_PRENUM & _
             m_numberCore & DELIMITER_POSTNUM & _
             m_numberSuffix
    
    If Protocol.typeDoc = measurement Then _
        result = result & SUFFIX_MEASUREMENT
    
    FullNumber = result

End Function
' ----------------------------------------------------------------
' Дата: 18.02.2023 15:40
' Назначение: Промежуточная процедура для организации работы
' ----------------------------------------------------------------
Friend Function GetFromJournal( _
    ) As Boolean
    
    GetFromJournal = False
    
    Dim srcData As cJournalItem
    Set srcData = JournalRegData

    Dim jrData As cJournalItem, _
        journal As New cJournalManager
        
    Dim status As Boolean
    Set jrData = journal.RegisterNewItem(srcData, status)

    If status Then _
        numberCore = ProtocolShortNumber(jrData.protocolFullNumber): _
        GetFromJournal = True
    
End Function
    ' ----------------------------------------------------------------
    ' Дата: 18.02.2023 18:59
    ' Назначение: данные для регистрации в журнале
    ' Возвращаемый тип: cJournalItem
    ' ----------------------------------------------------------------
    Private Function JournalRegData( _
        ) As cJournalItem
        
        Dim result As New cJournalItem
        
        result.organisation = Protocol.organisation
        result.typeInstrument = Protocol.typeInstrument
        result.factoryNumber = Protocol.factoryNumber
        result.employees = EmployeeData
        
        Set JournalRegData = result
        
    End Function
        Private Function EmployeeData( _
            ) As String
            
            EmployeeData = Base.defaultValue
            
            Dim result As String
            result = Config.empMajor
            
            If Cache.useEmpSecond Then _
                result = result & EMP_DELITIMER & Config.empSecond
            
            If Cache.useEmpThird Then _
                result = result & EMP_DELITIMER & Config.empThird
                
            If Cache.useEmpOther Then _
                result = result & EMP_DELITIMER & Cache.empOther
            
            EmployeeData = result
            
        End Function
    ' ----------------------------------------------------------------
    ' Дата: 18.02.2023 15:50
    ' Назначение: выделить сквозной номер из полного номера протокола
    '    параметр protFullNumber: 2104/0000-2023
    ' Возвращаемый тип: String
    ' ----------------------------------------------------------------
    Private Function ProtocolShortNumber( _
        protFullNumber As String _
        ) As String

        ProtocolShortNumber = Base.defaultValue
        
        Dim result As String
        result = protFullNumber
        
        Dim instrData As Integer
        instrData = InStr(protFullNumber, DELIMITER_PRENUM)
        
        If CBool(instrData) Then _
        result = Mid(protFullNumber, instrData + 1, 4)
        
        ProtocolShortNumber = result
        
    End Function
' ----------------------------------------------------------------
' Дата: 17.04.2023 09:35
' ----------------------------------------------------------------
Public Function PutInJournal( _
    data As cJournalItem _
    ) As Boolean
    
    PutInJournal = False
    
    Dim jr As New cJournalManager
    jr.RegisterNewItem data, PutInJournal
    
End Function
'todo:[+] NumName Register In Journal
