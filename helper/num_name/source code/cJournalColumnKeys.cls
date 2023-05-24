VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJournalColumnKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ORGANISATION_KEY = "Заказчик"
Const TYPE_KEY = "Тип СИ"
Const SERIAL_KEY = "Зав"

Const PROTOCOL_NUMBER_KEY = "Номер протокола"
Const DATE_KEY = "Дата"
Const FGIS_DOC_KEY = "Аршин"
Const LINK_KEY = "Ссылка"
Const ENTRY_NUM_KEY = "омер запис"
Const EMPLOYEE_KEY = "Исполнители"

Friend Property Get organisation() As String
    organisation = ORGANISATION_KEY
End Property

Friend Property Get typeInstrument() As String
    typeInstrument = TYPE_KEY
End Property

Friend Property Get serial() As String
    serial = SERIAL_KEY
End Property

Friend Property Get ProtocolNumber() As String
    ProtocolNumber = PROTOCOL_NUMBER_KEY
End Property

Friend Property Get dateVerifying() As String
    dateVerifying = DATE_KEY
End Property

Friend Property Get fgisDocumentNumber() As String
    fgisDocumentNumber = FGIS_DOC_KEY
End Property

Friend Property Get fgisLink() As String
    fgisLink = LINK_KEY
End Property

Friend Property Get fgisEntryNumber() As String
    fgisEntryNumber = ENTRY_NUM_KEY
End Property

Friend Property Get employees() As String
    employees = EMPLOYEE_KEY
End Property
