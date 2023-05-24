VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cErrorMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 25.03.2023 16:44
' Автор: akolodka
' Назначение:
' ----------------------------------------------------------------
Option Explicit

Public Property Get journalBusy( _
    userName As String _
    ) As String
    
    journalBusy = "Прервано: журнал занят «" & userName & "»."
End Property

Public Property Get journalNotFinded() As String
    journalNotFinded = "Прервано: журнал недоступен по адресу " & Config.journalPath & "."
End Property

Public Property Get columnTypeNotFinded() As String
    columnTypeNotFinded = "Прервано: в файле журнала отсутствует столбец типа СИ."
End Property

Public Property Get columnSerialNotFinded() As String
    columnSerialNotFinded = "Прервано: в файле журнала отсутствует столбец заводского номера СИ."
End Property

Public Property Get columnProtocolNumberNotFinded() As String
    columnProtocolNumberNotFinded = "Прервано: в файле журнала отсутствует столбец номера протокола."
End Property

Public Property Get columnFgisNumberNotFinded() As String
    columnFgisNumberNotFinded = "Прервано: в файле журнала отсутствует столбец номера документа ФГИС."
End Property

Public Property Get columnFgisLinkNotFinded() As String
    columnFgisLinkNotFinded = "Прервано: в файле журнала отсутствует столбец ссылки на запись во ФГИС."
End Property

Public Property Get incorrectTypeInstrument() As String
    incorrectTypeInstrument = "Прервано: некорректный тип СИ."
End Property

Public Property Get incorrectFactoryNumber() As String
    incorrectFactoryNumber = "Прервано: некорректный заводской номер СИ."
End Property

Public Property Get incorrectProtocolNumber() As String
    incorrectProtocolNumber = "Прервано: некорректный номер протокола."
End Property

Public Property Get incorrectFgisDocumentNumber() As String
    incorrectFgisDocumentNumber = "Прервано: некорректный номер документа ФГИС."
End Property

Public Property Get incorrectFgisLinkNumber() As String
    incorrectFgisLinkNumber = "Прервано: некорректный номер записи ФГИС."
End Property

Public Property Get incorrectDateVerifying() As String
    incorrectDateVerifying = "Прервано: некорректная дата поверки."
End Property
