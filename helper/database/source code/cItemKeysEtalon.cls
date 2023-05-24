VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemKeysEtalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 17.03.2023 11:11
' Автор: akolodka
' Назначение: ключи доступа к элементу базы эталонов
' ----------------------------------------------------------------

Option Explicit

Const TYPE_DEVICE_KEY = "Обозначение типа"
Const FACTORY_NUMBER_KEY = "Заводской номер"

Const FGIS_REG_NUMBER_KEY = "Рег. номер ФИФ"
Const NUMBER_DOCUMENT_KEY = "Номер документа"
Const DATE_EXPIRE_KEY = "Срок действия"

Const COMMENT_KEY = "Комментарий"
Public Property Get typeDevice() As String
    typeDevice = TYPE_DEVICE_KEY
End Property
Public Property Get factoryNumber() As String
    factoryNumber = FACTORY_NUMBER_KEY
End Property
' ----------------------------------------------------------------
Public Property Get regFgisNum() As String
    regFgisNum = FGIS_REG_NUMBER_KEY
End Property
Public Property Get numberDocument() As String
    numberDocument = NUMBER_DOCUMENT_KEY
End Property
Public Property Get dateExpire() As String
    dateExpire = DATE_EXPIRE_KEY
End Property
' ----------------------------------------------------------------
Public Property Get comment() As String
    comment = COMMENT_KEY
End Property
