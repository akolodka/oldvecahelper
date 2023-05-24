VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemKeysInstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 18.03.2023 13:40
' Автор: akolodka
' Назначение:
' ----------------------------------------------------------------
Option Explicit
    
Const FGIS_REG_NUMBER_KEY = "Номер в ФИФ"

Const NAME_KEY = "Наменование"
Const TYPE_KEY = "Обозначение типа"
Const METHODIC_KEY = "Методика"

Const MANUFACTURER_KEY = "Изготовитель"
Const DATE_EXPIRE_KEY = "Срок свидетельства"

Public Property Get regFgisNum() As String
    regFgisNum = FGIS_REG_NUMBER_KEY
End Property
' ----------------------------------------------------------------
Public Property Get name() As String
    name = NAME_KEY
End Property
Public Property Get typeDevice() As String
    typeDevice = TYPE_KEY
End Property
Public Property Get methodic() As String
    methodic = METHODIC_KEY
End Property
' ----------------------------------------------------------------
Public Property Get manufacturer() As String
    manufacturer = MANUFACTURER_KEY
End Property
Public Property Get dateExpire() As String
    dateExpire = DATE_EXPIRE_KEY
End Property

