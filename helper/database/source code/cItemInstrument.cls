VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemInstrument"
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
Implements IDatabaseItem

Private Type TInstrumentFields
    
    regFgisNum As String
    name As String
    type As String
    methodic As String
    
    modification As String
    factoryNumber As String
    
    dateExpire As String
    manufacturer As String
    
End Type

Private this As TInstrumentFields
Private Sub Class_Initialize()
    
    this.regFgisNum = Base.defaultValue
    this.name = Base.defaultValue
    this.type = Base.defaultValue
    this.methodic = Base.defaultValue
    
    this.modification = Base.defaultValue
    this.factoryNumber = Base.defaultValue
    
    this.dateExpire = Base.defaultValue
    this.manufacturer = Base.defaultValue
    
End Sub
' ----------------------------------------------------------------
Public Property Get regFgisNum() As String
    regFgisNum = this.regFgisNum
End Property
Public Property Let regFgisNum(ByVal NewValue As String)
    this.regFgisNum = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get name() As String
    name = this.name
End Property
Public Property Let name(ByVal NewValue As String)
    this.name = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get typeDevice() As String
    typeDevice = this.type
End Property
Public Property Let typeDevice(ByVal NewValue As String)
    this.type = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get methodic() As String
    methodic = this.methodic
End Property
Public Property Let methodic(ByVal NewValue As String)
    this.methodic = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get dateExpire() As String
    dateExpire = this.dateExpire
End Property
Public Property Let dateExpire(ByVal NewValue As String)
    this.dateExpire = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get manufacturer() As String
    manufacturer = this.manufacturer
End Property
Public Property Let manufacturer(ByVal NewValue As String)
    this.manufacturer = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get factoryNumber() As String
    factoryNumber = this.factoryNumber
End Property
Public Property Let factoryNumber(ByVal NewValue As String)
    this.factoryNumber = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get modification() As String
    modification = this.modification
End Property
Public Property Let modification(ByVal NewValue As String)
    this.modification = NewValue
End Property
