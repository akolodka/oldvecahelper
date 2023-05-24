VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemEtalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 18.03.2023 12:04
' Автор: akolodka
' Назначение:
' ----------------------------------------------------------------
Option Explicit
Implements IDatabaseItem

Private Type TEtalonsFields
    
    type As String
    factoryNumber As String
    
    regFgisNum As String
    numberDocument As String
    dateExpire As String

    comment As String

End Type

Public Enum EServiceStatuses
    
    legal
    illegal
    
End Enum

Private this As TEtalonsFields
        
Private Sub Class_Initialize()
    
    this.type = Base.defaultValue
    this.factoryNumber = Base.defaultValue
    
    this.regFgisNum = Base.defaultValue
    this.numberDocument = Base.defaultValue
    this.dateExpire = Base.defaultValue
    
    this.comment = Base.defaultValue
    
End Sub
' ----------------------------------------------------------------
Public Property Get typeDevice() As String
    typeDevice = this.type
End Property

Public Property Let typeDevice(ByVal NewValue As String)
    this.type = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get factoryNumber() As String
    factoryNumber = this.factoryNumber
End Property
Public Property Let factoryNumber(ByVal NewValue As String)
    this.factoryNumber = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get regFgisNum() As String
    regFgisNum = this.regFgisNum
End Property
Public Property Let regFgisNum(ByVal NewValue As String)
    this.regFgisNum = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get numberDocument() As String
    numberDocument = this.numberDocument
End Property
Public Property Let numberDocument(ByVal NewValue As String)
    this.numberDocument = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get dateExpire() As String
    dateExpire = this.dateExpire
End Property
Public Property Let dateExpire(ByVal NewValue As String)
    this.dateExpire = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get comment() As String
    comment = this.comment
End Property
Public Property Let comment(ByVal NewValue As String)
    this.comment = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get status() As EServiceStatuses
    
    status = illegal
    If Not IsDate(this.dateExpire) Then _
        Exit Property
    
    If Now > CDate(this.dateExpire) Then _
        Exit Property
    
    status = legal
    
End Property

