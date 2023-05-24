VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemOrganisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ------------------------------------------------------
' Наименование: cItemOrganisation (Class Module)
' Назначение:
' Автор: akolodka
' Date: 24.02.2023 21:35
' ------------------------------------------------------
Option Explicit
Implements IDatabaseItem

Const DEFAULT_COUNTRY = "Россия"

Private Type TOrgranisationFields
    
    shortName As String
    fullName As String
    legalAddress As String
    taxNumber As String
    country As String

End Type

Private this As TOrgranisationFields

Private Sub Class_Initialize()
    
    this.shortName = Base.defaultValue
    this.fullName = Base.defaultValue
    this.legalAddress = Base.defaultValue
    this.taxNumber = Base.defaultValue
    this.country = DEFAULT_COUNTRY
    
End Sub
' ------------------------------------------------------

Friend Property Get shortName() As String
    shortName = this.shortName
End Property
Friend Property Let shortName(ByVal NewValue As String)
    this.shortName = NewValue
End Property
' ------------------------------------------------------
Friend Property Get fullName() As String
    fullName = this.fullName
End Property
Friend Property Let fullName(ByVal NewValue As String)
    this.fullName = NewValue
End Property
' ------------------------------------------------------
Friend Property Get legalAddress() As String
    legalAddress = this.legalAddress
End Property
Friend Property Let legalAddress(ByVal NewValue As String)
    this.legalAddress = NewValue
End Property
' ------------------------------------------------------
Friend Property Get taxNumber() As String
    taxNumber = this.taxNumber
End Property
Friend Property Let taxNumber(ByVal NewValue As String)
    this.taxNumber = NewValue
End Property
' ------------------------------------------------------
Friend Property Get country() As String
    country = this.country
End Property
Friend Property Let country(ByVal NewValue As String)
    this.country = NewValue
End Property
