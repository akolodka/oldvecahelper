VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 12.03.2023 21:29
' Автор: akolodka
' Назначение:
' ----------------------------------------------------------------
Option Explicit

Implements IDatabaseItem

Const EMAIL_DEFAULT As String = "email"
Const PHONE_DEFAULE As String = "phone"

Private Type TPersonFields
    
    lastName As String
    firstName As String
    middleName As String
    
    state As String
    organisation As String
    
    birthDay As String
    
    phoneNumber As String
    email As String
    
End Type

Private this As TPersonFields
        
Private Sub Class_Initialize()
    
    this.lastName = Base.defaultValue
    this.firstName = Base.defaultValue
    this.middleName = Base.defaultValue
    
    this.state = Base.defaultValue
    this.organisation = Base.defaultValue
    
    this.phoneNumber = PHONE_DEFAULE
    this.email = EMAIL_DEFAULT
    
    this.birthDay = Base.defaultValue
    
End Sub
' ------------------------------------------------------
Public Property Get lastName() As String
    lastName = this.lastName
End Property
Public Property Let lastName(ByVal NewValue As String)
    this.lastName = NewValue
End Property
' ------------------------------------------------------
Public Property Get firstName() As String
    firstName = this.firstName
End Property
Public Property Let firstName(ByVal NewValue As String)
    this.firstName = NewValue
End Property
' ------------------------------------------------------
Public Property Get middleName() As String
    middleName = this.middleName
End Property
Public Property Let middleName(ByVal NewValue As String)
    this.middleName = NewValue
End Property
' ------------------------------------------------------
Public Property Get state() As String
    state = this.state
End Property
Public Property Let state(ByVal NewValue As String)
    this.state = NewValue
End Property
' ------------------------------------------------------
Public Property Get organisation() As String
    organisation = this.organisation
End Property
Public Property Let organisation(ByVal NewValue As String)
    this.organisation = NewValue
End Property
' ------------------------------------------------------
Public Property Get birthDay() As String
    birthDay = this.birthDay
End Property
Public Property Let birthDay(ByVal NewValue As String)
    this.birthDay = NewValue
End Property
' ------------------------------------------------------
Public Property Get email() As String
    email = this.email
End Property
Public Property Let email(ByVal NewValue As String)
    this.email = NewValue
End Property
' ------------------------------------------------------
Public Property Get phoneNumber() As String
    phoneNumber = this.phoneNumber
End Property
Public Property Let phoneNumber(ByVal NewValue As String)
    this.phoneNumber = NewValue
End Property
