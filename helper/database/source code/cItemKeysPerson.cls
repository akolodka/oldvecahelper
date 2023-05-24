VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemKeysPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 12.03.2023 21:18
' Автор: akolodka
' Назначение: Ключи для бд сотрудников
' ----------------------------------------------------------------
Option Explicit

Const LASTNAME_KEY = "Фамилия"
Const FIRSTNAME_KEY = "Имя"
Const MIDDLENAME_KEY = "Отчество"
Const STATE_KEY = "Должность"
Const PHONE_NUMBER_KEY = "Телефон"
Const EMAIL_KEY = "Email"
Const ORGANISATION_KEY = "Организация"
Const BIRTHDAY_KEY = "День рождения"

Friend Property Get lastName() As String
    lastName = LASTNAME_KEY
End Property

Friend Property Get firstName() As String
    firstName = FIRSTNAME_KEY
End Property

Friend Property Get middleName() As String
    middleName = MIDDLENAME_KEY
End Property

Friend Property Get state() As String
    state = STATE_KEY
End Property

Friend Property Get organisation() As String
    organisation = ORGANISATION_KEY
End Property

Friend Property Get birthDay() As String
    birthDay = BIRTHDAY_KEY
End Property

Friend Property Get phoneNumber() As String
    phoneNumber = PHONE_NUMBER_KEY
End Property

Friend Property Get email() As String
    email = EMAIL_KEY
End Property
