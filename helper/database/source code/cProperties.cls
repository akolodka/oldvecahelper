VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cProperties (Class Module)
' Назначение: Доступ к свойствам рабочей книги Excel
' Автор: akolodka
' Date: 23.10.2022 18:17
' ------------------------------------------------------

Option Explicit

Const KEY_CATEGORY = "category"
Const KEY_KEYWORDS = "keywords"
Const KEY_COMMENTS = "comments"

Const KEY_VERSION = "document version"
Const KEY_BUILD = "revision number"
Const KEY_SUBJECT = "subject"

Const KEY_COMPANY = "company"

Private targetWb As Workbook
Private Sub Class_Initialize()
    Set targetWb = ActiveWorkbook
End Sub
Private Sub Class_Terminate()
    Set targetWb = Nothing
End Sub
' ------------------------------------------------------
Friend Function TargetWorkbook() As Workbook
    Set TargetWorkbook = targetWb
End Function
Friend Sub SetTargetWorkbook(ByVal NewValue As Workbook)
    Set targetWb = NewValue
End Sub
' ------------------------------------------------------
Friend Property Get Category() As String
    Category = targetWb.BuiltinDocumentProperties(KEY_CATEGORY)
End Property
Friend Property Let Category(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_CATEGORY) = NewValue
End Property
' ------------------------------------------------------
Friend Property Get Keywords() As String
    Keywords = targetWb.BuiltinDocumentProperties(KEY_KEYWORDS)
End Property
Friend Property Let Keywords(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_KEYWORDS) = NewValue
End Property
' ------------------------------------------------------
Friend Property Get Comments() As String
    Comments = targetWb.BuiltinDocumentProperties(KEY_COMMENTS)
End Property
Friend Property Let Comments(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_COMMENTS) = NewValue
End Property
' ------------------------------------------------------
Friend Property Get Version() As String
    Version = targetWb.BuiltinDocumentProperties(KEY_VERSION)
End Property
Friend Property Let Version(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_VERSION) = NewValue
End Property
' ------------------------------------------------------
Friend Property Get Build() As String
    Build = targetWb.BuiltinDocumentProperties(KEY_BUILD)
End Property
Friend Property Let Build(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_BUILD) = NewValue
End Property
' ------------------------------------------------------
Friend Property Get Subject() As String
    Subject = targetWb.BuiltinDocumentProperties(KEY_SUBJECT)
End Property
Friend Property Let Subject(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_SUBJECT) = NewValue
End Property
' ------------------------------------------------------
Public Property Get Company() As String
    Company = targetWb.BuiltinDocumentProperties(KEY_COMPANY)
End Property
Public Property Let Company(ByVal NewValue As String)
    targetWb.BuiltinDocumentProperties(KEY_COMPANY) = NewValue
End Property
