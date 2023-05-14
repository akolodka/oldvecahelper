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

Const PROPERTY_CATEGORY = "category"
Const PROPERTY_KEYWORDS = "keywords"
Const PROPERTY_COMMENTS = "comments"

Const PROPERTY_VERSION = "document version"
Const PROPERTY_BUILD = "revision number"
Const PROPERTY_SUBJECT = "subject"

Public Enum ETargetWorkbooks
    thisWb
    activeWb
End Enum

Private m_targetWb As Workbook

Private Sub Class_Terminate()
    Set m_targetWb = Nothing
End Sub

Friend Function TargetWorkbook() As Workbook
    TargetWorkbook = m_targetWb
End Function

Friend Sub Initialize( _
    targetWb As ETargetWorkbooks _
    )
    
    If targetWb = thisWb Then _
        Set m_targetWb = ThisWorkbook
        
    If targetWb = activeWb Then _
        Set m_targetWb = ActiveWorkbook
        
End Sub

Friend Property Get category() As String
    category = m_targetWb.BuiltinDocumentProperties(PROPERTY_CATEGORY)
End Property
Friend Property Let category(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_CATEGORY) = NewValue
End Property

Friend Property Get keywords() As String
    keywords = m_targetWb.BuiltinDocumentProperties(PROPERTY_KEYWORDS)
End Property
Friend Property Let keywords(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_KEYWORDS) = NewValue
End Property

Friend Property Get comments() As String
    comments = m_targetWb.BuiltinDocumentProperties(PROPERTY_COMMENTS)
End Property
Friend Property Let comments(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_COMMENTS) = NewValue
End Property

Friend Property Get version() As String
    version = m_targetWb.BuiltinDocumentProperties(PROPERTY_VERSION)
End Property
Friend Property Let version(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_VERSION) = NewValue
End Property

Friend Property Get build() As String
    build = m_targetWb.BuiltinDocumentProperties(PROPERTY_BUILD)
End Property
Friend Property Let build(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_BUILD) = NewValue
End Property

Friend Property Get subject() As String
    subject = m_targetWb.BuiltinDocumentProperties(PROPERTY_SUBJECT)
End Property
Friend Property Let subject(ByVal NewValue As String)
    m_targetWb.BuiltinDocumentProperties(PROPERTY_SUBJECT) = NewValue
End Property
