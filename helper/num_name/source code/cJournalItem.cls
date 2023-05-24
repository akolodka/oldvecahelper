VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJournalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cJournalItem (Class Module)
' Назначение: Регистрационный элемент журнала поверок
' Автор: akolodka
' Date: 27.01.2023 22:22
' ------------------------------------------------------

Option Explicit

Private m_organisation As String, _
        m_typeInstrument As String, _
        m_factoryNumber As String, _
        m_protocolFullNumber As String

Private m_dateVerifying As String, _
        m_fgisDocNumber As String, _
        m_employees As String, _
        m_fgisLinkNumber As String


Private Sub Class_Initialize()

    m_organisation = Base.defaultValue
    m_typeInstrument = Base.defaultValue
    m_factoryNumber = Base.defaultValue
    
    m_dateVerifying = Base.defaultValue
    m_fgisDocNumber = Base.defaultValue
    m_employees = Base.defaultValue
    
    m_protocolFullNumber = Base.defaultValue
    
End Sub
'###################################################################
Friend Property Get organisation() As String
    organisation = m_organisation
End Property
Friend Property Let organisation(ByVal NewValue As String)
    m_organisation = NewValue
End Property
'###################################################################
Friend Property Get typeInstrument() As String
    typeInstrument = m_typeInstrument
End Property
Friend Property Let typeInstrument(ByVal NewValue As String)
    m_typeInstrument = NewValue
End Property
'###################################################################
Friend Property Get factoryNumber() As String
    factoryNumber = m_factoryNumber
End Property
Friend Property Let factoryNumber(ByVal NewValue As String)
    m_factoryNumber = NewValue
End Property
'###################################################################
Friend Property Get dateVerifying() As String
    dateVerifying = m_dateVerifying
End Property
Friend Property Let dateVerifying(ByVal NewValue As String)
    m_dateVerifying = NewValue
End Property
'###################################################################
Friend Property Get fgisDocNumber() As String
    fgisDocNumber = m_fgisDocNumber
End Property
Friend Property Let fgisDocNumber(ByVal NewValue As String)
    m_fgisDocNumber = NewValue
End Property
'###################################################################
Friend Property Get employees() As String
    employees = m_employees
End Property
Friend Property Let employees(ByVal NewValue As String)
    m_employees = NewValue
End Property
'###################################################################
Friend Property Get protocolFullNumber() As String
    protocolFullNumber = m_protocolFullNumber
End Property
Friend Property Let protocolFullNumber(ByVal NewValue As String)
    m_protocolFullNumber = NewValue
End Property
'###################################################################
Friend Property Get fgisLinkNumber() As String
    fgisLinkNumber = m_fgisLinkNumber
End Property
Friend Property Let fgisLinkNumber(ByVal NewValue As String)
    m_fgisLinkNumber = NewValue
End Property
