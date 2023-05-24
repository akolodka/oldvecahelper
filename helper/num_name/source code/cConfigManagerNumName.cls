VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfigManagerNumName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 09.03.2023 17:15
' Автор: akolodka
' Назначение: Обработчик конфигурации
' ----------------------------------------------------------------
Option Explicit

Private cs As New cConfigSource
Private fso As New FileSystemObject
Private keys As New cConfigKeysNumName

Private m_journalPath As String
        
Private m_numLaboratory As String, _
        m_numDepartment As String

Private m_empMajor As String, _
        m_empSecond As String, _
        m_empThird As String
' ----------------------------------------------------------------
Private Sub Class_Initialize()

    cs.Initialize _
        sourceType:=configData, _
        nameKey:=keys.header

    m_numLaboratory = cs.FindValue(keys.laboratoryNumber)
    m_numDepartment = cs.FindValue(keys.departmentNumber)
    
    m_journalPath = cs.FindValue(keys.journalPath)

    m_empMajor = cs.FindValue(keys.employeeMajor)
    m_empSecond = cs.FindValue(keys.employeeSecond)
    m_empThird = cs.FindValue(keys.employeeThird)
    
End Sub
Private Sub Class_Terminate()
    
    Set cs = Nothing
    Set fso = Nothing
    Set keys = Nothing
    
End Sub
' ----------------------------------------------------------------
Public Property Get numLaboratory() As String
    numLaboratory = m_numLaboratory
End Property
Friend Property Let numLaboratory(ByVal NewValue As String)
    
    m_numLaboratory = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        m_numLaboratory = NewValue
    
End Property
' ----------------------------------------------------------------
Friend Property Get numDepartment() As String
    numDepartment = m_numDepartment
End Property
Friend Property Let numDepartment(ByVal NewValue As String)
    
    m_numDepartment = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        m_numDepartment = NewValue
    
End Property
' ----------------------------------------------------------------
Public Property Get journalPath( _
    ) As String
    
    journalPath = m_journalPath
    
End Property
Friend Property Let journalPath( _
    ByVal NewValue As String _
    )
    
    m_journalPath = Base.defaultValue
    
    If NewValue <> vbNullString Then _
        m_journalPath = NewValue
        
End Property
' ----------------------------------------------------------------
Friend Sub Save()
    
    cs.PrepareToWrite 'обязательно, чтобы не дублировались ключи
    
    cs.PrepareToWrite keys.laboratoryNumber, m_numLaboratory
    cs.PrepareToWrite keys.departmentNumber, m_numDepartment
    cs.PrepareToWrite keys.journalPath, m_journalPath
    cs.PrepareToWrite keys.employeeMajor, m_empMajor
    cs.PrepareToWrite keys.employeeSecond, m_empSecond
    cs.PrepareToWrite keys.employeeThird, m_empThird
    
    cs.WriteToFile
    
End Sub
Public Function IsCorrect() As Boolean
    
    IsCorrect = False
    
    If m_numDepartment = Base.defaultValue Then _
        Exit Function
    
    If m_numLaboratory = Base.defaultValue Then _
        Exit Function
    
    If Not fso.FileExists(m_journalPath) Then _
        Exit Function
    
    IsCorrect = True
                
End Function
' ----------------------------------------------------------------
Friend Property Get empMajor() As String
    empMajor = m_empMajor
End Property
Friend Property Let empMajor(ByVal NewValue As String)
    m_empMajor = NewValue
End Property
' ----------------------------------------------------------------
Friend Property Get empSecond() As String
    empSecond = m_empSecond
End Property
Friend Property Let empSecond(ByVal NewValue As String)
    m_empSecond = NewValue
End Property
' ----------------------------------------------------------------
Friend Property Get empThird() As String
    empThird = m_empThird
End Property
Friend Property Let empThird(ByVal NewValue As String)
    m_empThird = NewValue
End Property
' ----------------------------------------------------------------
' ----------------------------------------------------------------
' Наименование: ChooseJournalPath (Public Sub)
' Назначение: Выбор файла журнала
'    параметр initialSomniumPath:
' Дата: 31.01.2023 21:32
' ----------------------------------------------------------------
Friend Function ChooseJournalPath( _
    Optional initialSomniumPath As Boolean = False _
    ) As Boolean
    
    ChooseJournalPath = False
    
    Dim fd As fileDialog
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    fd.Filters.Clear
    fd.InitialView = msoFileDialogViewDetails
    
    fd.Title = "Выбор источника данных"
    fd.AllowMultiSelect = False
    
    fd.InitialFileName = Base.desktopPath
    
    If initialSomniumPath Then _
        fd.InitialFileName = Base.serverInitialPath

    fd.Filters.Add "Книги Excel", "*.xls*", 1

    If fd.Show = False Then _
        Set fd = Nothing: _
        Exit Function
        
    m_journalPath = fd.SelectedItems(1) 'полный путь к файлу
    
    Set fd = Nothing
    ChooseJournalPath = True
    
End Function
