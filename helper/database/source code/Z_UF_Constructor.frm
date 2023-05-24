VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_UF_Constructor 
   Caption         =   "Конструктор базы данных заказчиков / средств измерений / эталонов"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9315
   OleObjectBlob   =   "Z_UF_Constructor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Z_UF_Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit 'запрет на использование неявных переменных

Const DEFAULT_VALUE = "nodata"
Const REFERENCE_FIFNUM_FILENAME = "fifRegNum.ref"

Private myBase As New Z_clsmBase, WorkClsm As New Z_clsmConstructor, sConstrSrch As String 'данные для быстрого перехода в конструкторе
Dim sArrDataBase() As String, sArrKeyCode() As String, bolUserInput As Boolean, sLabelText As String, bolTbEntry As Boolean
Dim iDataBaseIndex As Integer, sArrSelResults() As String


Private Sub UserForm_Initialize()
    Set_UF_Cunstructor_Properties 'изменить свойства в соответствии с загружаемым типом формы
    GetMyConfigFromFile 'загрузить параметры в модуль классы
        
    FillArrKeycodeFromFile 'заполнить массив кейкодов
    FillArrDataBaseFromFile 'заполнить массив базы данных
    RefreshLabelInfo ' исходные данные информационной метки
' ------------------------------------------------------
'todo: отвязать workclsm
    If UMenu.typе = instrumentsOLD Then CheckTempDir WorkClsm.startDir, sArrDataBase
    If UMenu.typе = etalonsOLD Then SortMassBiv sArrDataBase, UBound(sArrDataBase) 'сортировать массив по ключевому слову
    
    bolAlreadySaved = True: UpdateListDataBase sArrDataBase 'заполнить листбокс данным массива организаций
    SetEventControls Me 'инициировать групповые события для всех контролов

    If Me.listDataBase.ListCount = 0 Then Me.cmbAdd.Enabled = False: Me.cmbAdd.caption = "Внести"
    ' ------------------------------------------------------
'todo: отвязать workclsm
    If Me.listDataBase.ListCount > 0 Then sConstrSrch = WorkClsm.constrSrch: Me.listDataBase.Selected(0) = True
    
    bolUF_Cnstr_Load = True: If sConstrSrch <> "недоступно" Then SelByConstrSrch sConstrSrch 'строка, выбранная в форме поиска
    bolUF_Cnstr_Load = False: Me.tboxSearchConstr = "": bolUF_Cnstr_Load = True 'очистить поле быстрого перехода
End Sub
'#########################################################
'процедура корректно выбирает последний поисковый запрос
Sub SelByConstrSrch(sConstrSrch As String)
    Dim i As Integer, iArrayDimention As Byte 'по умолчанию выбирается первое измерение = 0
    
    If UMenu.typе = instrumentsOLD Then iArrayDimention = 2 'номер в фиф
    If UMenu.typе = etalonsOLD Then iArrayDimention = 4 'ключевое слово

    For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
        If InStr(sArrDataBase(iArrayDimention, i), sConstrSrch) > 0 Then Me.listDataBase.Selected(i) = True: Exit Sub
    Next i
End Sub


'#########################################################
'процедура срабатывает при активации формы
Private Sub UserForm_Activate()
    Me.tboxSearchConstr.SetFocus
End Sub
'#########################################################
'процедура срабатывает при выгрузке формы
Private Sub UserForm_Terminate()
    bolUF_Cnstr_Load = False: bolAlreadySaved = False: Set myBase = Nothing: Set WorkClsm = Nothing
End Sub
'#########################################################
'загрузить конфигурацию из файлов
Sub GetMyConfigFromFile()
    

    With myBase
        .AddP "startDir", "templatesDir"
        
        Select Case UMenu.typе  'параметр загрузки элементов формы
            
            Case organisationsOLD
                .AddP "cusDB" 'заказчики
                
            Case instrumentsOLD
                .AddP "measInstrDB" 'средства измерений
                
            Case etalonsOLD
                .AddP "etalDB" 'эталоны
                
            Case personsOLD
                .AddP "empDB" 'фамилии и должности
                
        End Select
        
        .GetArrFF setDir, Environ("USERNAME") & ".uCfg" 'загрузить в класс конфигурацию
        .FillValues 'обязательно: найти значения выходных параметров по ключам
        
        WorkClsm.FillConfiguration .Parameters, .values 'передать конфигурацию в класс
        .ClearParameters
        
        .AddP "constrSrch": .GetArrFF setDir, "settings.ini" 'загрузить в класс конфигурацию
        .FillValues 'обязательно: найти значения выходных параметров по ключам
        
        WorkClsm.FillSettings .Parameters, .values 'передать конфигурацию в класс
    End With
End Sub
'#########################################################
'процедура заполняет массив кейкодов
Sub FillArrKeycodeFromFile()
    If FileExist(setDir, "keycode.npDb") Then 'если опознано наличие в каталоге надстройки файла кейкодов
 ' ------------------------------------------------------
'todo: отвязать workclsm
        sArrKeyCode = WorkClsm.FillDataBase( _
            myBase.GetArrFF(setDir, "keycode.npDb"))  'получить массив кейкодов(если файл обнаружен)
    Else 'если файл не был обнаружен и загружен
        ReDim sArrKeyCode(0, 1)
    End If
End Sub
'#########################################################
'процедура заполняет массив базы данных
Sub FillArrDataBaseFromFile()
' ------------------------------------------------------
'todo: отвязать workclsm

    With WorkClsm
        If FileExist(.startDir, .DbName) Then 'если база данных обнаружена по указанному пути
            sArrDataBase = .FillDataBase(myBase.GetArrFF(.startDir, .DbName), True)  'преобразовать массив файла в массив базы данных
            sLabelText = .DbName & ", " & RusPadejPozition(UBound(sArrDataBase, 2) + 1)
        Else 'файл БД не был загружен
            ReDim sArrDataBase(0)
            RefreshLabelInfo "Файл данных не загружен.", &H80& 'красный
        End If
    End With
End Sub

'#########################################################
'процедура обновляет сведения информационной метки
Sub RefreshLabelInfo(Optional sCaption As String, Optional sColor As String, Optional TrueElemFore As Boolean)

    With Me.LabelInfo
        .caption = sLabelText: TrueElementForeColor Me.LabelInfo
        
        If sCaption <> "" Then _
            .caption = sCaption: If sColor <> "" Then .foreColor = sColor
        
        If TrueElemFore Then TrueElementForeColor Me.LabelInfo
    End With
End Sub
'#########################################################
'быстрый фильтр
Private Sub tboxSearchConstr_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If bolTbEntry = False Then _
        With Me.tboxSearchConstr: .SelStart = 0: .SelLength = Len(.text): bolTbEntry = True: End With 'переменная первого входа
End Sub
Private Sub tboxSearchConstr_Change()
    
    If bolUF_Cnstr_Load = False Then Exit Sub
    If Me.tboxSearchConstr = "" Then Me.listDataBase.ListIndex = -1: Exit Sub 'если поле пустое, то убрать выделение строки
    
    Dim sInputRus As String, sInputEng As String, sInputString As String
    ReDim sArrSelResults(UBound(sArrDataBase), 0) 'выделить память под временный массив
    
    If sArrKeyCode(0, 0) = "" Then 'массив кейкодов не был загружен
        FindInBivArr sArrDataBase, sArrSelResults, Me.tboxSearchConstr  'получить массив поисковых совпадений
    Else ' массив кейкодов был загружен
        FillInputData sArrKeyCode, Me.tboxSearchConstr, sInputRus, sInputEng   'получить значения для поиска по массиву
        sInputString = sInputRus & " " & sInputEng: FindInBivArr sArrDataBase, sArrSelResults, sInputString  'получить массив поисковых совпадений значениями на русском языке
    End If
        
    TrueSelectionlistDataBase 'выделить элемент листбокса - первый поиск
End Sub
Private Sub tboxSearchConstr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0: TrueSelectionlistDataBase True 'второй и последующий поиск
End Sub
Private Sub tboxSearchConstr_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    bolTbEntry = False
End Sub

'#########################################################
'процедура выделяет элемент листбокса
Sub TrueSelectionlistDataBase(Optional bolSecondSearch As Boolean)

    Dim myInd As Integer, i As Integer, j As Integer, K As Integer
    myInd = -1 'по умолчанию
    
    With Me.listDataBase
        For i = LBound(.List) To UBound(.List)
            'при первом поиске найти совпадение по первому элементу найденных значений
            If InStr(.List(i), sArrSelResults(LBound(sArrSelResults), 0)) > 0 And _
                InStr(.List(i), sArrSelResults(LBound(sArrSelResults) + 1, 0)) > 0 And _
                    InStr(.List(i), sArrSelResults(LBound(sArrSelResults) + 2, 0)) > 0 Then myInd = i: Exit For
        Next i
    
        If bolSecondSearch Then 'поиск дополнительных совпадений
            K = Me.listDataBase.ListIndex
            
            For j = LBound(sArrSelResults, 2) To UBound(sArrSelResults, 2)
                'найти элемент в массиве поисковых значений
                If InStr(.List(K), sArrSelResults(LBound(sArrSelResults), j)) > 0 And _
                    InStr(.List(K), sArrSelResults(LBound(sArrSelResults) + 1, j)) > 0 And _
                        InStr(.List(K), sArrSelResults(LBound(sArrSelResults) + 2, j)) > 0 Then Exit For
            Next j

            If j = UBound(sArrSelResults, 2) Then
                j = LBound(sArrSelResults, 2) 'переход от последнего элемента к первому
            Else
                j = j + 1
            End If
            
            For i = LBound(.List) To UBound(.List)
                'при первом поиске найти совпадение по первому элементу найденных значений
                If InStr(.List(i), sArrSelResults(LBound(sArrSelResults), j)) > 0 And _
                    InStr(.List(i), sArrSelResults(LBound(sArrSelResults) + 1, j)) > 0 And _
                        InStr(.List(i), sArrSelResults(LBound(sArrSelResults) + 2, j)) > 0 Then myInd = i: Exit For
            Next i
        End If
    End With

    Me.listDataBase.ListIndex = -1 'убрать выделение строки
    If myInd >= 0 Then Me.listDataBase.Selected(myInd) = True
End Sub

'#########################################################
'процедура обновляет содержимое листбокса, исходя из представленного массива
Sub UpdateListDataBase(sArrDBase() As String) 'обновить содержимое листбокса1
    Me.listDataBase.Clear
    
    If UBound(sArrDBase) > 0 Then 'массив данных был получен
    
        Dim i As Integer, sTempStr As String
        For i = LBound(sArrDBase, 2) To UBound(sArrDBase, 2) 'для каждого блока
            sTempStr = vbNullString 'очищение переменной
            
            Select Case UMenu.typе
            
                Case instrumentsOLD 'средства измерений
                    sTempStr = GetListStringForInstruments(sArrDataBase, i) 'получить строку для базы данных средств измерений
                    
                Case Else
                    sTempStr = GetListStringOtherDB(sArrDataBase, i)
            End Select
            
            If sTempStr <> "" Then _
                sTempStr = TrueSpace(i + 1) & sTempStr: Me.listDataBase.AddItem sTempStr
        Next
    End If
End Sub
'#########################################################
'получить строку для базы данных средств измерений
Function GetListStringForInstruments(sArrDBase() As String, iDbIndex As Integer) As String
    
    Dim sTempStr As String
    If sArrDBase(LBound(sArrDBase), iDbIndex) <> "" Then
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 2, iDbIndex) & " \\ " 'номер в фиф
        
        If sArrDBase(UBound(sArrDBase), iDbIndex) <> "nodata" Then _
            sTempStr = sTempStr & sArrDBase(UBound(sArrDBase), iDbIndex) & " \\ " 'наличие шаблона для данного СИ
        
        If sArrDBase(LBound(sArrDBase) + 1, iDbIndex) <> DEFAULT_VALUE And _
            sArrDBase(LBound(sArrDBase) + 1, iDbIndex) <> "" Then _
            sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 1, iDbIndex) & " \\ " 'тип СИ
            
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase), iDbIndex) & " \\ " 'полное наименование
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 3, iDbIndex) 'методика поверки
        
        If sArrDBase(LBound(sArrDBase) + 5, iDbIndex) <> "nodata" Then _
            sTempStr = sTempStr & " \\ " & sArrDBase(UBound(sArrDBase) - 1, iDbIndex) 'текущий СИ является эталоном
    End If

    GetListStringForInstruments = sTempStr
End Function
'#########################################################
'получить строку для базы данных средств измерений
Function GetListStringOtherDB(sArrDBase() As String, iDbIndex As Integer) As String
    
    Dim sTempStr As String, j As Byte, bolStopReplaceJ As Boolean
    
    If UMenu.typе = etalonsOLD And sArrDBase(LBound(sArrDBase), iDbIndex) <> "" Then _
        sTempStr = sTempStr & sArrDBase(UBound(sArrDBase), iDbIndex) & " \\ " 'ключевое слово в наименовании эталона
                
    For j = LBound(sArrDBase) To UBound(sArrDBase) 'передать значения из подстрок
        If UMenu.typе = etalonsOLD And j = UBound(sArrDBase) Then Exit For 'не вносить ключевое слово эталонов повторно
        
        If UMenu.typе = organisationsOLD And j = 2 Then j = 3: bolStopReplaceJ = True 'ключевое слово
        If bolStopReplaceJ = False And UMenu.typе = organisationsOLD And j = 3 Then j = 2: bolStopReplaceJ = True 'адрес
            
            If j > LBound(sArrDBase) And sArrDBase(j, iDbIndex) <> DEFAULT_VALUE And sArrDBase(j, iDbIndex) <> "" Then
                sTempStr = sTempStr & " \\ " 'разделитель
                If UMenu.typе = organisationsOLD Then If j = 1 Then sTempStr = sTempStr & "ИНН " 'данные ИНН
            End If
            
            If sArrDBase(j, iDbIndex) <> DEFAULT_VALUE And sArrDBase(j, iDbIndex) <> "" Then _
                sTempStr = sTempStr & sArrDBase(j, iDbIndex) 'переменная для добавления в листбокс
            
        If bolStopReplaceJ And UMenu.typе = organisationsOLD And j = 3 Then j = 2: bolStopReplaceJ = False 'ключевое слово
        If bolStopReplaceJ And UMenu.typе = organisationsOLD And j = 2 Then j = 3 'адрес
        
    Next j

    GetListStringOtherDB = sTempStr
End Function
'#########################################################
'функция вставляет пробелы рядом с индексом элемента в листбоксе
Function TrueSpace(number As Integer)
    Select Case Len(CStr(number))
        Case 1 'числа от 1 до 10
            TrueSpace = number & "     "
        Case 2 'числа свыше 10
            TrueSpace = number & "   "
        Case 3
            TrueSpace = number & " "
    End Select
End Function
'#########################################################
'результат поиска в базе данных
Private Sub listDataBase_Click()

    Dim myData As New MSForms.DataObject
    Me.TextBox2 = "": Me.TextBox3 = "": Me.TextBox4 = "": Me.TextBox5 = ""
    
    ' 0 - наименование Организации / СИ / эталона / Фамилия
    ' 1 - ИНН / модификация СИ / тип эталона
    ' 2 - ключевое слово для поиска / номер в ФИФ СИ / номер в ФИФ эталона / Должность
    ' 3 - адрес организации / методика поверки СИ / доп. сведения эталона
    Me.TextBox1 = sArrDataBase(0, listDataBase.ListIndex)

    If UMenu.typе <> personsOLD Then
        If sArrDataBase(1, listDataBase.ListIndex) <> "nodata" Then _
            myData.SetText sArrDataBase(1, listDataBase.ListIndex): Me.TextBox2.text = CheckDataBeforepaste(myData, Me.TextBox2.name, 0, 0)
    
        If sArrDataBase(2, listDataBase.ListIndex) <> "nodata" Then _
            myData.SetText sArrDataBase(2, listDataBase.ListIndex): Me.TextBox3.text = CheckDataBeforepaste(myData, Me.TextBox3.name, 0, 0)

        If sArrDataBase(3, listDataBase.ListIndex) <> "nodata" Then _
            myData.SetText sArrDataBase(3, listDataBase.ListIndex): Me.TextBox4.text = CheckDataBeforepaste(myData, Me.TextBox4.name, 0, 0)
        
        If UMenu.typе = instrumentsOLD Then 'средства измерений
    
            If Me.cmbImport.caption <> "Выбрать" Then 'выбор перекрёстной ссылки
                Me.chbVerRefer = False
            
                Dim sReferenceStr As String 'перекрёстная ссылка
                sReferenceStr = "nodata"
                
                If listDataBase.ListIndex >= 0 Then _
                    sReferenceStr = sArrDataBase(LBound(sArrDataBase) + 4, listDataBase.ListIndex)
                    
                If sReferenceStr <> "nodata" Then
                    Me.cmbImport.caption = sReferenceStr
                    Me.chbVerRefer = True
                End If
                
            End If
            
            Me.chbEtalon = False:  If listDataBase.ListIndex >= 0 Then _
                If sArrDataBase(LBound(sArrDataBase) + 5, listDataBase.ListIndex) <> "nodata" Then Me.chbEtalon = True
            
        End If

        If UMenu.typе = etalonsOLD Then 'эталоны
            If sArrDataBase(UBound(sArrDataBase), listDataBase.ListIndex) <> "nodata" Then _
                myData.SetText sArrDataBase(UBound(sArrDataBase), listDataBase.ListIndex): Me.TextBox5.text = CheckDataBeforepaste(myData, Me.TextBox5.name, 0, 0)
        End If
        
    Else 'ввод сведений о фамилиях и должностях
        If sArrDataBase(1, listDataBase.ListIndex) <> "nodata" Then _
            myData.SetText sArrDataBase(1, listDataBase.ListIndex): Me.TextBox3.text = CheckDataBeforepaste(myData, Me.TextBox3.name, 0, 0)
        
        Me.chbVerRefer = False: If sArrDataBase(2, listDataBase.ListIndex) <> "nodata" Then Me.chbVerRefer = True
    End If
    
    With Me.cmbAdd
        If Me.cmbImport.caption = "Импорт" Then .caption = "Добавить": .Enabled = True: .BackColor = &HFFFFFF 'белый цвет
    End With
    
    Set myData = Nothing: RefreshLabelInfo
    If Me.cmbImport.caption = "Импорт" Then PreSaveSetButton True
End Sub
Private Sub listDataBase_Change()

    Me.cmbDelete.Enabled = False 'кнопка "удалить"
    
    If UMenu.typе = instrumentsOLD Then _
        Me.cmbFillTempProp.Enabled = False: Me.cmbOpenTemplateFolder.Enabled = False  'кнопка сформировать шаблон

    If Me.listDataBase.ListIndex >= 0 Then
        Me.cmbDelete.Enabled = True
        
        If UMenu.typе = instrumentsOLD And myWdDoc = False Then
            Me.cmbFillTempProp.Enabled = True
            
            Dim sTempStr As String
            If listDataBase.ListIndex >= 0 Then sTempStr = sArrDataBase(LBound(sArrDataBase) + 2, listDataBase.ListIndex)
        ' ------------------------------------------------------
'todo: отвязать workclsm
            If Dir(WorkClsm.templatesDir & "*" & sTempStr & "*", vbDirectory) <> "" Then Me.cmbOpenTemplateFolder.Enabled = True
        End If
        
    End If
End Sub
'#########################################################
'процедура реагирует на ввод значений в текстбоксы
Sub PreSaveSetButton(Optional NoRecapture As Boolean)
' ------------------------------------------------------
'todo: отвязать workclsm
    If NoRecapture = False Then 'опция ввода дополнительных значений
        If WorkClsm.DbName <> "недоступно" Then _
            Me.cmbAdd.caption = "Внести": EnableAddButton
    End If

    With Me.cmbReady
        If UBound(listDataBase.List) >= 0 And bolUserInput And listDataBase.ListIndex >= 0 Then 'можно обновить текущую позицию
            .caption = "Обновить": .Font.Size = 12: .BackColor = &HC0E0FF   'коралловый
        Else
            If bolAlreadySaved = False Then 'измерения не были сохранены
                .caption = "Сохранить": .Font.Size = 11: .BackColor = &HC0FFFF 'желтый цвет
            Else 'изменения были сохранены
                .caption = "Готово": .Font.Size = 12: .BackColor = &HFFFFFF 'белый цвет
            End If
        End If
    End With
End Sub
'#########################################################
'наименование Организации / СИ / эталона
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = True
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = False
End Sub
Private Sub TextBox1_Change()

    With Me.TextBox1
        .BackColor = &HFFFFFF: If .text = "" Then .BackColor = &HC0FFFF 'жёлтый

        If bolUserInput = False Then Exit Sub
        .text = Replace(.text, Chr(45) & Chr(45), Chr(150)) 'заменить двойной дефис на тире
    End With

    PreSaveSetButton
End Sub
Private Sub TextBox1_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True

    With Me.Controls(ActiveControl.name): _
        .text = CheckDataBeforepaste(data, ActiveControl.name, .SelStart, .SelLength): End With
End Sub
'#########################################################
' ИНН / модификация СИ / тип эталона / И.О.Фамилия
Private Sub TextBox2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = False
End Sub
Private Sub TextBox2_Change()
    With Me.TextBox2
        .BackColor = &HFFFFFF 'белый
        
        Select Case UMenu.typе
        
            Case organisationsOLD
                If Len(.text) < .maxLength Then .BackColor = &HC0FFFF   'жёлтый
                
            Case instrumentsOLD, etalonsOLD
                If .text = "" Then .BackColor = &HC0FFFF 'жёлтый
        End Select
        
        If bolUserInput = False Then Exit Sub
        .text = Replace(.text, Chr(45) & Chr(45), Chr(150)) 'заменить двойной дефис на тире
    End With
    
    PreSaveSetButton
End Sub
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = True
    Select Case KeyCode
        Case 8, 27, 35, 36, 37, 39 'эти нажатия контролируются в классе oTxtBx
        Case Else
            If UMenu.typе = organisationsOLD Then 'только для БД заказчиков
                With Me.TextBox2
                    Select Case Len(.text)
                        Case 3, 7, 10
                            .text = .text & " " 'вставить доп.пробелы в ИНН для удобочитаемости
                    End Select
                End With
            End If
    End Select
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger) 'макрос проверяет ввод значений по нажатию клавиш в текстбоксе
    If UMenu.typе = organisationsOLD Then
        'если вводится НЕ символы "0123456789", не допустить ввод этого символа
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    End If
End Sub
Private Sub TextBox2_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    With Me.Controls(ActiveControl.name)
        .text = CheckDataBeforepaste(data, _
            ActiveControl.name, .SelStart, .SelLength)
    End With
End Sub
'#########################################################
' сокращённое наименование для архива / номер в ФИФ СИ / номер в ФИФ эталона
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = True
End Sub
Private Sub TextBox3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = False
End Sub
Private Sub TextBox3_Enter()
    If UMenu.typе = etalonsOLD Then _
        RefreshLabelInfo "Режим ввода прописных букв.", &H800080
End Sub
Private Sub TextBox3_Change()
    With Me.TextBox3
    
        .BackColor = &HFFFFFF 'белый
        
        Select Case UMenu.typе
            
            Case instrumentsOLD  'средства измерений
                If bolUserInput And Len(.text) < .maxLength Then .BackColor = &HC0FFFF  'жёлтый
                
            Case organisationsOLD, etalonsOLD, personsOLD  ' заказчики, эталоны, фамилии и должности,
                If .text = "" Then .BackColor = &HC0FFFF 'жёлтый
                
        End Select
    End With
    
    If bolUserInput = False Then Exit Sub
    PreSaveSetButton
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case UMenu.typе
        Case instrumentsOLD  'средства измерений - ввод номера в ФИФ (метки каталога)
        
            Select Case KeyAscii
                Case 34, 42, 47, 58, 60, 62, 63, 92, 124: KeyAscii = 0 ' \/:*?"<>| запрещённые в наименовании файла символы
            End Select
            
        Case etalonsOLD  ',1 'эталоны - ввод только заглавных символов
            
            Select Case KeyAscii
                Case 34, 42, 47, 58, 60, 62, 63, 92, 124: KeyAscii = 0: Exit Sub ' \/:*?"<>| запрещённые в наименовании файла символы
            End Select
            
            Dim iCode As Integer, bArrTemp(1) As Byte, sSym As String
            iCode = KeyAscii: KeyAscii = 0
            '##################################################################### 'хз, как это работает
            bArrTemp(0) = iCode Mod 256: bArrTemp(1) = iCode / 256: sSym = bArrTemp
            '##################################################################### 'но работает
            With Me.TextBox3
            
                If Len(.text) < .maxLength Then
                    Dim sLeftStr As String, sRightStr As String, iselSt As Integer
                    
                    iselSt = .SelStart: sLeftStr = Left(.text, .SelStart) 'левая часть текста
                    sRightStr = Right(.text, Len(.text) - (.SelStart + .SelLength)) 'правая часть текста
                    .text = sLeftStr & UCase(sSym) & sRightStr: .SelStart = iselSt + 1
                End If
            End With
    End Select

End Sub
Private Sub TextBox3_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    
    With Me.Controls(ActiveControl.name)
        .text = CheckDataBeforepaste(data, ActiveControl.name, .SelStart, .SelLength)
    End With
End Sub
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UMenu.typе = etalonsOLD Then RefreshLabelInfo
End Sub
'#########################################################
' адрес организации / методика поверки СИ / доп. сведения эталона
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = True
End Sub
Private Sub TextBox4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = False
End Sub
Private Sub TextBox4_Change()
    
    With Me.TextBox4
        .BackColor = &HFFFFFF 'белый
        
        Select Case UMenu.typе
            
            Case organisationsOLD, instrumentsOLD, etalonsOLD
                If .text = "" Then .BackColor = &HC0FFFF 'жёлтый
                
        End Select
    End With
    
    If bolUserInput = False Then Exit Sub
    PreSaveSetButton
End Sub
Private Sub TextBox4_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    With Me.Controls(ActiveControl.name)
        .text = CheckDataBeforepaste(data, _
            ActiveControl.name, .SelStart, .SelLength)
    End With
End Sub
'#########################################################
' ключевое слово
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = True
End Sub
Private Sub TextBox5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    bolUserInput = False
End Sub
Private Sub TextBox5_Change()
    With Me.TextBox5
        .BackColor = &HFFFFFF 'белый
        
        Select Case UMenu.typе
            
            Case etalonsOLD
                If .text = "" Then .BackColor = &HC0FFFF 'жёлтый
                
                If bolUserInput = False Then Exit Sub
                .text = Replace(.text, Chr(45) & Chr(45), Chr(150)) 'заменить двойной дефис на тире
                
        End Select
    End With
    
    PreSaveSetButton
End Sub
'#########################################################
'функция возвращает скорректированное значение буфера обмена в текущее поле ввода
Function CheckDataBeforepaste(ByVal data As MSForms.DataObject, _
    controlName As String, SelStart As Integer, SelLength As Integer) As String
    
    Dim sBaseStr As String, sLeftStr As String, sRightStr As String
    sBaseStr = Me.Controls(controlName).text 'текущее содержимое поля ввода
    sLeftStr = Left(sBaseStr, SelStart): If SelStart <= Len(sBaseStr) Then sRightStr = Right(sBaseStr, Len(sBaseStr) - SelStart - SelLength)
    
    Dim sTempStr As String
    sTempStr = data.GetText 'получить значение из буфера обмена
    sTempStr = DeleteSpaceStEnd(sTempStr) 'убрать нечитаемые символы в начале и конце
    ReplaceInvCommas sTempStr 'правильно заменить кавычки
    
    Select Case UMenu.typе
        
        Case organisationsOLD 'заказчики
            CheckCustomersData controlName, sTempStr, sBaseStr
            
        Case instrumentsOLD 'средства измерений
            CheckInstrumentsData controlName, sTempStr, sBaseStr
            
        Case etalonsOLD  'эталоны
            CheckEtalonsData controlName, sTempStr
    End Select
    
    sBaseStr = sLeftStr & sTempStr & sRightStr 'новое значение для передачи
    
    InsertSpaceToBaseStr controlName, sBaseStr

    If Me.Controls(controlName).maxLength > 0 Then 'если у текущего поля есть ограничения по вводу символов
        If Len(sBaseStr) > Me.Controls(controlName).maxLength Then _
            sBaseStr = Left(sBaseStr, Me.Controls(controlName).maxLength)
    End If
        
    CheckDataBeforepaste = sBaseStr
End Function

'#########################################################
'частная процедура для конструктора заказчиков
Sub CheckCustomersData(sCtrlName As String, ByRef sTempStr As String, ByRef sBaseStr As String)
    
    If sCtrlName = "TextBox2" Then  'только при вводе ИНН
        sTempStr = Replace(LCase(sTempStr), "и", ""): sTempStr = Replace(LCase(sTempStr), "н", "") 'убрать инн
        sBaseStr = Replace(sBaseStr, " ", "") 'убрать пробелы в исходной строке
        sTempStr = Replace(sTempStr, " ", ""): sTempStr = Replace(sTempStr, ":", "") 'убрать символы в добавляемой строке
        
        If IsNumeric(sTempStr) = False Or Len(sBaseStr & sTempStr) > 10 Then _
            sTempStr = "": RefreshLabelInfo "Ввод некорректных данных", &HFF& 'красный цвет
    End If
End Sub
'#########################################################
'частная процедура для конструктора средств измерений
Sub CheckInstrumentsData(sCtrlName As String, ByRef sTempStr As String, ByRef sBaseStr As String)
     'только при вводе номера в ФИФ СИ
    If sCtrlName = "TextBox3" Then _
        If Len(sBaseStr & sTempStr) > 8 Then sTempStr = "": RefreshLabelInfo "Ввод некорректных данных", &HFF& 'красный цвет
End Sub
'#########################################################
'частная процедура для конструктора эталонов
Sub CheckEtalonsData(sCtrlName As String, ByRef sTempStr As String)
                
    If sCtrlName = "TextBox3" Then  'только при вводе номера в ФИФ эталона
        sTempStr = Replace(sTempStr, " ", "")
        
        If Len(sTempStr) > Me.Controls(sCtrlName).maxLength Then _
            sTempStr = "": RefreshLabelInfo "Ввод некорректных данных", &HFF& 'красный цвет
    End If
End Sub
'#########################################################
'частная процедура для вставки пробелов в строку
Sub InsertSpaceToBaseStr(sCtrlName As String, ByRef sBaseStr As String)
                
    Dim i As Integer
    If UMenu.typе = organisationsOLD And sCtrlName = "TextBox2" Then
        i = Len(sBaseStr) 'общее количество символов в строке
        
        Do 'вставка дополнительных пробелов
            If i < 3 Then Exit Do
            Select Case i
                Case 3, 6, 8
                    sBaseStr = Left(sBaseStr, i) & " " & Right(sBaseStr, Len(sBaseStr) - i)
            End Select
            i = i - 1
        Loop
    End If
End Sub
'#########################################################
'процедура отвечает за поведение кнопки "Добавить"
Sub EnableAddButton()
    Dim bolStopEnable As Boolean, sTempStr As String

    With Me.cmbAdd  'кнопка "добавить"
        If .caption = "Внести" Then .Enabled = False: .BackColor = &HFFFFFF 'белый цвет

        If Me.TextBox1 = "" Then bolStopEnable = True 'полное наименование
        sTempStr = "Заполните обязательные поля"
        
        Select Case UMenu.typе
        
            Case organisationsOLD '1 - заказчики
                If Len(Me.TextBox2.text) < Me.TextBox2.maxLength Or _
                    Len(Me.TextBox4.text) = 0 Then bolStopEnable = True: 'ИНН / адрес
                    
                If Not IsItemUnique(Me.TextBox3) Then _
                    bolStopEnable = True: sTempStr = "Для внесения измените сокращение."
                    
            Case instrumentsOLD '2 - средства измерений
                If Len(Me.TextBox3.text) < Me.TextBox3.maxLength - 1 Or _
                   Me.TextBox2 = "" Or Me.TextBox4 = "" Then bolStopEnable = True 'номер в ФИФ пустой / тип СИ / МП
                   
                If AlreadyInBase(Me.TextBox3) Then _
                    bolStopEnable = True: sTempStr = "Для внесения измените номер ФИФ."
            Case etalonsOLD 'эталоны
            
                If Me.TextBox2 = "" Or Me.TextBox3 = "" Or _
                    Me.TextBox4 = "" Or Me.TextBox5 = "" Then bolStopEnable = True 'тип / зав № / примечание / ключевая фраза
                    
                If AlreadyInBase(Me.TextBox3) Then _
                    bolStopEnable = True: sTempStr = "Для внесения измените номер."
                    
            Case personsOLD 'фамилии
                If Me.TextBox3 = "" Then bolStopEnable = True 'должность
                
                If AlreadyInBase(Me.TextBox1) Then _
                    bolStopEnable = True: sTempStr = "Сотрудник уже присутствует в базе."
        End Select
        
        If bolStopEnable = False Then .Enabled = True: .BackColor = &HC0FFC0 'зелёный
        If bolStopEnable Then RefreshLabelInfo sTempStr, &H800080     'фиолетовый цвет '.BackColor = &HC0FFFF
    End With
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.02.2023 17:20
    ' Назначение:
    '    параметр key:
    ' Возвращаемый тип: Boolean
    ' ----------------------------------------------------------------
    Private Function IsItemUnique( _
        key As String _
        ) As Boolean
        
        IsItemUnique = False
        
        If key = vbNullString Then _
            Exit Function
        
        Dim i As Integer
        For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
        
            If sArrDataBase(2, i) = key Then _
                Exit Function 'поиск ключевого слова организаций
                
        Next
        
        IsItemUnique = True
    End Function
'#########################################################
'проверка наличия в массиве
Function AlreadyInBase(sFindStr As String) As Boolean
    
    Dim i As Integer

    With listDataBase
        
        For i = LBound(.List) To UBound(.List)
            If InStr(.List(i), sFindStr) > 0 Then _
                AlreadyInBase = True: Exit Function 'если номер в фиф уже есть в базе
        Next
        
    End With
End Function


'#########################################################
'импорт данных из существующей базы данных
Private Sub cmbImport_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then bolUserInput = False
End Sub
Private Sub cmbImport_Click()
        
    

    Select Case cmbImport.caption

        Case "Импорт"
        
            MsgBox "Функционал недоступен, сорямба :("
'            Dim sJoinDBName As String, iJoinCnt As Integer
'            sJoinDBName = GetFileFPath(myMenu.ty pe * 2, WorkClsm.startDir, "Присоединить к базе данных " & WorkClsm.DbName)
'
'            If sJoinDBName <> "NoPath" Then 'если был выбран путь к файлу
'                iJoinCnt = JoinDataBase(sJoinDBName, sArrDataBase) 'импортировать БД и получить количество импортированных файлов
'
'                Select Case iJoinCnt 'количество импортированных значений
'                    Case 0: RefreshLabelInfo "Нет сведений для импорта", &H800080    'фиолетовый цвет
'                    Case Else
'                        Me.listDataBase.Selected(0) = True: RefreshLabelInfo RusPadej3(iJoinCnt), &H8000& 'зелёный цвет
'                        bolAlreadySaved = False: PreSaveSetButton True 'Вызвать кнопку "Сохранить"
'                End Select
'            End If
        Case "Выбрать" 'Выбрать перекрёстную ссылку

            If iDataBaseIndex = listDataBase.ListIndex Then
                RefreshLabelInfo "Выберите другой элемент.", &H800080
            Else

                Dim referenceFifNum As String, _
                    currentFifNum As String

                referenceFifNum = Me.TextBox3
                currentFifNum = sArrDataBase(LBound(sArrDataBase) + 2, iDataBaseIndex)

                sArrDataBase(LBound(sArrDataBase) + 4, iDataBaseIndex) = referenceFifNum 'передать ссылку на объект
                Me.cmbImport.caption = referenceFifNum
                
                ' ------------------------------------------------------
                Dim targetDir As String
                targetDir = Dir(Config.instrumentsPath & "*" & currentFifNum & "*", vbDirectory) 'каталог СИ
                
                Dim fso As New FileSystemObject
                
                Dim targetPath As String
                targetPath = fso.BuildPath(Config.instrumentsPath, targetDir)
                
                Dim targetFilePath As String
                targetFilePath = fso.BuildPath(targetPath, REFERENCE_FIFNUM_FILENAME)
                
                Base.WriteContent targetFilePath, referenceFifNum
                ' ------------------------------------------------------

                bolAlreadySaved = False
                PreSaveSetButton True 'Вызвать кнопку "Сохранить"
            End If

        Case Else 'перейти к номеру в фиф
            Dim sTempStr As String, i As Integer
            sTempStr = Me.cmbImport.caption

            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                If sArrDataBase(2, i) = sTempStr Then Me.listDataBase.Selected(i) = True: Exit For
            Next i
    End Select
End Sub
'#########################################################
'функция импортирует массив базы данных и возвращает количество импортированных строк
Function JoinDataBase(JoinFPath As String, ByRef BaseDataArr() As String) As Integer
    Dim sArrFile() As String, sInstrDel As String, bolOldFileFormat As Boolean, bolInstrFound As Boolean
    
    sArrFile = myBase.GetArrFF(JoinFPath)    'загрузить во временный массив данные добавляемого файла
    If sArrFile(LBound(sArrFile)) = "newFile" Then Exit Function 'свежесозданный файл JoinDataBase = 0
    
    sInstrDel = "%" 'поиск данных по разделителю старого типа
    If InStr(sArrFile(LBound(sArrFile)), sInstrDel) > 0 Then bolInstrFound = True: bolOldFileFormat = True
    
    If bolInstrFound = False Then sInstrDel = InStrDelimiter: _
        If InStr(sArrFile(LBound(sArrFile)), sInstrDel) > 0 Then bolInstrFound = True 'поиск данных нового типа, если не опознан старый тип
    
    If bolInstrFound Then 'если было опознано наличие хоть одного разделителя в импортируемом файле
        JoinDataBase = SubJoinDB(BaseDataArr, sArrFile, sInstrDel, bolOldFileFormat) 'частная процедура импорта
        SortMassBiv BaseDataArr(): UpdateListDataBase BaseDataArr  'отсортировать двумерный массив по тексту по убыванию, начиная с нулевого элемента
    End If
End Function
'#########################################################
'субфункция импорта БД
Function SubJoinDB(ByRef BaseDataArr() As String, _
    sArrFile() As String, sInstrDel As String, Optional OldFileFormat As Boolean) As Integer

    Dim sArrTemp() As String, i As Integer, j As Integer, bolStopLineImport As Boolean
    For i = LBound(sArrFile) To UBound(sArrFile) 'пройтись по всему массиву
    
        bolStopLineImport = False: sArrTemp = Split(sArrFile(i), sInstrDel) 'разбить строку на специальных пробелах
        
        If UBound(BaseDataArr) = 0 Then _
            ReDim BaseDataArr(UBound(sArrTemp), 0): If UMenu.typе = instrumentsOLD Then ReDim BaseDataArr(UBound(sArrTemp) + 1, 0)

        For j = LBound(BaseDataArr, 2) To UBound(BaseDataArr, 2) 'пройтись по второму измерению массива
            
            If UMenu.typе = organisationsOLD Then 'заказчики
            
                If OldFileFormat Then 'для старого формата файлов заказчиков
                    If sArrTemp(LBound(sArrTemp) + 1) = BaseDataArr(LBound(BaseDataArr), j) Then _
                        bolStopLineImport = True: Exit For 'если наименование уже есть в массиве - не добавлять элемент
                Else 'современный формат файла
                    If sArrTemp(LBound(sArrTemp)) = BaseDataArr(LBound(BaseDataArr), j) Then _
                        bolStopLineImport = True: Exit For 'если наименование уже есть в массиве - не добавлять элемент
                End If
                
            ElseIf UMenu.typе = instrumentsOLD Then 'база данных средств измерений - проверка по номеру в ФИФ
                If sArrTemp(2) = BaseDataArr(2, j) Then bolStopLineImport = True: Exit For
                
            Else 'современный формат файла
                If sArrTemp(LBound(sArrTemp)) = BaseDataArr(LBound(BaseDataArr), j) Then _
                    bolStopLineImport = True: Exit For 'если наименование уже есть в массиве - не добавлять элемент
            End If
        Next j
        
        If bolStopLineImport = False Then
            If BaseDataArr(LBound(BaseDataArr), UBound(BaseDataArr, 2)) <> "" Then _
                ReDim Preserve BaseDataArr(UBound(BaseDataArr), UBound(BaseDataArr, 2) + 1) 'расширить массив, если он полный
            
            If OldFileFormat And UMenu.typе = organisationsOLD Then 'для старого формата файлов заказчиков
                BaseDataArr(0, UBound(BaseDataArr, 2)) = sArrTemp(LBound(sArrTemp) + 1) 'наименование
                BaseDataArr(1, UBound(BaseDataArr, 2)) = sArrTemp(LBound(sArrTemp) + 2) 'ИНН
                BaseDataArr(2, UBound(BaseDataArr, 2)) = sArrTemp(LBound(sArrTemp) + 3) 'сокращение
            Else 'современный формат файла
                
                For j = LBound(sArrTemp) To UBound(sArrTemp) 'передать значения из временного массива
                    BaseDataArr(j, UBound(BaseDataArr, 2)) = sArrTemp(j) 'заполнить массив сведенями базы данных
                Next j
                
                If UMenu.typе = instrumentsOLD Then BaseDataArr(UBound(BaseDataArr), UBound(BaseDataArr, 2)) = "nodata" 'сведения о наличии шаблонов
            End If
            
            SubJoinDB = SubJoinDB + 1
        End If
    Next i
End Function
'#########################################################
'кнопка "удалить" выбранную строку из базы данных
Private Sub cmbDelete_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then bolUserInput = False
End Sub
Private Sub cmbDelete_Click()
    
    Dim ask As Integer
    ask = Handler.ask("Удалить элемент из базы?")
    
    If ask = vbYes Then _
        cmbDeleteClck
        
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 15:53
' Назначение:
'    параметр bolUpdateString:
' ----------------------------------------------------------------
Private Sub cmbDeleteClck( _
    Optional bolUpdateString As Boolean = False _
    )
    With Me.listDataBase
    
        If .ListIndex >= 0 Then 'если выбрана хотя бы первая строка
        
            Dim iLiInd As Integer
            iLiInd = .ListIndex
            ' ----------------------------------------------------------------
            If bolUpdateString Then 'только обновить строку
           
                UpdateArrDB sArrDataBase, .ListIndex
                bolAlreadySaved = True
                ' ------------------------------------------------------
'todo: отвязать workclsm
                ArrDBaseToFile sArrDataBase, WorkClsm.startDir & WorkClsm.DbName  'изменения были сохранены
                PreSaveSetButton True  'Вызвать кнопку "Сохранить"
                
                Exit Sub
            End If
            ' ----------------------------------------------------------------
            ReduceArrDB sArrDataBase, .ListIndex 'удалить текущий элемент
            UpdateListDataBase sArrDataBase 'заполнить листбокс данным массива организаций
            
            If .ListCount > 0 Then 'если в листбоксе остался хоть один элемент
                If iLiInd <= UBound(.List()) Then .Selected(iLiInd) = True 'оставить выделение на текущем элементе
                If iLiInd > UBound(.List()) Then .Selected(UBound(.List())) = True 'выделить последний доступный элемент
            Else 'в листбоксе не осталось элементов
                With Me.cmbAdd: .caption = "Внести": .Enabled = False: End With
                Me.cmbDelete.Enabled = False: ClearTxtBxes
            End If
            
            bolAlreadySaved = False
            PreSaveSetButton True  'Вызвать кнопку "Сохранить"
            
        Else
            RefreshLabelInfo "Сперва выберите элемент."
        End If
        
    End With
    TextBox1.SetFocus
End Sub
'#########################################################
'процедура уменьшает текущий массив на один элемент
Sub ReduceArrDB(ByRef BaseArrData() As String, RIndex As Integer)

    Dim j As Integer
    For j = LBound(BaseArrData) To UBound(BaseArrData)
        BaseArrData(j, RIndex) = "" 'очистить все элементы текущей строки
    Next j
    
    SortMassBiv BaseArrData 'отсортировать массив по умолчанию
    If UBound(BaseArrData, 2) > 0 Then _
        ReDim Preserve BaseArrData(UBound(BaseArrData), UBound(BaseArrData, 2) - 1) 'уменьшить массив на один элемент с сохранением данных
End Sub
























' ----------------------------------------------------------------
' Дата: 25.02.2023 15:55
' Назначение: процедура обновляет значения текущего массива
'    параметр BaseArrData:
'    параметр RIndex:
' ----------------------------------------------------------------
Private Sub UpdateArrDB( _
    ByRef BaseArrData() As String, _
    RIndex As Integer _
    )

    Dim sSelSaved As String 'заполнить элемент для выделения его в списке впоследствии
    Select Case UMenu.typе
    
        Case organisationsOLD, personsOLD
            sSelSaved = Me.TextBox1 'наименование
            
        Case instrumentsOLD, etalonsOLD
            sSelSaved = Me.TextBox3 'номер в фиф / заводской номер
            
    End Select

    Dim sTb1Text As String, _
        sTb2Text As String, _
        sTb3Text As String, _
        sTb4Text As String, _
        sTb5Text As String
        
    sTb1Text = Me.TextBox1
    sTb2Text = Me.TextBox2
    sTb3Text = Me.TextBox3
    sTb4Text = Me.TextBox4
    sTb5Text = Me.TextBox5
    
    If UMenu.typе = organisationsOLD Then _
        sTb2Text = Replace(sTb2Text, " ", "")  'если вводится ИНН
    
    DeleteSpaceStEnd sTb1Text, True: DeleteSpaceStEnd sTb2Text, True: DeleteSpaceStEnd sTb3Text, True
    DeleteSpaceStEnd sTb4Text, True: DeleteSpaceStEnd sTb5Text, True
    
    ' ----------------------------------------------------------------
    If UMenu.typе = organisationsOLD Then
    
        Dim key As String
        key = BaseArrData(2, RIndex)
        
        Dim newKey As String
        newKey = sTb3Text
         
        Dim item As New cItemOrganisation
        item.shortName = sTb1Text
        item.taxNumber = sTb2Text
        item.legalAddress = sTb4Text

        DataBase.AddItem _
            key:=key, _
            itemData:=item

    End If

    ' ----------------------------------------------------------------
    Select Case UMenu.typе
    
        Case organisationsOLD, instrumentsOLD, etalonsOLD
            BaseArrData(0, RIndex) = sTb1Text 'наименование
            BaseArrData(1, RIndex) = sTb2Text 'инн / тип
            BaseArrData(2, RIndex) = sTb3Text 'КЛЮЧ ОРГАНИЗАЦИЙ / рег. № фиф
            BaseArrData(3, RIndex) = sTb4Text 'адрес / МП
            
        Case personsOLD 'фамилии и должности
            BaseArrData(0, RIndex) = sTb1Text: BaseArrData(1, RIndex) = sTb3Text 'должность
            BaseArrData(2, RIndex) = "nodata": If Me.chbVerRefer Then BaseArrData(2, RIndex) = "поверитель"
            
    End Select
    ' ----------------------------------------------------------------
    If UMenu.typе = etalonsOLD Then _
        BaseArrData(4, RIndex) = sTb5Text
    
    Select Case UMenu.typе
        
        Case instrumentsOLD
            SortMassBiv BaseArrData, 1 'отсортировать массив по типу СИ
            
        Case etalonsOLD
            SortMassBiv BaseArrData, UBound(BaseArrData) 'отсортировать массив по ключевому слову эталона
        
        Case Else
            SortMassBiv BaseArrData 'отсортировать массив по умолчанию 'заказчики, фамилии - сортировать по полному наименованию
            
    End Select
    
  
    UpdateListDataBase sArrDataBase 'заполнить листбокс данным массива организаций
    
    Dim i As Integer
    For i = LBound(listDataBase.List) To UBound(listDataBase.List)
        If InStr(listDataBase.List(i), sSelSaved) > 0 Then _
            listDataBase.Selected(i) = True: Exit For 'выделить добавленную строку
    Next i
    
    bolAlreadySaved = False
    PreSaveSetButton True 'Вызвать кнопку "Сохранить"
End Sub





























'#########################################################
'кнопка "добавить"
Private Sub cmbAdd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then bolUserInput = False
End Sub
Private Sub cmbAdd_Click()
    
    With Me.cmbAdd

        If .caption = "Добавить" Then 'очистить поля для внесения данных в базу
            
            Me.TextBox1.SetFocus
            If UMenu.typе = instrumentsOLD Then _
                Me.TextBox3.SetFocus

            .caption = "Внести"
            .Enabled = False
            
            ClearTxtBxes 'bolUserInput = True: ClearTxtBxes: bolUserInput = False
            
            Me.chbVerRefer = False
            Me.chbEtalon = False
            PreSaveSetButton
        Else 'внести сведения в базу данных

            Dim sTempSelect As String, _
                sTb1Text As String, _
                sTb2Text As String, _
                sTb3Text As String, _
                sTb4Text As String, _
                sTb5Text As String
                
            Select Case UMenu.typе 'запомнить строку для повторного выделения
                
                Case organisationsOLD, personsOLD
                    sTempSelect = Me.TextBox1 'полное наименование или фамилия
                    
                Case instrumentsOLD, etalonsOLD
                    sTempSelect = Me.TextBox3 'номер фиф или же сокращение
                    
            End Select
            
            sTb1Text = Trim(Me.TextBox1)
            sTb2Text = Trim(Me.TextBox2)
            sTb3Text = Trim(Me.TextBox3)
            sTb4Text = Trim(Me.TextBox4)
            sTb5Text = Trim(Me.TextBox5)
            ' ------------------------------------------------------
            If UMenu.typе = organisationsOLD Then _
                sTb2Text = Replace(sTb2Text, " ", "")  'если вводится ИНН
            ' ------------------------------------------------------
            With myBase
                .ClearParameters
                .AddP DeleteSpaceStEnd(sTb1Text, True) 'полное наименование
                
                Select Case UMenu.typе
                
                    Case organisationsOLD 'организации
                        .AddP DeleteSpaceStEnd(sTb2Text, True) 'ИНН
                        
                        If sTb3Text = "" Then _
                            sTb3Text = "БЕЗИМЕНИ" ' Сокращение
                            
                        .AddP DeleteSpaceStEnd(sTb3Text, True) 'ключевое слово — идентификатор
                        .AddP DeleteSpaceStEnd(sTb4Text, True) 'адрес
                        
                    Case instrumentsOLD 'средства измерений
                        .AddP DeleteSpaceStEnd(sTb2Text, True)
                        .AddP DeleteSpaceStEnd(sTb3Text, True)
                        .AddP DeleteSpaceStEnd(sTb4Text, True)
                        .AddP DeleteSpaceStEnd("", True) 'наличие шаблона протокола и свидетельства - '- по умолчанию отсутствует.
                        
                    Case etalonsOLD 'эталоны
                        .AddP DeleteSpaceStEnd(sTb2Text, True)
                        .AddP DeleteSpaceStEnd(sTb3Text, True)
                        .AddP DeleteSpaceStEnd(sTb4Text, True)
                        .AddP DeleteSpaceStEnd(sTb5Text, True) 'ключевое слово для поиска
                        
                    Case personsOLD 'фамилии
                        .AddP DeleteSpaceStEnd(sTb3Text, True) 'должность
                        
                        Dim sTempStr As String
                        If Me.chbVerRefer Then sTempStr = "поверитель"
                        
                        .AddP DeleteSpaceStEnd(sTempStr, True) 'статус поверителя
                End Select
                
                AddToArrDataBase sArrDataBase, .Parameters
            End With
            ' ----------------------------------------------------------------
            ' Дата: 24.02.2023 22:57
            ' Назначение: создание каталога нового элемента
            ' ----------------------------------------------------------------
            If UMenu.typе = organisationsOLD Then
                
                Dim item As New cItemOrganisation
                item.shortName = sTb1Text
                item.taxNumber = sTb2Text
                item.legalAddress = sTb4Text
                
                
                DataBase.AddItem _
                    key:=sTb3Text, _
                    itemData:=item
                                
            End If
             ' ------------------------------------------------------
            If UMenu.typе = instrumentsOLD Then
            
                Dim regFifNum As String, _
                    regTypeName As String, _
                    templateNewPath As String
                
                regFifNum = DeleteSpaceStEnd(sTb3Text)
                regTypeName = DeleteSpaceStEnd(sTb2Text)
            ' ------------------------------------------------------
'todo: отвязать workclsm
                templateNewPath = WorkClsm.templatesDir & regTypeName & "_" & regFifNum & Application.PathSeparator
                If FolderNotExist(templateNewPath) Then _
                    MkDir templateNewPath 'создать каталог для данного СИ
                
                Dim templateNewSomniumPath As String
                ' ------------------------------------------------------
'todo: отвязать workclsm
                templateNewSomniumPath = WorkClsm.templatesDir & regTypeName & "_" & regFifNum & Application.PathSeparator
                
                If FolderNotExist(templateNewSomniumPath) Then _
                    MkDir templateNewSomniumPath 'создать каталог для данного СИ
                    
       '         Dim referenceFifNum As String
        '        referenceFifNum = sArrDataBase(LBound(sArrDataBase) + 4, iDataBaseIndex)
                 ' ------------------------------------------------------
'todo: отвязать workclsm
                CheckTempDir WorkClsm.startDir, sArrDataBase
            End If
             ' ------------------------------------------------------
            UpdateListDataBase sArrDataBase 'заполнить листбокс данным массива

            Dim i As Integer
            For i = LBound(listDataBase.List) To UBound(listDataBase.List)
            
                If listDataBase.List(i) Like "*" & sTempSelect & "*" Then _
                    listDataBase.Selected(i) = True: _
                    Exit For 'выделить добавленную строку
                    
            Next i
            
            .caption = "Добавить"
            bolAlreadySaved = False
            PreSaveSetButton True 'Вызвать кнопку "Сохранить"
        End If
    End With
End Sub
' ----------------------------------------------------------------
' Дата: 24.02.2023 17:56
' Назначение: создание каталога для элемента базы данных
'    параметр itemName: ключевой слово элемента БД
' ----------------------------------------------------------------
    Private Function CreateItemDir( _
        itemName As String _
        ) As String
        
        Dim fso As New FileSystemObject, _
            sourceDataPath As String
  ' ------------------------------------------------------
'todo: отвязать workclsm
        sourceDataPath = fso.BuildPath(WorkClsm.startDir, sourceDataPath)
        
        If Not fso.FolderExists(sourceDataPath) Then _
            fso.CreateFolder sourceDataPath
            
        Dim targetPath As String
        targetPath = fso.BuildPath(sourceDataPath, itemName)
        
        If Not fso.FolderExists(targetPath) Then _
            fso.CreateFolder targetPath
        
        CreateItemDir = targetPath
        
    End Function
        Private Function sourceDataPath( _
            ) As String
            
            sourceDataPath = Base.defaultValue
            
            Select Case True
                    
                Case UMenu.typе = organisationsOLD
                    sourceDataPath = "organisations"
                    
                Case UMenu.typе = instrumentsOLD
                    sourceDataPath = "instruments"
                
                Case UMenu.typе = etalonsOLD
                    sourceDataPath = "etalons"
                    
                Case UMenu.typе = personsOLD
                    sourceDataPath = "persons"
                
            End Select
            
        End Function

'#########################################################
'процедура присоединяет к исходному массиву дополнительный массив
Sub AddToArrDataBase(ByRef BaseArrData() As String, InpArr() As String)

    Select Case UMenu.typе
        
        Case organisationsOLD
            
            If UBound(BaseArrData) = 0 Then _
                ReDim BaseArrData(3, 0)
            
        Case instrumentsOLD, etalonsOLD
        
            If UBound(BaseArrData) = 0 Then _
                ReDim BaseArrData(4, 0)
            
        Case personsOLD
            If UBound(BaseArrData) = 0 Then _
                ReDim BaseArrData(2, 0)
                
    End Select
    
    If BaseArrData(0, UBound(BaseArrData, 2)) <> "" Then _
        ReDim Preserve BaseArrData(UBound(BaseArrData), UBound(BaseArrData, 2) + 1) 'расширить массив, если он полный
    
    ReDim Preserve InpArr(UBound(BaseArrData))
    
    Dim i As Integer
    For i = LBound(InpArr) To UBound(InpArr)
        If InpArr(i) = "" Then InpArr(i) = "nodata" ' заменить пустые значения
        BaseArrData(i, UBound(BaseArrData, 2)) = InpArr(i) 'добавить сведения в массив базы данных
    Next i
    
    SortMassBiv BaseArrData() 'отсортировать двумерный массив по тексту по убыванию, начиная с нулевого элемента
End Sub
'############################################################
'кнопка "готово"/"сохранить"
Private Sub cmbReady_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then bolUserInput = False
End Sub
Private Sub cmbReady_Click()
    Select Case Me.cmbReady.caption
        Case "Обновить": cmbDeleteClck True
        Case "Сохранить"
        ' ------------------------------------------------------
'todo: отвязать workclsm
            bolAlreadySaved = True: ArrDBaseToFile sArrDataBase, WorkClsm.startDir & WorkClsm.DbName  'изменения были сохранены
            'bolAlreadySaved = True: ArrDBaseToFile sArrDataBase, WorkClsm.startDir & WorkClsm.DbName, True  ' добавить новый параметр
        Case Else
        
            VBA.Unload Me


    End Select
End Sub
'############################################################
'процедура добавляет элемент в базу данных
Sub ArrDBaseToFile( _
    BaseArr() As String, _
    fPath As String, _
    Optional myNewParameter As Boolean _
    )
    
    Dim i As Integer, j As Integer, sTempStr As String, bUpperBound As Byte
    bUpperBound = UBound(BaseArr): If UMenu.typе = instrumentsOLD Then bUpperBound = UBound(BaseArr) - 1 'индивидуально для бд СИ

    For i = LBound(BaseArr, 2) To UBound(BaseArr, 2)
        
        For j = LBound(BaseArr) To bUpperBound  'подстроки
            sTempStr = sTempStr & BaseArr(j, i): If j < bUpperBound Then sTempStr = sTempStr & InStrDelimiter 'разделитель внутри строки
        Next j
        
        '######################################
        If myNewParameter Then sTempStr = sTempStr & InStrDelimiter & "nodata" 'ТОЛЬКО ДЛЯ ДОБАВЛЕНИЯ НОВОГО ПАРАМЕТРА
        '######################################
        If i < UBound(BaseArr, 2) Then sTempStr = sTempStr & vbNewLine 'разделитель между строк
    Next i
    
    If fPath = "недоступно" Then
        MsgBox "Путь недоступен, получить новый никак. Вот говно, правда?"
'        Select Case uMenu.typе
'
'            Case organisations, instruments, etalons
'                fPath = GetSaveAsFname(myMe nu.type * 2, WorkClsm.startDir)
'
'            Case persons
'                fPath = GetSaveAsFname(myMen u.type + 2, WorkClsm.startDir)
'
'        End Select
    End If
    
    If fPath <> "NoPath" Then
        Open fPath For Output As #1: Print #1, sTempStr: Close
        
        PreSaveSetButton True 'Вызвать кнопку "Готово"
        RefreshLabelInfo "Изменения сохранены.", , True
    End If
End Sub
'############################################################
'процедура очищает поля ввода
Sub ClearTxtBxes()
    Me.TextBox1 = "": Me.TextBox2 = "": Me.TextBox3 = "": Me.TextBox4 = "": Me.TextBox5 = ""
    Me.tboxSearchConstr = "": Me.listDataBase.ListIndex = -1
End Sub
'############################################################
'функция корректирует падеж при выводе инфомации о количестве источников
Function RusPadej3(value As Integer)
    Select Case value
        Case 1, 21, 31, 41, 51, 61, 71, 81, 91, 101, 121
            RusPadej3 = "Импортировано " & value & " наименование."
        Case 2 To 4, 22 To 24, 32 To 34, 42 To 44, 52 To 54, 62 To 64, 72 To 74, 82 To 84, 92 To 94
            RusPadej3 = "Импортировано " & value & " наименования."
        Case Else
            RusPadej3 = "Импортировано " & value & " наименований."
    End Select
End Function
'############################################################
'процедура открывает каталог текущего шаблона
Private Sub cmbOpenTemplateFolder_Click()

'    Dim sTempDir As String, sBaseDir As String, sFifNum As String, sTypeName As String
'    sBaseDir = WorkClsm.templatesDir
'    sFifNum = sArrDataBase(2, Me.listDataBase.ListIndex): sTypeName = sArrDataBase(1, Me.listDataBase.ListIndex)
'    sTempDir = Dir(sBaseDir & "*" & sFifNum & "*", vbDirectory)  'каталог СИ
'
'    Explorer.OpenFolder sBaseDir & sTempDir & "\", True
'
'
    Dim sTempDir As String, _
        templatesDir As String, _
        fifRegNum As String, _
        sTypeName As String
     ' ------------------------------------------------------
'todo: отвязать workclsm
    templatesDir = WorkClsm.templatesDir & Application.PathSeparator
    fifRegNum = sArrDataBase(2, Me.listDataBase.ListIndex)
        
    Dim currTemplateDir As String
    currTemplateDir = Dir(templatesDir & "*" & fifRegNum & "*", vbDirectory)   'каталог выбранного СИ
    
    Dim refFifPath As String
    refFifPath = templatesDir & currTemplateDir & Application.PathSeparator & REFERENCE_FIFNUM_FILENAME
    
    If FileExist(refFifPath) Then _
        fifRegNum = Base.ContentFromFile(refFifPath)
    
    sTypeName = sArrDataBase(1, Me.listDataBase.ListIndex)
    sTempDir = Dir(templatesDir & "*" & fifRegNum & "*", vbDirectory)  'каталог СИ
    
    Explorer.OpenFolder templatesDir & sTempDir & "\", True
End Sub
'#########################################################
'передать в свойства открытого файла данные для шаблона
Private Sub cmbFillTempProp_Click()
' ------------------------------------------------------
'todo: отвязать workclsm
    Me.cmbFillTempProp.Enabled = False: FillIfXl WorkClsm, sArrDataBase, Me.listDataBase.ListIndex
End Sub
'#########################################################
'человек является поверителем
Private Sub chbVerRefer_Change()

    Me.chbVerRefer.foreColor = &H80000007
    
    If Me.chbVerRefer Then _
        TrueElementForeColor Me.chbVerRefer
    
    If UMenu.typе = instrumentsOLD Then
    
        With Me.cmbImport
            
            .foreColor = Me.chbVerRefer.foreColor
            
            If Me.chbVerRefer = False Then
                .caption = "Импорт"
                .ControlTipText = "Импорт данных из существующей базы"
                
                If bolUF_Cnstr_Load And bolUserInput Then
                    'удалить ссылку на перекрёстную ссылку, если таковая имеется
                    Dim regFifNum As String
                    regFifNum = Me.TextBox3
                    
                    Dim currentInstrumentDir As String
                    ' ------------------------------------------------------
'todo: отвязать workclsm
                    currentInstrumentDir = Dir(WorkClsm.templatesDir & "*" & regFifNum & "*", vbDirectory) 'каталог СИ
                    
                    Dim referenceFilePath As String
                    ' ------------------------------------------------------
'todo: отвязать workclsm
                    referenceFilePath = WorkClsm.templatesDir & currentInstrumentDir & Application.PathSeparator & REFERENCE_FIFNUM_FILENAME
                    
                    If FileExist(referenceFilePath) Then _
                        Kill referenceFilePath
                End If
                
            End If
            
            If Me.chbVerRefer And .caption = "Импорт" Then
                .caption = "Выбрать"
                .ControlTipText = "Выбрать элемент базы для перекрёстной ссылки"
            End If
            
            If Me.chbVerRefer And .caption <> "Выбрать" Then
                .ControlTipText = "Перейти к элементу перекрёстной ссылки"
            End If
        End With
        
    End If
    
    If bolUF_Cnstr_Load And bolUserInput Then
    
        Dim iUboundArr As Integer
        iUboundArr = UBound(sArrDataBase) - 2 'по умолчанию для средств измерений - предпоследний элемент
        
        If UMenu.typе = personsOLD Then _
            iUboundArr = UBound(sArrDataBase) 'для фамилий и должностей
        
        If Me.chbVerRefer = False Then _
            sArrDataBase(iUboundArr, listDataBase.ListIndex) = "nodata"
            
        If UMenu.typе = personsOLD And Me.chbVerRefer Then _
            sArrDataBase(iUboundArr, listDataBase.ListIndex) = "поверитель"
        
        If UMenu.typе = instrumentsOLD Then
            
            iDataBaseIndex = 0
            
            If Me.chbVerRefer Then _
                iDataBaseIndex = listDataBase.ListIndex 'запонить позицию
        End If
    
        bolAlreadySaved = False: bolUserInput = False: PreSaveSetButton
    End If
End Sub
Private Sub chbVerRefer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    bolUserInput = True
End Sub
'#########################################################
'средство измерений является эталоном
Private Sub chbEtalon_Change()
    Me.chbEtalon.foreColor = &H80000007: If Me.chbEtalon Then TrueElementForeColor Me.chbEtalon
    
    If bolUF_Cnstr_Load And bolUserInput Then
    
        Dim iUboundArr As Integer
        iUboundArr = UBound(sArrDataBase) - 1 'по умолчанию для средств измерений

        If Me.chbEtalon = False Then sArrDataBase(iUboundArr, listDataBase.ListIndex) = "nodata"
        If Me.chbEtalon Then sArrDataBase(iUboundArr, listDataBase.ListIndex) = "etalon"
        
        iDataBaseIndex = 0: If Me.chbEtalon Then iDataBaseIndex = listDataBase.ListIndex 'запонить позицию
        bolAlreadySaved = False: bolUserInput = False: PreSaveSetButton
    End If
End Sub
Private Sub chbEtalon_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    bolUserInput = True
End Sub

Private Sub FillInputData(sArrKeyCode() As String, StrToSearch As String, ByRef RusInput As String, ByRef EngInput As String)
    'RusInput - переменная поискового значения на русском языке
    'EngInput - переменная поискового значения на английском языке
    Dim i As Integer, j As Integer, sSym As String
    StrToSearch = Replace(StrToSearch, "[", "{") 'чтобы не было ошибки поиска
    For i = 1 To Len(StrToSearch) 'для каждого символа, записанного в строке
        sSym = Mid(StrToSearch, i, 1) 'обрабатываемый символ
        For j = LBound(sArrKeyCode, 2) To UBound(sArrKeyCode, 2) 'пройтись по массиву кейкодов
            If sSym = sArrKeyCode(0, j) Then 'найдено соответствие среди символов на русском языке
                RusInput = RusInput & sArrKeyCode(0, j) 'искомое значение на русской раскладке
                EngInput = EngInput & sArrKeyCode(1, j) 'искомое значение на английской раскладке
                Exit For
            ElseIf sSym = sArrKeyCode(1, j) Then 'найдено соответствие среди символов на английском языке
                RusInput = RusInput & sArrKeyCode(0, j) 'искомое значение на русской раскладке
                EngInput = EngInput & sArrKeyCode(1, j) 'искомое значение на английской раскладке
                Exit For
            End If
        Next j
    Next i
End Sub
