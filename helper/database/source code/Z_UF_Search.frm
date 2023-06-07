VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_UF_Search 
   Caption         =   "База данных заказчиков/средств измерений/эталонов"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   OleObjectBlob   =   "Z_UF_Search.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Z_UF_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit 'запрет на использование неявных переменных

Const REFERENCE_FIFNUM_FILENAME = "fifRegNum.ref"

Private myBase As New Z_clsmBase, _
        WorkClsm As New Z_clsmSearch, _
        myMi As MeasInstrument

Dim sArrKeyCode() As String, _
    sArrDataBase() As String

Private itemIndex As String, _
        bolTbEntry As Boolean
    
Private fso As New FileSystemObject

Private Sub UserForm_Initialize() 'загрузка формы
    
    Set_Z_UF_Search_Size 'задать индивидуальные свойства для формы поиска по БД
    GetMyConfigFromFile myBase, WorkClsm 'загрузить настройки из файлов
    ' ------------------------------------------------------
    TransferConfig
    ' ------------------------------------------------------

' ------------------------------------------------------
'todo: привязать к файлу данных cache.db
    If UMenu.typе <> archiveOLD Then _
        FillArrDataBaseFromFile sArrDataBase(), WorkClsm.DbName  'заполнить массив основной базы данных
' ------------------------------------------------------
    If UMenu.typе = personsOLD Then 'только для формы фамилий и должностей
        ' ------------------------------------------------------
'todo: привязать к файлу данных cache.db
        If WorkClsm.isFullName = "True" Then _
            WorkClsm.FullFirstName = True 'если в настройкх выбрано указание полного имени
' ------------------------------------------------------
'todo: привязать к файлу данных cache.db
        If WorkClsm.FullFirstName = "недоступно" Then _
            WorkClsm.FullFirstName = False
            
        Me.chbFullName = WorkClsm.FullFirstName 'передать параметр в чекбокс
        
    End If
    
    If UMenu.typе = archiveOLD Then 'архив

        FillArchivedata sArrDataBase 'основной массив с данными найденных архивных работ
        Me.LabelInfo.caption = DataBase.LabelInfoText
        
    End If
        

    
    TrueElementForeColor Me.LabelInfo 'покрасить информационную метку в зависимости от типа приложения
' ------------------------------------------------------
'todo: привязать к серверу
    FillArrKeycodeFromFile 'заполнить массив кейкодов
' ------------------------------------------------------

    If UMenu.typе = etalonsOLD Then _
        SortMassBiv sArrDataBase, UBound(sArrDataBase) 'сортировать массив по ключевому слову
        
    UpdateTboxAndListbox 'передать последние сведения в текстбокс
    
    If Me.listResults.ListCount = 0 Then _
        UpdateListResults sArrDataBase

    SetEventControls Me 'инициировать групповые события для всех контролов

    InsertSearchData
    
    With Me.tboxSearch
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 26.02.2023 10:46
    ' Назначение: для перехода от старых настроек к новым
    ' ----------------------------------------------------------------
    Private Sub TransferConfig()
        
        If Config.sourceDataPath = Base.defaultValue Then _
            Config.sourceDataPath = WorkClsm.startDir
        
        If Config.sandboxPath = Base.defaultValue Then _
            Config.sandboxPath = WorkClsm.workDir
        
        If WorkClsm.ArchivePath <> vbNullString Then
        
            If Config.archiveLocalPath = Base.defaultValue Then _
                Config.archiveLocalPath = WorkClsm.ArchivePath
            
        End If
        
        Config.Save
        
    End Sub
Private Sub UserForm_Activate()
    
    If Not UMenu.isLoaded Then _
        VBA.Unload Me
    
End Sub
'#########################################################
'вставить в строку поиска данные о СИ
Private Sub InsertSearchData()

    If UMenu.typе = instrumentsOLD Then 'только для БД СИ
        
        Properties.SetTargetWorkbook ActiveWorkbook
        
        Dim sKeyWord As String
        sKeyWord = Properties.Keywords 'получить номер в фиф

        If sKeyWord <> vbNullString Then
                        
            If InStr(sKeyWord, "-") = 0 Then 'заполнение номера в фиф из АРШИНА
                
                If Len(sKeyWord) >= 8 Then
                    sKeyWord = Left(sKeyWord, 8)
                    sKeyWord = Replace(sKeyWord, ".", "-")
                    
                End If
                
            End If
            
            Me.tboxSearch = sKeyWord
            
            Dim i As Integer
            For i = LBound(Me.listResults.List) To UBound(Me.listResults.List)
                If InStr(Me.listResults.List(i), sKeyWord) > 0 Then Me.listResults.Selected(i) = True: Exit Sub
            Next i
        End If
    End If
End Sub

    ' ----------------------------------------------------------------
    ' Дата: 09.03.2023 18:31
    ' Назначение: получить массив данных архива
    ' ----------------------------------------------------------------
    Private Sub FillArchivedata( _
        ByRef sArrToFill() As String, _
        Optional key As String _
        )
        
        Dim sArchivePath As String, sTempStr As String
        sArchivePath = Config.archiveLocalPath

        Dim arc As New Collection
        Set arc = DataBase.FilterArchive(key)
        
        If Not CBool(arc.count) Then _
            ReDim sArrToFill(0, 0): _
            Exit Sub
            
            
        Dim tempCol As New Collection
        
            
        ReDim sArrToFill(0, arc.count - 1)
        
        Dim i As Integer
        For i = 1 To arc.count
            sArrToFill(0, i - 1) = arc(i)
        Next i
        
        
'        If sArrDataBase(LBound(sArrDataBase), UBound(sArrDataBase, 2)) <> vbNullString Then _
'            SortMassBiv sArrToFill, , , False
'
    End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If Shift = 2 And UMenu.typе = instrumentsOLD Then _
        Me.cmb1.caption = "*"
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
End Sub
Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If UMenu.typе = instrumentsOLD Then _
        Me.cmb1.caption = "Н"
        
End Sub

'#########################################################
'процедура срабатывает при закрытии формы
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    myBase.GetArrFF setDir, "settings.ini" 'загрузить в класс файл настроек
    myBase.SaveProperties WorkClsm.myParameters, WorkClsm.myValues 'записать изменения в файл настроек
    ' ------------------------------------------------------
    'todo: сохранить Cache файл -- db.close
    
End Sub
'#########################################################
'процедура срабатывает при выгрузке формы
Private Sub UserForm_Terminate()

    UMenu.isLoaded = False
    
    Set myBase = Nothing
    Set WorkClsm = Nothing
    
    ClearSingletone
End Sub

'#########################################################
'поиск совпадений по базе данных
Private Sub tboxSearch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If bolTbEntry = False Then 'переменная первого входа
    
        With Me.tboxSearch
            
            .SelStart = 0
            .SelLength = Len(.text)
            
        End With
        
        bolTbEntry = True
        
    End If
End Sub
Private Sub tboxSearch_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    ) 'поиск по организациям
    
    Dim i As Integer, _
    bolSeleceted As Boolean
    
    If KeyCode = 13 Then 'нажатие enter
        
        For i = LBound(listResults.List) To UBound(listResults.List)
            
            If listResults.Selected(i) = True Then _
                bolSeleceted = True: _
                Exit For 'была ли выделена хоть одна строка
                
        Next i
        
        If bolSeleceted Then  'найдена выделенная строка
            
            If Shift = 2 Then 'зажат CTRL
                TransferManufacturer
                
            ElseIf Shift = 1 Then 'shift
                
                If Me.btnOpenFolder.Enabled Then _
                    btnOpenFolder_Click
            
            Else
            
                If Me.cmbProtSv.Enabled Then _
                    cmbProtSv_Click: Exit Sub 'шаблон
                    
                If Me.cmbDescription.Enabled Then _
                    cmbDescription_Click: _
                    Exit Sub 'описание типа
                    
                If Me.cmbMetodic.Enabled Then _
                    cmbMetodic_Click: _
                    Exit Sub 'методика поверки
                
                cmbOneClick 'полное наименование объекта БД
                
            End If
        Else 'строка не была выделена
            If UMenu.typе = etalonsOLD And myWdDoc = False And _
                Me.cmbProtSv.Enabled Then cmbProtSv_Click: Exit Sub 'поиск и заполнение сведений об эталонном оборудовании
                
            Me.listResults.Selected(0) = True 'выделить первый элемент поискового запроса
        End If
        
    End If
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
End Sub
    Private Sub TransferManufacturer()
        
        If Not UMenu.typе = organisationsOLD Then _
            Exit Sub
            
        Application.ScreenUpdating = False
        
        Dim customerCell As Range
        If FindCellRight("Изготовитель:", customerCell, ThisCell:=True) Then
            
            customerCell.Offset(0, 1) = sArrDataBase(0, CInt(itemIndex)) ' наименование
        Else
            ActiveCell = sArrDataBase(0, CInt(itemIndex)) ' наименование
            ActiveCell.Offset(1, 0).Select
        End If
        
        Application.ScreenUpdating = True
        VBA.Unload Me
        
    End Sub
    
Private Sub tboxSearch_Change()
    DisableButtons 'отключить все кнопки формы
    
    If tboxSearch = "" Then
        UpdateListResults sArrDataBase  'обновить содержимое листбокса1
    Else
        Dim sTempArr() As String
        ReDim sTempArr(UBound(sArrDataBase), 0) 'выделить память под временный массив
        
        If sArrKeyCode(0, 0) = "" Then 'массив кейкодов не был загружен
            FindInBivArr sArrDataBase, sTempArr, Me.tboxSearch  'получить массив поисковых совпадений
        Else ' массив кейкодов был загружен
            
            Dim sInputRus As String, sInputEng As String, sInputString As String
            FillInputData sArrKeyCode, Me.tboxSearch, sInputRus, sInputEng  'получить значения для поиска по массиву
            sInputString = sInputRus & " " & sInputEng

            FindInBivArr sArrDataBase, sTempArr, sInputString  'получить массив поисковых совпадений значениями на русском языке

        End If

        UpdateListResults sTempArr  'обновить данные совпадений по массиву sTempArr
    End If
    ' ------------------------------------------------------
    'todo: привязать к db.LastSearch
    WorkClsm.LastSearch = Me.tboxSearch 'передать значение в массив
    
    Dim iListCount As Integer
    iListCount = listResults.ListCount - 1
    
    If UMenu.typе = archiveOLD Then _
        iListCount = listResults.ListCount
        
    Me.labelUnderSearchField.caption = RusPadejCoincidence(iListCount, Me.labelUnderSearchField)

    If tboxSearch = "" Then _
        Me.labelUnderSearchField.caption = "поиск совпадений": Me.labelUnderSearchField.foreColor = &H80000012 '- чёрный
    
    If listResults.ListCount = 2 Then _
        If listResults.List(1) = "внести изменения в базу данных..." Then listResults.Selected(0) = True 'выбрать первый элемент
    
    If listResults.ListCount = 1 Then _
        If listResults.List(0) <> "внести изменения в базу данных..." Then listResults.Selected(0) = True
End Sub
Private Sub tboxSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    bolTbEntry = False
End Sub
'#########################################################
'инвертировать отображение фамилии и инициалов
Private Sub chbFullName_Change()
    
    If Me.chbFullName Then
        TrueElementForeColor Me.chbFullName
        
        If UMenu.typе = personsOLD Then _
            Me.cmb1.caption = "ФИО полностью"
    Else
        Me.chbFullName.foreColor = &H80000007
        
        If UMenu.typе = personsOLD Then _
            Me.cmb1.caption = "Фамилия И.О."
    End If
    
    If UMenu.isLoaded Then

        If UMenu.typе = personsOLD Then 'фамилии
            ' ------------------------------------------------------
            'todo: привязать к db
            WorkClsm.FullFirstName = Me.chbFullName
            tboxSearch_Change
        End If
        
        If UMenu.typе = archiveOLD Then 'архив работ
            If Me.chbFullName Then
            
                If MsgBox("Защитить книгу?", vbYesNo) = vbYes Then _
                    ProtectSheets: _
                    Me.chbFullName = Not Me.chbFullName: _
                    VBA.Unload Me
                    
            End If
        End If
    End If
End Sub
Private Sub chbFullName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then cmbOneClick
End Sub
'############################
'результат поиска совпадений
Private Sub listResults_Click()

    Dim sLB1str As String
    Call DisableButtons: sLB1str = listResults.text  'выбранная строка
    
    If sLB1str <> "внести изменения в базу данных..." Then
    ' ------------------------------------------------------
    'todo: привязать к db.lastsearch
        WorkClsm.LastIndex = listResults.ListIndex  ' индекс текущей выбранной строки
        WorkClsm.constrSrch = GetconstrSrch(sLB1str)  'получить элемент для поиска в конструктора
        
        GetDBindAndMyMi sLB1str
    End If
End Sub
'############################
'дополнительная подфункция
Private Sub GetDBindAndMyMi(sLB1str As String, Optional DontEnableButtons As Boolean)

    itemIndex = DataBaseIndex(sLB1str) 'получить индекс элемента в общем массиве
    If IsNumeric(itemIndex) Then
    
        If DontEnableButtons = False Then
            EnableButtons 'активировать кнопки в завивисимости от состояния заполнения ячеек данных массива базы данных
            Me.labelUnderSearchField.caption = "нажмите Enter": TrueElementForeColor Me.labelUnderSearchField
        End If
       
        With myMi
            .sName = sArrDataBase(0, CInt(itemIndex))
            If UMenu.typе <> archiveOLD Then
                .sType = sArrDataBase(1, CInt(itemIndex)): .sFif = sArrDataBase(2, CInt(itemIndex))
                
                If UMenu.typе <> personsOLD Then
                    .sMetodic = sArrDataBase(3, CInt(itemIndex)): .sRef = sArrDataBase(UBound(sArrDataBase) - 2, CInt(itemIndex))
                    .bolEtal = False: If sArrDataBase(UBound(sArrDataBase) - 1, CInt(itemIndex)) <> "nodata" Then .bolEtal = True
                End If
            End If
        End With
    End If
End Sub

'############################
'функция выделяет компонент для передачи в переменную поиска в конструкторе
Function GetconstrSrch(ByVal sSelectedLBstr As String) As String
    
    Dim sArrTemp() As String
    sArrTemp = Split(sSelectedLBstr, " \\ ")
    
    Select Case UMenu.typе
    
        Case organisationsOLD
            Dim i As Integer 'поисковый запрос = наименование организации
            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                If sArrDataBase(0, i) = sArrTemp(0) Then GetconstrSrch = sArrDataBase(0, i): Exit Function
            Next i
           
        Case instrumentsOLD, etalonsOLD
            GetconstrSrch = sArrTemp(0)   'номер фиф или же ключевое слово
            
        Case personsOLD
            sArrTemp = Split(sArrTemp(0), " "): GetconstrSrch = sArrTemp(0) 'фамилия
            
    End Select
    
End Function
Private Sub listResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'enter
        If Me.cmbProtSv.Enabled Then Me.cmbProtSv.SetFocus: Exit Sub 'шаблон
        If Me.cmbDescription.Enabled Then Me.cmbDescription.SetFocus: Exit Sub 'описание типа
        If Me.cmbMetodic.Enabled Then Me.cmbMetodic.SetFocus: Exit Sub 'методика поверки
        
        Me.cmb1.SetFocus 'полное наименование объекта БД
    End If
End Sub
Private Sub listResults_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If listResults.text = "внести изменения в базу данных..." Then _
        VBA.Unload Me: Z_UF_Constructor.Show 0 'загрузка формы конструктора БД
End Sub
Private Sub listResults_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If listResults.text = "внести изменения в базу данных..." Then _
        VBA.Unload Me: Z_UF_Constructor.Show 0 'загрузка формы конструктора БД
End Sub
'#########################################################
'функция определяет позицию элемента в общем массиве
Function DataBaseIndex(sFindStr As String) As Integer
    DataBaseIndex = -1 'по умолчанию
    
    If UMenu.typе <> archiveOLD Then _
        sFindStr = Left(sFindStr, InStr(sFindStr, " \\ ") - 1) ' строка для поиска в массиве базы данных
    
    Dim i As Integer, j As Integer, bolExitFor As Boolean, iUbound As Integer
    iUbound = UBound(sArrDataBase)
    
    If UMenu.typе = instrumentsOLD Then _
        iUbound = LBound(sArrDataBase) + 3
    
    For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
        For j = LBound(sArrDataBase) To iUbound
            If InStr(sArrDataBase(j, i), sFindStr) > 0 Then bolExitFor = True: Exit For 'найдена строка базы данных
        Next j
        If bolExitFor Then Exit For
    Next i

    If i > UBound(sArrDataBase, 2) Then 'второй проход - попытка найти совпадение при его отсутствии
        If InStr(sFindStr, " ") > 0 Then 'только если есть пробел в поисковой строке
        
            Dim sArrTemp() As String
            sArrTemp = Split(sFindStr, " "): sFindStr = sArrTemp(0)
            
            For i = LBound(sArrTemp) To UBound(sArrTemp)
                If Len(sArrTemp(i)) > Len(sFindStr) Then sFindStr = sArrTemp(i)
            Next i
    
            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                For j = LBound(sArrDataBase) To UBound(sArrDataBase)
                    If InStr(sArrDataBase(j, i), sFindStr) > 0 Then bolExitFor = True: Exit For 'найдена строка базы данных
                Next j
                If bolExitFor Then Exit For
            Next i
            
        End If
    End If
    
    If i <= UBound(sArrDataBase, 2) Then
        If sArrDataBase(0, i) <> "" Then DataBaseIndex = i
    End If
    
End Function
' ----------------------------------------------------------------
' Дата: 25.02.2023 13:01
' Назначение:
'    параметр DbIndex:
' ----------------------------------------------------------------
Private Sub EnableButtons( _
    )
    Select Case True
    
        Case UMenu.typе = organisationsOLD
            EnableForOrganisations
        ' ----------------------------------------------------------------
        Case UMenu.typе = instrumentsOLD
            EnableForInstruments
        ' ----------------------------------------------------------------
        Case Else
            EnableForArchive
            
    End Select

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.02.2023 13:27
    ' Назначение:
    ' ----------------------------------------------------------------
    Private Sub EnableForOrganisations()

        Dim itemKey As String
        itemKey = sArrDataBase(LBound(sArrDataBase) + 2, CInt(itemIndex)) 'ключевое слово для организации
        
        If DataBase.IsDirAvailable(itemKey) Then

            Me.btnOpenFolder.Enabled = True
            Me.btnOpenFolder.BackColor = Colors.oragnePastel
            
        End If

        If sArrDataBase(0, itemIndex) <> "nodata" Then 'наименование краткое
        
            Me.cmb1.Enabled = True 'передать на лист
            Me.cmb1.BackColor = Colors.greenPastel
            
            Me.cmb2.Enabled = True 'наименование
            
        End If
        
        If sArrDataBase(1, itemIndex) <> "nodata" Then _
            Me.cmb3.Enabled = True 'инн
        
        If sArrDataBase(3, itemIndex) <> "nodata" Then _
            Me.cmb4.Enabled = True 'адрес
    
    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.02.2023 13:36
    ' Назначение:
    ' ----------------------------------------------------------------
    Private Sub EnableForInstruments()

        If sArrDataBase(0, itemIndex) <> "nodata" Then _
            Me.cmb1.Enabled = True 'наименование СИ
        
        If sArrDataBase(1, itemIndex) <> "nodata" Then _
            Me.cmb2.Enabled = True 'Тип СИ
            
        If sArrDataBase(2, itemIndex) <> "nodata" Then _
            Me.cmb3.Enabled = True 'рег. номер ФИФ
        
        If sArrDataBase(3, itemIndex) <> "nodata" Then _
            Me.cmb4.Enabled = True 'методика поверки
        ' ----------------------------------------------------------------
        Dim keySearch As String 'ключ поиска — рег номер
        keySearch = sArrDataBase(LBound(sArrDataBase) + 2, CInt(itemIndex))
        
        Dim targetDir As String
        targetDir = ItemDirectory(keySearch) '\instruments\Canberra_18509-04
          
        If targetDir <> vbNullString Then
            
            Me.btnOpenFolder.Enabled = True

            
            EnableDescriptionButton targetDir
            EnableMethodicButton targetDir
            EnableLoadTemplateButton
               
        End If
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 25.02.2023 14:54
        ' Назначение:
        '    параметр itemKey:
        ' Возвращаемый тип: String
        ' ----------------------------------------------------------------
        Private Function ItemDirectory( _
            itemKey As String _
            ) As String
            
            
'            Dim itemPath As String
'            itemPath = Config.instrumentsPath
'            ' ----------------------------------------------------------------
'            If Right(itemPath, 1) <> Application.PathSeparator Then _
'                itemPath = itemPath & Application.PathSeparator
'            ' ----------------------------------------------------------------
'            Dim result As String
'            result = Dir(itemPath & "*" & itemKey & "*", vbDirectory) '\instruments\Canberra_18509-04
            
            Dim itemPath As String
            itemPath = fso.BuildPath(Config.instrumentsPath, "*" & itemKey & "*")
            
            Dim result As String
            result = Dir(itemPath, vbDirectory) '\instruments\Canberra_18509-04
            
            ItemDirectory = result

            
        End Function
        Private Sub EnableDescriptionButton( _
            targetDir As String _
            )
            
            Dim targetPath As String
            targetPath = fso.BuildPath(Config.instrumentsPath, targetDir)
        
            Dim targetTypeDesc As String
            targetTypeDesc = Dir(targetPath & "\" & "*" & "ot_" & "*") '\instruments\Canberra_18509-04\mp_xxxxx
            
            If targetTypeDesc <> vbNullString Then
               
               Me.cmbDescription.Enabled = True
               Me.cmbDescription.BackColor = Colors.greenPastel
               
            End If
        End Sub
        Private Sub EnableMethodicButton( _
            targetDir As String _
            )
            
            Dim targetPath As String
            targetPath = fso.BuildPath(Config.instrumentsPath, targetDir)
            
            Dim targetMp As String
            targetMp = Dir(targetPath & "\" & "*" & "mp_" & "*") '\instruments\Canberra_18509-04\mp_xxxxx
            
            If targetMp <> vbNullString Then
                
                Me.cmbMetodic.Enabled = True
                Me.cmbMetodic.BackColor = Colors.greenPastel
                
            End If
               
        End Sub
        Private Sub EnableLoadTemplateButton( _
            )
            
            Dim itemKey As String 'ключ поиска — рег номер
            itemKey = sArrDataBase(LBound(sArrDataBase) + 2, CInt(itemIndex))
            
            Dim itemReferenceKey As String 'перекрёстная ссылка
            itemReferenceKey = sArrDataBase(LBound(sArrDataBase) + 4, CInt(itemIndex))
            ' ----------------------------------------------------------------
            Dim targetKey As String
            targetKey = itemKey
            
            If itemReferenceKey <> "nodata" Then _
                targetKey = itemReferenceKey
            ' ----------------------------------------------------------------
            Dim targetDir As String
            targetDir = ItemDirectory(targetKey) '\instruments\Canberra_18509-04
            ' ----------------------------------------------------------------
            Dim targetPath As String
            targetPath = fso.BuildPath(Config.instrumentsPath, targetDir)
            ' ----------------------------------------------------------------
            Dim targetTemplate As String
            targetTemplate = Dir(targetPath & Application.PathSeparator & "*" & "body_" & "*") '\instruments\Canberra_18509-04\mp_xxxxx
            
            If targetTemplate = vbNullString Then _
                targetTemplate = Dir(targetPath & Application.PathSeparator & "*" & "pr_" & "*")  '\instruments\Canberra_18509-04\mp_xxxxx
            ' ----------------------------------------------------------------
            If targetTemplate = vbNullString Then _
                Exit Sub
            ' ----------------------------------------------------------------
            Me.cmbProtSv.Enabled = True
            Me.cmbProtSv.BackColor = Colors.oragnePastel
            
        End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.02.2023 13:38
    ' Назначение:
    ' ----------------------------------------------------------------
    Private Sub EnableForArchive()
           
        Dim j As Byte
        For j = LBound(sArrDataBase) To UBound(sArrDataBase) 'проход по всем полям массива
        
            If sArrDataBase(j, itemIndex) <> "nodata" Then
            
                If j < 4 Then _
                    Me.Controls("cmb" & j + 1).Enabled = True

            End If
        Next j
        
        
    End Sub
    
    

'        sFifNum = sArrData(LBound(sArrData) + 2, i)
'        referenceRegFifNum = sArrData(LBound(sArrData) + 4, i) 'перекрёстная ссылка
'
'        sTempDir = Dir(sStartDir & "*" & sFifNum & "*", vbDirectory) 'каталог СИ
'        sRefDir = Dir(sStartDir & "*" & referenceRegFifNum & "*", vbDirectory) 'каталог перекрёстной ссылки
'
'        If sTempDir <> vbNullString Then 'каталог с номером в фиф обнаружен
'
'            If Dir(sTempPath & "\pr" & "*" & sFifNum & "*.xls*") <> vbNullString Or _
'                    Dir(sTempPath & "\body" & "*" & sFifNum & "*.xls*") <> vbNullString Then bolTMP = True 'наличие шаблонов
'
'            If Dir(sRefPath & "\pr" & "*" & referenceRegFifNum & "*.xls*") <> vbNullString Or _
'                    Dir(sRefPath & "\body" & "*" & referenceRegFifNum & "*.xls*") <> vbNullString Then bolRef = True 'наличие перекрёстного шаблона
'
'            If bolRef Then 'опознано наличие перекрёстной ссылки - приоритет загрузки
'                If sTempStr <> vbNullString Then sTempStr = sTempStr & "+"
'                sTempStr = sTempStr & "ШБ*"
'            Else
'                If bolTMP Then 'опознано наличие шаблона
'                    If sTempStr <> vbNullString Then sTempStr = sTempStr & "+"
'                    sTempStr = sTempStr & "ШБ"
'                End If
'            End If

'        End If
'    Next
'End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

'########Зкшмфеу#################################################
'процедура правильно переводит фокус на кнопки
Sub TrueBtnFocus(ACtrlName As String)
    Dim iCtrlInd As Integer
    If ACtrlName Like "cmb*" And IsNumeric(Right(ACtrlName, 1)) Then 'только для кнопок
        iCtrlInd = Right(ACtrlName, 1) 'номер кнопки
        Do
            iCtrlInd = iCtrlInd + 1 'следующая кнопка
            If iCtrlInd = 5 Then Me.cmbReady.SetFocus:  Exit Do
            If iCtrlInd < 5 Then Me.labelUnderSearchField.caption = "нажмите Enter": Call TrueElementForeColor(Me.labelUnderSearchField) 'Me.labelUnderSearchField.ForeColor = &H8000& ' - зелёный
            
            If Me.Controls("cmb" & iCtrlInd).Enabled = True Then
                If Me.Controls("cmb" & iCtrlInd).Visible = True Then
                    Me.Controls("cmb" & iCtrlInd).SetFocus
                    Exit Do
                End If
            End If
        Loop
        
    Else 'если активный контрол - не кнопка, то выбрать первую кнопку
        TrueBtnFocus "cmb1"
    End If
End Sub
'#########################################################
'наименование Организации+ИНН / наименование СИ / наименование эталона / фамилия сотрудника
Private Sub cmb1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If myWdDoc = False Then _
        InsertInstrumentName Shift: _
        Exit Sub
        
    cmbOneClick
End Sub
Private Sub cmb1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If myWdDoc = False Then InsertInstrumentName Shift: Exit Sub
    cmbOneClick
End Sub
'#########################################################
'процедура обрабатывает клик в форме поиска по средствам измерений
Sub InsertInstrumentName( _
    ByVal Shift As Integer _
    )
    
    If Shift = 2 Then
        
        If UMenu.typе <> instrumentsOLD Or myWdDoc Then _
            cmbOneClick: _
            Exit Sub
            
        If MsgBox("Будет произведён поиск и замена параметров наименований. Продолжить?", vbYesNo) = vbNo Then _
            Me.cmb1.caption = "Н": _
            Exit Sub
        
        ChangeXlPropertyComment myMi 'передать тип СИ
        'SetBuiltInProperty "Comments", sArrDataBase(1, CInt(itemIndex)) 'тип СИ
        SetBuiltInProperty "Keywords", sArrDataBase(2, CInt(itemIndex)) 'номер в фиф
        
        FillNameInstrument myMi
        Commit_factory_number 'добавить заводской номер в типу СИ
        ActiveWorkbook.Save
        
        VBA.Unload Me
        Exit Sub
    End If
    
    cmbOneClick
End Sub
Private Sub cmbOneClick()
    Select Case UMenu.typе
        Case organisationsOLD 'наименование организации заказчика + ИНН передать
            
            Application.ScreenUpdating = False
            
            Dim customerCell As Range
            If FindCellRight("Заказчик:", customerCell, ThisCell:=True) Then
                
                customerCell.Offset(0, 1) = sArrDataBase(0, CInt(itemIndex)) ' наименование
                
                If InStr(customerCell.Offset(1, 0), "ИНН") > 0 Then _
                    customerCell.Offset(1, 1) = sArrDataBase(1, CInt(itemIndex)) ' ИНН
                
                If sArrDataBase(3, CInt(itemIndex)) <> "nodata" Then
                
                    If InStr(customerCell.Offset(1, 0), "адрес") > 0 Then customerCell.Offset(1, 1) = sArrDataBase(3, CInt(itemIndex)) 'Адрес
                    If InStr(customerCell.Offset(1, 2), "адрес") > 0 Then customerCell.Offset(1, 3) = sArrDataBase(3, CInt(itemIndex)) 'Адрес
                End If
                
                
                SetBuiltInProperty "Company", sArrDataBase(0, CInt(itemIndex))                 'передать сведения заказчика в свойство книги
              '  SetBuiltInProperty "Category", sArrDataBase(0, CInt(itemIndex))                 'передать сведения заказчика в свойство книги
              
               ' TransferManufacturer
                
                Application.ScreenUpdating = True
                VBA.Unload Me
            Else
                MsgBox "Поле заказчика не найдено"
                Application.ScreenUpdating = True
                
                Exit Sub
            End If
            
            
        ' ------------------------------------------------------
        'todo: привязать к db
        Case personsOLD 'фамилии
            DataTransfer sArrDataBase(0, CInt(itemIndex)), True, WorkClsm.FullFirstName 'наименование СИ / фамилия сотрудника
            TrueBtnFocus ActiveControl.name
        
        Case archiveOLD 'архив работ
            Explorer.OpenFolder Config.archiveLocalPath & Application.PathSeparator & sArrDataBase(0, CInt(itemIndex)), True
            
        Case instrumentsOLD 'средства измерений
            DataTransfer sArrDataBase(0, CInt(itemIndex)) & " " & sArrDataBase(1, CInt(itemIndex)), True
            TrueBtnFocus ActiveControl.name
        Case Else
            DataTransfer sArrDataBase(0, CInt(itemIndex)), True  'наименование СИ / фомилия сотрудника
            TrueBtnFocus ActiveControl.name
    End Select
End Sub
'#########################################################
'Наименование организации / модификация СИ / ?????тип эталона
Private Sub cmb2_Click()

    Select Case UMenu.typе
    
        Case organisationsOLD 'сведения заказчиков
            DataTransfer sArrDataBase(0, CInt(itemIndex)), True, , True
            
        Case Else 'СИ / эталоны
            DataTransfer sArrDataBase(1, CInt(itemIndex)), , , False
        '    Z_UF_Search_Cmb2 sArrDataBase, CInt(itemIndex) 'передать данные в книгу /лист
    End Select
    
    Call TrueBtnFocus(ActiveControl.name)
End Sub
'#########################################################
'ИНН / номер в ФИФ СИ / номер в ФИФ эталона  /должность сотрудника
Private Sub cmb3_Click()
    
    Select Case UMenu.typе
        
        Case organisationsOLD

            Application.ScreenUpdating = False
            
                ActiveCell = sArrDataBase(1, CInt(itemIndex)) ' ИНН
                ActiveCell.numberFormat = "0"
                
            Application.ScreenUpdating = True
            
        Case personsOLD
            DataTransfer sArrDataBase(1, CInt(itemIndex)), True, True
            
        Case archiveOLD
        ' ------------------------------------------------------
        'todo: привязать к db & config
            If Me.chbFullName Then ProtectSheets
            TrueMkDir myBase, WorkClsm, sArrDataBase 'создать каталог в архиве
            
        Case Else
        
            Dim sTempStr As String
            sTempStr = sArrDataBase(2, itemIndex)
            If myWdDoc And InStr(sTempStr, "ZZB") > 0 Then sTempStr = "рег. № " & sTempStr
            
            DataTransfer sTempStr, True
    End Select
    
    If UMenu.typе <> archiveOLD Then _
        TrueBtnFocus ActiveControl.name
End Sub
    Private Sub ProtectSheets()
        
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            
            Dim arrSearch(), i As Byte, rSearchCell As Range
            '###################################################################
            arrSearch = Array("ОКОК(", _
                                "РАСПАД(", _
                                "ВЗВЕШСР(", _
                                "ДАТАПРИВЕД(", _
                                "ОСКОВЗВ(", _
                                "ОСКОСА(", _
                                "ВЗВЕШНЕОПР(") 'ячейки для поиска значений
            '###################################################################
            For i = LBound(arrSearch) To UBound(arrSearch) 'для всех поисковых значений
                
                Set rSearchCell = ws.Cells.Find(arrSearch(i))
                Do While Not rSearchCell Is Nothing
                    rSearchCell.value = rSearchCell.value
                    Set rSearchCell = ws.Cells.Find(arrSearch(i))
                Loop
            Next i
            
    '        ws.[l:s].EntireColumn.Hidden = True
            ws.EnableSelection = xlNoSelection
            
            Dim password As String
            password = NewPassword

            ws.Protect password, _
                UserInterfaceOnly:=True

            Debug.Print "Sheet «" & ws.name & "» protected, password = " & password
        Next
        
        Dim protectedDirectoryPath As String
        protectedDirectoryPath = ActiveWorkbook.path & "\somnium\"
        
        If Dir(protectedDirectoryPath, vbDirectory) = vbNullString Then _
            MkDir protectedDirectoryPath
            
        ActiveWorkbook.SaveAs protectedDirectoryPath & _
            GetFileNameWithOutExt(ActiveWorkbook.name) & "." & GetExt(ActiveWorkbook.name)
            
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End Sub
    Private Function NewPassword() As String
        
        Dim i As Byte, _
            index_second As Integer, _
            pass As String
            
        Do While Len(pass) < 10
        
            Randomize
            index_second = CInt(Rnd * UBound(sArrKeyCode, 2) - 1)
            If index_second = -1 Then index_second = 0
            
            pass = pass + sArrKeyCode(1, index_second)
            If Len(pass) < 10 And index_second < UBound(sArrKeyCode, 2) / 2 Then _
                pass = pass & Left(index_second, 1)
            
            i = i + 1
        Loop
        
        NewPassword = pass
    End Function

'#########################################################
'адрес организации / методика поверки СИ / доп. сведения эталона
Private Sub cmb4_Click()

    If UMenu.typе <> personsOLD Then
        DataTransfer sArrDataBase(3, CInt(itemIndex)), True
    Else
        DataTransfer "Поверитель", True, True
    End If
    
    TrueBtnFocus ActiveControl.name
End Sub
'#########################################################
Private Sub cmbReady_Click() 'Готово
    VBA.Unload Me
End Sub
Private Sub cmbReady_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Shift = 1 Or Shift = 2 Then 'зажат ctrl или shift
    
        If UMenu.typе = organisationsOLD Then
            

            DataBase.Refactor _
                oldDataBaseArray:=sArrDataBase
                
        ElseIf UMenu.typе = instrumentsOLD Then
        
            DataBase.Refactor _
                oldDataBaseArray:=sArrDataBase
                
        ElseIf UMenu.typе = personsOLD Then
        
            DataBase.Refactor _
                oldDataBaseArray:=sArrDataBase
            
        End If
        
        VBA.Unload Me
        
    End If

End Sub
'#########################################################
'процедура открывает каталог с файлами выбранного пункта шаблонов
Private Sub btnOpenFolder_Click()

    If UMenu.typе = archiveOLD Then _
        Explorer.OpenFolder Config.archiveLocalPath, True: _
        Exit Sub
        
    If UMenu.typе = organisationsOLD Then

        
        Dim itemKey As String
        itemKey = sArrDataBase(LBound(sArrDataBase) + 2, CInt(itemIndex))
        
        If DataBase.IsDirAvailable(itemKey) Then _
            DataBase.TargetItemKey = itemKey
            DataBase.OpenItemDir
        
        VBA.Unload Me
        Exit Sub
        
    End If
    
    Dim sTempDir As String, _
        templatesDir As String, _
        fifRegNum As String, _
        sTypeName As String
        
    templatesDir = Config.instrumentsPath & "\"
    fifRegNum = sArrDataBase(2, itemIndex)
        
    Dim currTemplateDir As String
    currTemplateDir = Dir(templatesDir & "*" & fifRegNum & "*", vbDirectory)   'каталог выбранного СИ
    
    Dim targetPath As String
    targetPath = fso.BuildPath(templatesDir, currTemplateDir)
    
    Dim refFifPath As String
    refFifPath = templatesDir & currTemplateDir & Application.PathSeparator & REFERENCE_FIFNUM_FILENAME
    
    Dim refRegNumber As String
    
    If FileExist(refFifPath) Then
        
        refRegNumber = Base.ContentFromFile(refFifPath): _
        refRegNumber = Replace(refRegNumber, vbNewLine, vbNullString)
    
        If refRegNumber <> fifRegNum Then
                
            Dim newTargetDir As String
            newTargetDir = Dir(templatesDir & "*" & refRegNumber & "*", vbDirectory)  'каталог СИ
            
            targetPath = fso.BuildPath(templatesDir, newTargetDir)
            
        End If
        
    End If
    
    Explorer.OpenFolder targetPath, True
    VBA.Unload Me
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 13:31
' Назначение:
' ----------------------------------------------------------------
Private Sub DisableButtons()
    
    Me.cmb1.Enabled = False
    Me.cmb1.BackColor = Colors.white
    
    Me.cmb2.Enabled = False
    Me.cmb4.Enabled = False
    
    Me.cmbProtSv.Enabled = False
    Me.cmbProtSv.BackColor = Colors.white
    
    Me.cmbDescription.Enabled = False
    Me.cmbDescription.BackColor = Colors.white
    
    Me.cmbMetodic.Enabled = False
    Me.cmbMetodic.BackColor = Colors.white
    
    If UMenu.typе <> archiveOLD Then
        
        Me.btnOpenFolder.Enabled = False
        Me.btnOpenFolder.BackColor = Colors.white
        

        
    End If
'
'    If UMenu.typе = instrumentsOLD And myWdDoc = False Then _
'        Me.cmbProtSv.Caption = "Шаблон" & vbNewLine & "протокола"
    
    If UMenu.typе <> personsOLD Then
    
        If UMenu.typе <> archiveOLD Then _
            Me.cmb3.Enabled = False
            
        If myWdDoc = False And UMenu.typе = etalonsOLD Then _
            Me.cmbProtSv.Enabled = True
        
    End If
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 15:05
' Назначение: загрузка шаблона
' ----------------------------------------------------------------
Private Sub cmbProtSv_Click()
        
    Dim bolStopUnload As Boolean
    
    If UMenu.typе = instrumentsOLD Then 'средства измерений
    
        If myWdDoc Then
            cmbOneClick
            
        Else 'только для xl
            If myMi.sName = vbNullString Then _
                GetDBindAndMyMi listResults.text
            ' ------------------------------------------------------
            'todo: привязать к db, config
            If SaveAsTemplate(WorkClsm, myMi) = False Then _
                bolStopUnload = True
        End If
        
    ElseIf UMenu.typе = etalonsOLD Then 'эталоны
    
        If myWdDoc Then
            cmb3_Click
            cmb4_Click
        Else
            PasteEtalons sArrDataBase
        End If
    End If
    
    If bolStopUnload = False Then _
        VBA.Unload Me
End Sub
'#########################################################
'процедура загружает описание типа
Private Sub cmbDescription_Click()

    OpenPDF Config.sourceDataPath, myMi.sFif, "ot"
    VBA.Unload Me
    
End Sub
'#########################################################
'процедура загружает методику поверки
Private Sub cmbMetodic_Click()

    OpenPDF Config.sourceDataPath, myMi.sFif, "mp"
    VBA.Unload Me
    
End Sub

'#########################################################
'загрузить конфигурацию из файлов
Private Sub GetMyConfigFromFile(ByRef myBase As Z_clsmBase, ByRef WorkClsm As Z_clsmSearch)
    
    With myBase 'работа с классом настроек
        .AddP "startDir", "workDir"
        .AddP "depPrefix", "labNum"
        .AddP "templatesDir"
        
        Select Case True 'параметр загрузки элементов формы
        
            Case UMenu.typе = organisationsOLD
                .AddP "cusDB" 'заказчики
                
            Case UMenu.typе = instrumentsOLD
                .AddP "measInstrDB" 'средства измерений
                
            Case UMenu.typе = etalonsOLD
                .AddP "etalDB" 'эталоны
                
            Case UMenu.typе = personsOLD
                .AddP "empDB", "isFullName" 'фамилии
                
            Case UMenu.typе = archiveOLD
                .AddP "cusDB", "useArchiveDir", "ArchivePath" 'Архив ПКР
                
        End Select
        
        .GetArrFF setDir, Environ("USERNAME") & ".uCfg" 'загрузить в класс конфигурацию
        .FillValues 'обязательно: найти значения выходных параметров по ключам

        WorkClsm.FillConfiguration .Parameters, .values 'передать конфигурацию в класс
        .ClearParameters
        
        .AddP "constrSrch"
        Select Case True 'параметр загрузки элементов формы
            
            Case UMenu.typе = organisationsOLD
                .AddP "custSearch", "custIndex"  'заказчики
                
            Case UMenu.typе = instrumentsOLD
                .AddP "instrSearch", "instrIndex", "normalCondition" 'средства измерений
                
            Case UMenu.typе = etalonsOLD
                .AddP "etalSearch", "etalIndex" 'эталоны
                
            Case UMenu.typе = personsOLD
                .AddP "empSearch", "empIndex": AddInvertParameter myBase 'фамилии
                
            Case UMenu.typе = archiveOLD
                .AddP "archveSearch", "archiveIndex" 'Архив ПКР
        End Select
        
        .GetArrFF setDir, "settings.ini" 'загрузить в класс конфигурацию
        .FillValues 'обязательно: найти значения выходных параметров по ключам
        
        WorkClsm.FillSettings .Parameters, .values 'передать конфигурацию в класс
    End With
End Sub
'#########################################################
'процедура заполняет массив кейкодов
Private Sub FillArrKeycodeFromFile()
    
    Const fileName As String = "keycode.npDb"
     
    Dim charTablePath As String
    charTablePath = fso.BuildPath(Config.sourceDataPath, fileName)
    
    If fso.FileExists(charTablePath) Then 'если опознано наличие в каталоге надстройки файла кейкодов
        
        sArrKeyCode = WorkClsm.FillDataBase( _
            myBase.GetArrFF(charTablePath))  'получить массив кейкодов(если файл обнаружен)

    Else 'если файл не был обнаружен и загружен
        ReDim sArrKeyCode(0, 1)
    End If
    
End Sub
'#########################################################
'процедура заполняет массив базы данных
Private Sub FillArrDataBaseFromFile( _
    ByRef sArrDataName() As String, _
    sDbName As String _
    )
    ' ------------------------------------------------------
    'todo: привязать к db -- загрузка базы данных из файла

    If FileExist(WorkClsm.startDir, sDbName) Then 'если база данных обнаружена по указанному пути
        sArrDataName = WorkClsm.FillDataBase(myBase.GetArrFF(WorkClsm.startDir, sDbName), True)  'преобразовать массив файла в массив базы данных
        
        If UMenu.typе <> archiveOLD Then _
            Me.LabelInfo.caption = "Файл данных «" & sDbName & "», " & RusPadejPozition(UBound(sArrDataBase, 2) + 1)
        
    Else 'файл БД не был загружен
        ReDim sArrDataName(0)
        With Me.LabelInfo: .caption = "Файл данных не загружен.": .foreColor = &H80&: End With 'красный цвет
    End If
    
End Sub
'#########################################################
'процедура воспроизводит последний результат поиска значений в текстбоксе и выбранный элемент
Private Sub UpdateTboxAndListbox()
    
    If sArrDataBase(LBound(sArrDataBase), UBound(sArrDataBase, 2)) <> vbNullString Then 'если массив файла был получен
    
        If UMenu.typе = etalonsOLD And myWdDoc = False Then _
            Me.cmbProtSv.Enabled = True
            
  
' ------------------------------------------------------
'todo: привязать к Cache
        If WorkClsm.LastSearch <> "недоступно" Then _
            Me.tboxSearch = WorkClsm.LastSearch: 'последний поисковый запрос
' ------------------------------------------------------
'todo: привязать к Cache
        If WorkClsm.LastIndex <> "недоступно" Then 'последний индекс позиции выделения элемента
            If Me.listResults.ListCount = 0 Then UpdateListResults sArrDataBase
            

        ' ------------------------------------------------------
'todo: привязать к Cache
            
            If WorkClsm.LastIndex >= 0 Then
            
                Dim bolSelind As Boolean
                If WorkClsm.LastIndex < Me.listResults.ListCount - 1 Then _
                    bolSelind = True
                
                If bolSelind Then
                    
                    Me.listResults.Selected(CInt(WorkClsm.LastIndex)) = True 'выделить ранее выбранный элемент
                    TrueElementForeColor Me.labelUnderSearchField 'окрасить метку
                    Me.labelUnderSearchField.caption = "нажмите Enter" 'информационная подпись
                    
                End If
                
            End If
        End If

      
    End If
End Sub

'#########################################################
'процедура обновляет содержимое листбокса, исходя из представленного массива данных
Private Sub UpdateListResults(sArrDBase() As String, _
    Optional strException As String = "nodata") 'обновить содержимое листбокса1
    
    Dim i As Integer, sTempStr As String
    Me.listResults.Clear
    ' ------------------------------------------------------
    Dim lowerBound As Integer
    lowerBound = LBound(sArrDBase, 2)
    
    If UMenu.typе = archiveOLD Then _
        lowerBound = UBound(sArrDBase, 2)
    ' ------------------------------------------------------
    Dim upperBound As Integer
    upperBound = UBound(sArrDBase, 2)
    
    If UMenu.typе = archiveOLD Then _
        upperBound = LBound(sArrDBase, 2)
    ' ------------------------------------------------------
    Dim stepData As Integer
    stepData = 1
    
    If UMenu.typе = archiveOLD Then _
        stepData = -1
    ' ------------------------------------------------------
    If sArrDataBase(LBound(sArrDataBase), UBound(sArrDataBase, 2)) <> vbNullString Then 'массив был получен
        
        For i = lowerBound To upperBound Step stepData 'для каждого блока
            sTempStr = vbNullString 'очищение переменной
            
            Select Case UMenu.typе
                ' ------------------------------------------------------
                Case instrumentsOLD
                    sTempStr = GetTempStrBy2(sArrDBase, i, strException)  'средства измерений
                ' ------------------------------------------------------
                Case Else
                
                    If UMenu.typе = etalonsOLD And sArrDBase(LBound(sArrDBase), i) <> "" Then _
                        sTempStr = sTempStr & sArrDBase(UBound(sArrDBase), i) & " \\ " 'ключевое слово
                    
                    If UMenu.typе <> archiveOLD Then _
                        sTempStr = sTempStr & GetTempStrByOther(sArrDBase, i, strException) 'другие БД
                    
                    If UMenu.typе = archiveOLD Then _
                        sTempStr = sArrDBase(0, i)
                        
            End Select
            
            If sTempStr <> "" Then _
                Me.listResults.AddItem sTempStr

        Next i
        
        'If myMen u.type = 14 And sArrDBase(LBound(sArrDBase), UBound(sArrDBase, 2)) = "" Then Me.listResults.AddItem " "
    End If
    
    If UMenu.typе <> archiveOLD Then
    
        Me.listResults.AddItem "внести изменения в базу данных..."
        ' ------------------------------------------------------
        'todo: привязать к db cache
        If UMenu.isLoaded = True Then _
            WorkClsm.LastIndex = "недоступно"  'очищать переменную последнего поискового индекса после обновления списка
            
    End If
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 12:54
' Назначение: функция передаёт строку для добавления в список результатов для формы средств измерений
'    параметр sArrDBase:
'    параметр i:
'    параметр strException:
' Возвращаемый тип: String
' ----------------------------------------------------------------
Function GetTempStrBy2( _
    sArrDBase() As String, _
    i As Integer, _
    strException As String _
    ) As String
    
    Dim sTempStr As String
    
    If sArrDBase(LBound(sArrDBase), i) <> vbNullString Then 'чтобы не было пустой строки с данными
        
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 2, i) & " \\ " 'номер в фиф
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 1, i) & " \\ "
        
'        If sArrDBase(UBound(sArrDBase) - 1, i) <> "nodata" Then _
'            sTempStr = sTempStr & " \\ " & sTempStr & sArrDBase(UBound(sArrDBase) - 1, i) & " \\ " 'Си является эталоном
        
        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase), i) & " \\ " 'полное наименование
        
'        If sArrDBase(UBound(sArrDBase), i) <> "nodata" Then _
'            sTempStr = sTempStr & sArrDBase(UBound(sArrDBase), i) & " \\ " 'наличие шаблона для данного СИ
        
'        If sArrDBase(LBound(sArrDBase) + 1, i) <> strException And sArrDBase(LBound(sArrDBase) + 1, i) <> vbNullString Then _
'            sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 1, i) & " \\ " 'тип СИ
            

'        sTempStr = sTempStr & sArrDBase(LBound(sArrDBase) + 3, i) 'методика поверки
               

            
    End If
    
    GetTempStrBy2 = sTempStr
    
End Function
'#########################################################
'функция передаёт строку для добавления в список результатов для форм, кроме средств измерений
Function GetTempStrByOther(sArrDBase() As String, i As Integer, strException As String) As String
    Dim sTempStr As String, j As Byte, sArrTemp() As String

    For j = LBound(sArrDBase) To UBound(sArrDBase) 'передать значения из подстрок
        
        If UMenu.typе = etalonsOLD And j = UBound(sArrDBase) Then Exit For
        If UMenu.typе = organisationsOLD And j = 2 Then   'не вносить "ключевое слово" организации заказчика в результаты поиска
        
        ElseIf UMenu.typе = personsOLD And j = LBound(sArrDBase) Then 'передать ФИО
        
            If sArrDBase(j, i) <> strException And sArrDBase(j, i) <> "" Then
                If Me.chbFullName = False Then
                    sArrTemp = Split(sArrDBase(j, i), " "): sTempStr = sArrTemp(0) & " " 'разбить строку на пробелах
                    
                    Dim K As Byte
                    For K = LBound(sArrTemp) + 1 To UBound(sArrTemp)
                        sTempStr = sTempStr & Left(sArrTemp(K), 1) & "."
                    Next
                Else
                    sTempStr = sArrDBase(j, i) 'переменная для добавления в листбокс
                End If
            End If

        Else
            If j > LBound(sArrDBase) And sArrDBase(j, i) <> strException And sArrDBase(j, i) <> "" Then
                sTempStr = sTempStr & " \\ " 'разделитель
                If UMenu.typе = organisationsOLD And j = 1 Then sTempStr = sTempStr & "ИНН "
            End If
            If sArrDBase(j, i) <> strException And sArrDBase(j, i) <> "" Then _
                sTempStr = sTempStr & sArrDBase(j, i) 'переменная для добавления в листбокс
        End If
    Next j
    
    GetTempStrByOther = sTempStr
End Function

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
