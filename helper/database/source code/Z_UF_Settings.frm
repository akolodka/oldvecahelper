VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_UF_Settings 
   Caption         =   "Настройки работы с источниками данных"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   OleObjectBlob   =   "Z_UF_Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Z_UF_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'запрет на использование неявных переменных

Const bolFullFirstNames = True 'полная форма записи имён сотрудников

Private myBase As New Z_clsmBase, WorkClsm As New Z_clsmSettings, sArrDataBase() As String

Private Sub cmbChooseTemplatesDir_Click()
    Dim sMyPath As String, _
        sTempName As String
    
    sMyPath = GetFolderFPath(, , False) 'выбрать локальую рабочую директорию
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\"
        WorkClsm.templatesDir = sMyPath 'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub

Private Sub cmbOpenTemplatesDir_Click()
    Explorer.OpenFolder WorkClsm.templatesDir, True  'открыть и передать фокус
End Sub

'#######################################################
'переустановить программу
Private Sub cmbUninstall_Click()

    Dim sFileName As String
    sFileName = WorkClsm.startDir & "Installer\Uninstall.vbs": If FileExist(sFileName) Then Shell "explorer.exe " & sFileName
    VBA.Unload Me
End Sub
'#######################################################
'изменение вкладки настроек
Private Sub MultiPage1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        UMenu.Unload
        
End Sub
'#######################################################
'первичная загрузка формы
Private Sub UserForm_Initialize()
    ReDim sArrDataBase(0)
    
    TransformConfigFile 'для перехода от старого обновления к новому
    
    If FileExist(setDir, Environ("USERNAME") & ".uCfg") Then
        Me.cmbCfgExp.Enabled = True: Me.cmbCfgImp.BackColor = &HFFFFFF    'белый
    End If
    
    With myBase 'работа с классом настроек
        '#######################################################
        .GetArrFF setDir, Environ("USERNAME") & ".uCfg" 'загрузить в класс файл локальной конфигурации
        '#######################################################

        .AddP "startDir", "cusDB", "measInstrDB" 'входные параметры - ключи
        .AddP "etalDB", "empDB", "isFullName"
        
        .AddP "workDir", "depPrefix", "labNum"
        .AddP "headNAME", "verNAME", "isPortable"
        .AddP "empNAME", "empSTATE", "helpOption"
        .AddP "empSecNAME", "empSecSTATE", "templatesDir"
        
        .AddP "xlPrPath", "xlPrcPath"
        .AddP "wdSvPath", "wdSrtPath", "wdInPath"
        .AddP "useArchiveDir", "ArchivePath"
        
        .FillValues 'обязательно: найти значения выходных параметров по ключам
    End With
              
    WorkClsm.FillProperties myBase.Parameters, myBase.values 'передать извлечённые из настроек параметры в специальный класс
    UpdateDBLabels  'обновить метки выбранных файлов согласно загруженным данным
    
'    With WorkClsm
'        If .empState = "недоступно" Then _
'            Me.cmbxEmployee = "<выбрать>": PreSaveSettings 'для обратной совместимости
'
'        If .headSTATE = "недоступно" Then _
'            Me.cmbxHead = "<выбрать>": PreSaveSettings 'для обратной совместимости
'    End With
    
  '  Me.VersionLabel.Caption = "v " & GetCDProp("Version") & " от " & _
        Format(GetCDProp("VersionDate"), "[$-FC19]dd mmmm yyyy г.") & vbNewLine & "kolaa@vniim.ru; т/н 21-09"
        
    Me.VersionLabel.caption = Addin.VersionCaption
    
    SetEventControls Me
    
    WorkClsm.isFullName = bolFullFirstNames 'форма записи имени
    bolAlreadySaved = True: bolUF_Set_Load = True
End Sub
    Private Function VersionCaption( _
        ) As String
        
        Dim strData As String
        strData = "Версия " & Properties.Version & vbNewLine & _
            Format(Properties.Comments, "dd mmmm yyyy г.")
        
        VersionCaption = strData
        
    End Function
    
    Private Sub TransformConfigFile( _
        )
        
        Dim oldConfigName As String
        oldConfigName = "localConfig.uCfg"
        
        Dim newConfigName As String
        newConfigName = Environ("USERNAME") & ".uCfg"
        
        Dim configDir As String
        configDir = Environ("APPDATA") & "\Microsoft\Помощник ПКР\"
            
        If Dir(configDir & newConfigName) <> vbNullString Then _
            Exit Sub
        
        If Dir(configDir & oldConfigName) <> vbNullString Then
        
            FileCopy configDir & oldConfigName, _
                     configDir & newConfigName
        End If
        
    End Sub

'#######################################################
'выгрузка формы из памяти
Private Sub UserForm_Terminate()
    Set myBase = Nothing: Set WorkClsm = Nothing 'обязательно: очистка объектов
    bolUF_Set_Load = False: bolAlreadySaved = False
End Sub
'#######################################################
'процедура обновляет массив данных об исполнителях
Sub UpdateEmployees(bolFormIsLoad As Boolean)
    Dim i As Byte, sEmpName As String, sHeadName As String, sVerName As String, sempSecName As String, _
        sArrTemp() As String, sTempStr As String, sTempFull As String
    
    With WorkClsm
        If FileExist(.startDir, .empDB) Then 'если база данных фамилий обнаружена по указанному пути
            If UBound(sArrDataBase) = 0 Then sArrDataBase = .FillDataBase(myBase.GetArrFF(.startDir, .empDB), True)  'преобразовать массив файла в массив базы данных
            If UBound(sArrDataBase) > 0 Then 'если массив файла был получен
            
                Me.labelEmployee.Enabled = True: Me.labelHead.Enabled = True
                Me.labelVerifier.Enabled = True: Me.labelEmpSecond.Enabled = True
                
                Me.cmbxHead.Enabled = True: Me.cmbxEmployee.Enabled = True
                Me.cmbxVerifier.Enabled = True: Me.cmbxEmpSecond.Enabled = True
                
                If bolFormIsLoad Then Exit Sub
                
                Me.cmbxHead.Clear: Me.cmbxVerifier.Clear:  Me.cmbxEmpSecond.Clear: Me.cmbxEmployee.Clear
                Me.cmbxHead.AddItem "<выбрать>": Me.cmbxVerifier.AddItem "<выбрать>"
                Me.cmbxEmpSecond.AddItem "<выбрать>": Me.cmbxEmployee.AddItem "<выбрать>"
                
                sHeadName = .headNAME: sVerName = .verNAME: sempSecName = .empSecName: sEmpName = .empName 'получить данные из класса
                
                If sHeadName = "" Then .headNAME = "недоступно": sHeadName = "недоступно"
                If sVerName = "" Then .verNAME = "недоступно": sVerName = "недоступно"
                If sempSecName = "" Then .empSecName = "недоступно": sempSecName = "недоступно"
                If sEmpName = "" Then .empName = "недоступно": sEmpName = "недоступно"

                If InStr(sHeadName, " ") > 0 Then sHeadName = Left(sHeadName, InStr(sHeadName, " ")) 'получить чистую фамилию
                If InStr(sVerName, " ") > 0 Then sVerName = Left(sVerName, InStr(sVerName, " "))
                If InStr(sempSecName, " ") > 0 Then sempSecName = Left(sempSecName, InStr(sempSecName, " "))
                If InStr(sEmpName, " ") > 0 Then sEmpName = Left(sEmpName, InStr(sEmpName, " "))
                
                For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2) 'передать в комбобоксы сведения о сотрудниках
 
                    sArrTemp = Split(sArrDataBase(0, i), " "): sTempStr = sArrTemp(0) & " " 'разбить строку на пробелах
                    
                    Dim K As Byte
                    For K = LBound(sArrTemp) + 1 To UBound(sArrTemp)
                        sTempStr = sTempStr & Left(sArrTemp(K), 1) & "." 'добавить точки как разделители
                    Next
                        
                    sTempFull = sArrDataBase(0, i) 'полная форма записи имени
                    If bolFullFirstNames = False Then sTempFull = sTempStr  'если не выбрана опция полной формы записи имени
                    
                    If InStr(sArrDataBase(1, i), "уководит") > 0 Then
                        Me.cmbxHead.AddItem sTempFull
                        
                        If InStr(Me.cmbxHead.List(Me.cmbxHead.ListCount - 1), sHeadName) > 0 Then _
                            Me.cmbxHead = Me.cmbxHead.List(Me.cmbxHead.ListCount - 1)
                            
                    End If
                        
                    If InStr(sArrDataBase(2, i), "поверитель") > 0 Then
                        Me.cmbxVerifier.AddItem sTempFull
                        If InStr(Me.cmbxVerifier.List(Me.cmbxVerifier.ListCount - 1), sVerName) > 0 Then _
                            Me.cmbxVerifier = Me.cmbxVerifier.List(Me.cmbxVerifier.ListCount - 1)
                    End If
                        
                    Me.cmbxEmployee.AddItem sTempStr
                    If InStr(Me.cmbxEmployee.List(Me.cmbxEmployee.ListCount - 1), sEmpName) > 0 _
                        Then Me.cmbxEmployee = Me.cmbxEmployee.List(Me.cmbxEmployee.ListCount - 1)
                        
                    Me.cmbxEmpSecond.AddItem sTempStr
                    If InStr(Me.cmbxEmpSecond.List(Me.cmbxEmpSecond.ListCount - 1), sempSecName) > 0 Then _
                        Me.cmbxEmpSecond = Me.cmbxEmpSecond.List(Me.cmbxEmpSecond.ListCount - 1)
                Next i
                
                With Me.cmbxHead
                   If .text = "недоступно" Or .text = "" Then .text = "<выбрать>"
                End With
                
                With Me.cmbxVerifier
                   If .text = "недоступно" Or .text = "" Then .text = "<выбрать>"
                End With
                
                With Me.cmbxEmpSecond
                   If .text = "недоступно" Or .text = "" Then .text = "<выбрать>"
                End With
                
                With Me.cmbxEmployee
                   If .text = "недоступно" Or .text = "" Then .text = "<выбрать>"
                End With
                
                Exit Sub
            End If
        End If
        .empName = "недоступно": .headNAME = "недоступно": .verNAME = "недоступно": .empSecName = "недоступно"
    End With
End Sub
'#######################################################
'процедура возвращает исходный цвет и доступность всех контролов, которые изменяются при загрузке
Sub AllControlsToZeroPos()
    bolUF_Set_Load = False
    
    Me.cmbOpenFolder.Enabled = False 'открыть стартовую директорию
    Me.cmbChoseBaseDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    Me.cmbCreateCus1.Enabled = False: Me.cmbChoseCus2.Enabled = False
    Me.cmbCreateMi3.Enabled = False: Me.cmbChoseMi4.Enabled = False
    Me.cmbCreateEt5.Enabled = False: Me.cmbChoseEt6.Enabled = False
    Me.cmbCreateLn7.Enabled = False: Me.cmbChoseLn8.Enabled = False
    
    Me.labelEmployee.Enabled = False: Me.cmbxEmployee.Enabled = False
    Me.labelHead.Enabled = False: Me.cmbxHead.Enabled = False
    Me.labelVerifier.Enabled = False: Me.cmbxVerifier.Enabled = False
    Me.labelEmpSecond.Enabled = False: Me.cmbxEmpSecond.Enabled = False
    
    Me.chbUseArchiveDir = False
'    Me.chbHelp = False
  '  Me.chbPortable = False
    
    Me.cmbOpenPrDir.Enabled = False
    Me.cmbChoosePrDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    cmbOpenPrcDir.Enabled = False
    Me.cmbChoosePrcDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    cmbOpenSvDir.Enabled = False
    Me.cmbChooseSvDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    cmbOpenSrtDir.Enabled = False
    Me.cmbChooseSrtDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    cmbOpenInDir.Enabled = False
    Me.cmbChooseInDir.BackColor = &HC0FFFF    'желтый цвет - исходный
    
    bolUF_Set_Load = True
End Sub
'#######################################################
'отображение полного имени сотрудников
Private Sub chbUseArchiveDir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub
Private Sub chbUseArchiveDir_Change()
    TrueElementForeColor Me.chbUseArchiveDir, Not Me.chbUseArchiveDir, 1: If bolUF_Set_Load = False Then Exit Sub
    
    WorkClsm.ArchivePath = "недоступно" 'передать папаметр в рабочий класс
    WorkClsm.useArchiveDir = False
    
    If Me.chbUseArchiveDir Then
    
        Dim sMyPath As String
        sMyPath = GetFolderFPath("Выбор директории архива") 'выбрать путь к каталогу
        
        If sMyPath <> "NoPath" Then
            sMyPath = sMyPath & "\"
            
            WorkClsm.ArchivePath = sMyPath 'передать папаметр в рабочий класс
            WorkClsm.useArchiveDir = Me.chbUseArchiveDir
        End If
        If sMyPath = "NoPath" Then Me.chbUseArchiveDir = False
    End If
    
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
'#######################################################
'процедура обновляет информацию на метках выбранных файлов
Sub UpdateDBLabels()
    
    If bolUF_Set_Load Then AllControlsToZeroPos 'вернуть свойства контролов в нулевое положение
    
    Dim sMyDir As String
    
    With WorkClsm
        sMyDir = .startDir 'начальная директория файлов, где хранятся базы данных
        If sMyDir <> "недоступно" Then  'если директория была выбрана
        
            With Me.baseDirLabel
                PaintLabels Me.baseDirLabel.name, sMyDir
                
                If FolderNotExist(sMyDir) = False Then
                    Me.cmbOpenFolder.Enabled = True
                    Me.cmbChoseBaseDir.BackColor = &HFFFFFF 'white color
                
                    Me.cmbCreateCus1.Enabled = True: Me.cmbChoseCus2.Enabled = True
                    Me.cmbCreateMi3.Enabled = True: Me.cmbChoseMi4.Enabled = True
                    Me.cmbCreateEt5.Enabled = True: Me.cmbChoseEt6.Enabled = True
                    Me.cmbCreateLn7.Enabled = True: Me.cmbChoseLn8.Enabled = True
                End If
                
                PaintLabels Me.CusLabel.name, sMyDir, WorkClsm.cusDB
                PaintLabels Me.InstrLabel.name, sMyDir, WorkClsm.measInstrDB
                PaintLabels Me.EtalLabel.name, sMyDir, WorkClsm.etalDB
                PaintLabels Me.empLabel.name, sMyDir, WorkClsm.empDB
            End With
            
        End If
'вторая страница настроек
        
        sMyDir = .workDir 'каталог, в который копируются шаблоны с сервера
        
        If sMyDir = "недоступно" Then _
            .workDir = "C:\Users\" & Environ("USERNAME") & "\Desktop\": sMyDir = .workDir

        With Me.labelWorkDir
        
            PaintLabels .name, sMyDir
            If FolderNotExist(sMyDir) = False Then  'если указанный каталог доступен
                Me.cmbOpenWorkDir.Enabled = True
                Me.cmbChoseWorkDir.BackColor = &HFFFFFF 'white color
            End If
            
        End With
        
        '----------------------
        sMyDir = .templatesDir
        
        With Me.labelTemplatesDir
        
            PaintLabels .name, sMyDir
            
            If Not FolderNotExist(sMyDir) Then  'если указанный каталог доступен
            
                Me.cmbOpenTemplatesDir.Enabled = True
                Me.cmbChooseTemplatesDir.BackColor = &HFFFFFF 'white color
                
            End If
            
        End With

        If .depPrefix <> "недоступно" Then Me.tboxDepPrefix = .depPrefix
        If .labNum <> "недоступно" Then Me.tBoxLabNum = .labNum
        
        If bolUF_Set_Load Then
            bolUF_Set_Load = False
            
            If .useArchiveDir <> "недоступно" Then _
                If Dir(.ArchivePath, vbDirectory) <> "" Then Me.chbUseArchiveDir = .useArchiveDir
            bolUF_Set_Load = True
        Else
            If .useArchiveDir <> "недоступно" Then _
                If Dir(.ArchivePath, vbDirectory) <> "" Then Me.chbUseArchiveDir = .useArchiveDir
        End If
            
        If .isPortable = "недоступно" Then .isPortable = "False"
       ' Me.chbPortable = .isPortable
        
        If .helpOption = "недоступно" Then .helpOption = "True"
       ' Me.chbHelp = .helpOption
        
        UpdateEmployees (bolUF_Set_Load) 'обновить содержимое комбобокса исполнителя
        
'третья страница настроек
        If .xlPrPath = "недоступно" Then .xlPrPath = sMyDir
        PaintLabels Me.labelPrDir.name, .xlPrPath
        
        If FolderNotExist(.xlPrPath) = False Then  'если указанный каталог доступен
            Me.labelPrDir.foreColor = &H8000& 'зелёный цвет
            Me.cmbChoosePrDir.BackColor = &HFFFFFF 'white color
            Me.cmbOpenPrDir.Enabled = True
        End If

        If .xlPrcPath = "недоступно" Then .xlPrcPath = sMyDir
        PaintLabels Me.labelPrcDir.name, .xlPrcPath
        
        If FolderNotExist(.xlPrcPath) = False Then  'если указанный каталог доступен
            Me.labelPrcDir.foreColor = &H8000& 'зелёный цвет
            Me.cmbChoosePrcDir.BackColor = &HFFFFFF 'white color
            Me.cmbOpenPrcDir.Enabled = True
        End If
        
        If .wdSvPath = "недоступно" Then .wdSvPath = sMyDir
        PaintLabels Me.labelSvDir.name, .wdSvPath
            
        If FolderNotExist(.wdSvPath) = False Then  'если указанный каталог доступен
            Me.labelSvDir.foreColor = &HFF0000    'синий цвет
            Me.cmbChooseSvDir.BackColor = &HFFFFFF 'white color
            Me.cmbOpenSvDir.Enabled = True
        End If
        
        If .wdSrtPath = "недоступно" Then .wdSrtPath = sMyDir
        PaintLabels Me.labelSrtDir.name, .wdSrtPath
            
        If FolderNotExist(.wdSrtPath) = False Then  'если указанный каталог доступен
            Me.labelSrtDir.foreColor = &HFF0000    'синий цвет
            Me.cmbChooseSrtDir.BackColor = &HFFFFFF 'white color
            Me.cmbOpenSrtDir.Enabled = True
        End If
        
        If .wdInPath = "недоступно" Then .wdInPath = sMyDir
        PaintLabels Me.labelInDir.name, .wdInPath
            
        If FolderNotExist(.wdInPath) = False Then  'если указанный каталог доступен
            Me.labelInDir.foreColor = &HFF0000    'синий цвет
            Me.cmbChooseInDir.BackColor = &HFFFFFF 'white color
            Me.cmbOpenInDir.Enabled = True
        End If
    End With
End Sub
'#######################################################
'процедура закрашивает поля с именами файлов
Function PaintLabels( _
    labelName As String, _
    sDBPath As String, _
    Optional sDbName As String _
    ) As Boolean
    
    With Me.Controls(labelName)
        .foreColor = &H80&       'красный
        
        If sDbName <> "" Then 'передаётся наименование файла
        
            .caption = ShortedString(sDBPath & sDbName, 48)
            If FileExist(sDBPath, sDbName) Then _
                .TextAlign = fmTextAlignLeft: .foreColor = &H808000: PaintLabels = True 'бирюзовый
        Else 'передаётся наименование каталога
        
            .caption = ShortedString(sDBPath, 43)
            If FolderNotExist(sDBPath) = False Then _
                .TextAlign = fmTextAlignLeft: .foreColor = &H808000: PaintLabels = True   'бирюзовый
        End If
    End With
End Function
'#######################################################
'задать условия для сохранения настроек
Sub PreSaveSettings()
    With Me.cmbSaveReady 'кнопка готово
        .Font.Size = 11: .caption = "Сохранить": .BackColor = &HC0FFFF 'желтый цвет
    End With
    bolAlreadySaved = False
End Sub
'#######################################################
'выбор рабочей директории
Private Sub cmbChoseWorkDir_Click()
    
    Dim sMyPath As String, _
        sTempName As String
    
    sMyPath = GetFolderFPath(, , False) 'выбрать локальую рабочую директорию
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\"
        WorkClsm.workDir = sMyPath 'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
    
End Sub
'#######################################################
'открыть рабочую директорию
Private Sub cmbOpenWorkDir_Click()
    Explorer.OpenFolder WorkClsm.workDir, True
End Sub
'#######################################################
'процедура позволяет выбрать начальный каталог
Private Sub cmbChoseBaseDir_Click()
    Dim sMyPath As String, sTempName As String
    sMyPath = GetFolderFPath(, , True) 'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\"
        WorkClsm.startDir = sMyPath 'передать папаметр в рабочий класс
        
        WorkClsm.cusDB = "недоступно"
        sTempName = Dir(sMyPath & "\*.cuDb") 'поиск базы данных заказчиков
        If sTempName <> "" Then
            WorkClsm.cusDB = sTempName
            sTempName = Dir
            If sTempName <> "" Then WorkClsm.cusDB = "недоступно"
        End If
        
        WorkClsm.measInstrDB = "недоступно"
        sTempName = Dir(sMyPath & "\*.miDb") 'поиск базы данных заказчиков
        If sTempName <> "" Then
            WorkClsm.measInstrDB = sTempName
            sTempName = Dir
            If sTempName <> "" Then WorkClsm.measInstrDB = "недоступно"
        End If
        
        WorkClsm.etalDB = "недоступно"
        sTempName = Dir(sMyPath & "\*.etDb") 'поиск базы данных заказчиков
        If sTempName <> "" Then
            WorkClsm.etalDB = sTempName
            sTempName = Dir
            If sTempName <> "" Then WorkClsm.etalDB = "недоступно"
        End If
        
        WorkClsm.empDB = "недоступно"
        sTempName = Dir(sMyPath & "\*.nmDb") 'поиск базы данных заказчиков
        If sTempName <> "" Then
            WorkClsm.empDB = sTempName
            sTempName = Dir
            If sTempName <> "" Then WorkClsm.empDB = "недоступно"
        End If
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub
'#######################################################
'процедура позволяет выбрать или создать чистый файл базы данных
Sub ChooseFile(controlName As String, Optional CreateNew As Boolean)
    Dim sTempPath As String
    
    If CreateNew Then 'создание нового файла
        sTempPath = GetSaveAsFname(myCmbIndex(controlName), WorkClsm.startDir) 'путь к файлу для сохранения
    Else 'выбор существующего файла
        sTempPath = GetFileFPath(myCmbIndex(controlName), WorkClsm.startDir)   'получить полный путь к файлу
    End If
    
    If sTempPath <> "NoPath" Then
        
        If CreateNew Then
            Open sTempPath For Output As #1
                Print #1, "newFile"
            Close
        End If
        
        UpdateProperties controlName, sTempPath
    End If
End Sub
'#######################################################
'процдура корректно помещает свойства в класс
Sub UpdateProperties(controlName As String, sPath As String)
    With WorkClsm
        Select Case myCmbIndex(controlName)
            Case 1, 2: .cusDB = TrueName(sPath)  'заказчики
            Case 3, 4: .measInstrDB = TrueName(sPath)  'средства измерений
            Case 5, 6: .etalDB = TrueName(sPath) 'эталоны
            Case 7, 8: .empDB = TrueName(sPath) 'фамилии
        End Select
    End With
    
    UpdateDBLabels
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
'#######################################################
'открыть директорию с файлами баз данных
Private Sub cmbOpenFolder_Click()
    Explorer.OpenFolder WorkClsm.startDir, True  'открыть и передать фокус
End Sub
'#######################################################
'готово / сохранить
Private Sub cmbSaveReady_Click()
    With Me.cmbSaveReady
        
        If .caption = "Готово" Then
            VBA.Unload Me
        Else 'сохранить изменения
            '####################################
            With myBase
                .GetArrFF setDir, Environ("USERNAME") & ".uCfg" 'загрузить в класс файл локальной конфигурации
                .SaveProperties WorkClsm.myParameters, WorkClsm.myValues 'передать значения в настройки и сохранить в файле через класс
            End With
            '####################################
            .Font.Size = 12: .caption = "Готово": bolAlreadySaved = True: .BackColor = &HFFFFFF  'белый цвет
            
            Me.VersionLabel.caption = "Конфигурация сохранена."
            
            If Me.cmbCfgExp.Enabled = False Then _
                Me.cmbCfgExp.Enabled = True
        End If
    End With
End Sub
Private Sub cmbSaveReady_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Shift = 6 Then SaveNewVersion WorkClsm.startDir: VBA.Unload Me   'сратабывает при зажатых ctrl+alt
End Sub
'#######################################################
'функция возвращает индекс кнопки по её имени
Private Function myCmbIndex(ctrlName As String) As Byte
    myCmbIndex = Right(ctrlName, 1)
End Function
'#######################################################
'изменение руководителя по умолчанию
Private Sub cmbxHead_Change()
    With Me.cmbxHead
        If .text <> "<выбрать>" Then .BackColor = &H80000005 'белый
        
        If bolUF_Set_Load = False Then Exit Sub
        
        If .text = "<выбрать>" Then _
            WorkClsm.headNAME = "недоступно": .BackColor = &HC0FFFF 'желтый
        
        If .text <> "<выбрать>" Then
            WorkClsm.headNAME = .text
            
            Dim i As Byte, sTempStr As String
            sTempStr = .text: If InStr(sTempStr, " ") > 0 Then sTempStr = Left(sTempStr, InStr(sTempStr, " "))
            
            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                If InStr(sArrDataBase(0, i), sTempStr) > 0 Then WorkClsm.headSTATE = sArrDataBase(1, i): Exit For 'должность
            Next i
        End If
    End With
    
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
Private Sub cmbxHead_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then

    ElseIf KeyCode = 13 Then cmbSaveReady_Click
    Else: KeyCode = 0
    End If
End Sub
'#######################################################
'изменение поверителя по умолчанию
Private Sub cmbxVerifier_Change()
        
    With Me.cmbxVerifier
        If .text <> "<выбрать>" Then .BackColor = &H80000005 'белый
        
        If bolUF_Set_Load = False Then Exit Sub
        
        If .text = "<выбрать>" Then _
            WorkClsm.verNAME = "недоступно": .BackColor = &HC0FFFF 'желтый
        
        If .text <> "<выбрать>" Then WorkClsm.verNAME = .text
    End With
    
    If Me.cmbxEmployee = "<выбрать>" Then Me.cmbxEmployee = Me.cmbxVerifier
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
Private Sub cmbxVerifier_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then

    ElseIf KeyCode = 13 Then cmbSaveReady_Click
    Else: KeyCode = 0
    End If
End Sub
'#######################################################
'изменение исполнителя по умолчанию
Private Sub cmbxEmployee_Change()

    With Me.cmbxEmployee
        If .text <> "<выбрать>" Then .BackColor = &H80000005 'белый
        
        If bolUF_Set_Load = False Then Exit Sub
        
        If .text = "<выбрать>" Then _
            WorkClsm.empName = "недоступно": .BackColor = &HC0FFFF 'желтый
        
        If .text <> "<выбрать>" Then
            WorkClsm.empName = .text
            
            Dim i As Byte, sTempStr As String
            sTempStr = .text: If InStr(sTempStr, " ") > 0 Then sTempStr = Left(sTempStr, InStr(sTempStr, " "))
            
            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                If InStr(sArrDataBase(0, i), sTempStr) > 0 Then WorkClsm.empState = sArrDataBase(1, i): Exit For 'должность
            Next i
        End If
    End With
    
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
Private Sub cmbxEmployee_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then

    ElseIf KeyCode = 13 Then cmbSaveReady_Click
    Else: KeyCode = 0
    End If
End Sub
'#######################################################
'изменение дополнительного поверителя
Private Sub cmbxEmpSecond_change()

    With Me.cmbxEmpSecond
        If .text <> "<выбрать>" Then .BackColor = &H80000005 'белый
        
        If bolUF_Set_Load = False Then Exit Sub
        
        If .text = "<выбрать>" Then _
            WorkClsm.empSecName = "недоступно": .BackColor = &HC0FFFF 'желтый
        
        If .text <> "<выбрать>" Then
            WorkClsm.empSecName = .text
            
            Dim i As Byte, sTempStr As String
            sTempStr = .text: If InStr(sTempStr, " ") > 0 Then sTempStr = Left(sTempStr, InStr(sTempStr, " "))
            
            For i = LBound(sArrDataBase, 2) To UBound(sArrDataBase, 2)
                If InStr(sArrDataBase(0, i), sTempStr) > 0 Then WorkClsm.empSecState = sArrDataBase(1, i): Exit For 'должность
            Next i
        End If
    End With

    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
Private Sub cmbxEmpSecond_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then

    ElseIf KeyCode = 13 Then cmbSaveReady_Click
    Else: KeyCode = 0
    End If
End Sub
'#######################################################
'изменение отдела по умолчанию
Private Sub tboxDepPrefix_Change()
    If bolUF_Set_Load = False Then Exit Sub
    WorkClsm.depPrefix = tboxDepPrefix
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub
'#######################################################
'изменение лаборатории по умолчанию
Private Sub tBoxLabNum_Change()
    If bolUF_Set_Load = False Then Exit Sub
    WorkClsm.labNum = tBoxLabNum
    PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
End Sub

'#######################################################
'выбор директории протоколов поверки
Private Sub cmbChoosePrDir_Click( _
    )
    
    Dim sMyPath As String, _
        sTempPath As String
    
    sTempPath = WorkClsm.xlPrPath
    
    If FolderNotExist(sTempPath) Then _
        sTempPath = "недоступно"
    
    If sTempPath <> "недоступно" Then _
        sTempPath = Left(sTempPath, Len(sTempPath) - InStr(2, StrReverse(sTempPath), "\") + 1) 'подняться на каталог выше
        
    sMyPath = GetFolderFPath(, sTempPath)  'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
    
        sMyPath = sMyPath & "\"
        WorkClsm.xlPrPath = sMyPath  'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
    
End Sub
Private Sub cmbOpenPrDir_Click()
    Explorer.OpenFolder WorkClsm.xlPrPath, True
End Sub
'#######################################################
'выбор директории протоколов калибровки
Private Sub cmbChoosePrcDir_Click( _
    )
    Dim sMyPath As String, _
        sTempPath As String
    
    sTempPath = WorkClsm.xlPrcPath
    
    If FolderNotExist(sTempPath) Then _
        sTempPath = "недоступно"
    
    If sTempPath <> "недоступно" Then _
        sTempPath = Left(sTempPath, Len(sTempPath) - InStr(2, StrReverse(sTempPath), "\") + 1) 'подняться на каталог выше
        
    sMyPath = GetFolderFPath(, sTempPath)  'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
        
        sMyPath = sMyPath & "\"
        WorkClsm.xlPrcPath = sMyPath  'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
    
End Sub
Private Sub cmbOpenPrcDir_Click()
    Explorer.OpenFolder WorkClsm.xlPrcPath, True
End Sub
'#######################################################
'выбор директории свидетельств о поверке
Private Sub cmbChooseSvDir_Click( _
    )
    Dim sMyPath As String, _
        sTempPath As String
    
    sTempPath = WorkClsm.wdSvPath
    
    If FolderNotExist(sTempPath) Then _
        sTempPath = "недоступно"
    
    If sTempPath <> "недоступно" Then _
        sTempPath = Left(sTempPath, Len(sTempPath) - InStr(2, StrReverse(sTempPath), "\") + 1) 'подняться на каталог выше
        
    sMyPath = GetFolderFPath(, sTempPath)  'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\": WorkClsm.wdSvPath = sMyPath  'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub
Private Sub cmbOpenSvDir_Click()
    Explorer.OpenFolder WorkClsm.wdSvPath, True
End Sub
'#######################################################
'выбор директории сертификатов калибровки
Private Sub cmbChooseSrtDir_Click()
    Dim sMyPath As String, sTempPath As String
    
    sTempPath = WorkClsm.wdSrtPath
    If FolderNotExist(sTempPath) Then _
        sTempPath = "недоступно"
    
    If sTempPath <> "недоступно" Then _
        sTempPath = Left(sTempPath, Len(sTempPath) - InStr(2, StrReverse(sTempPath), "\") + 1) 'подняться на каталог выше
        
    sMyPath = GetFolderFPath(, sTempPath)  'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\": WorkClsm.wdSrtPath = sMyPath  'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub
Private Sub cmbOpenSrtDir_Click()
    Explorer.OpenFolder WorkClsm.wdSrtPath, True
End Sub
'#######################################################
'выбор директории извещений о непригодности к применению
Private Sub cmbChooseInDir_Click()
    Dim sMyPath As String, sTempPath As String
    
    sTempPath = WorkClsm.wdInPath
    
    If FolderNotExist(sTempPath) Then _
        sTempPath = "недоступно"
    
    If sTempPath <> "недоступно" Then _
        sTempPath = Left(sTempPath, Len(sTempPath) - InStr(2, StrReverse(sTempPath), "\") + 1) 'подняться на каталог выше
        
    sMyPath = GetFolderFPath(, sTempPath)  'выбрать путь к каталогу
    
    If sMyPath <> "NoPath" Then
        sMyPath = sMyPath & "\": WorkClsm.wdInPath = sMyPath  'передать папаметр в рабочий класс
        
        UpdateDBLabels
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub
Private Sub cmbOpenInDir_Click()
    Explorer.OpenFolder WorkClsm.wdInPath, True
End Sub
'#######################################################
'экспорт файла локальной конфигурации
Private Sub cmbCfgExp_Click()
    Dim sTempPath As String
    sTempPath = GetSaveAsFname(9, "C:\Users\" & Environ("USERNAME") & "\Desktop\") 'путь к файлу для сохранения
    
    If sTempPath <> "NoPath" Then
        myBase.GetArrFF sTempPath 'загрузить в класс файл новой конфигурации
        myBase.SaveProperties WorkClsm.myParameters, WorkClsm.myValues 'передать значения в настройки и сохранить в файле через класс
        
        Me.VersionLabel.caption = "Конфигурация экспортирована" & vbNewLine & "на рабочий стол."
        
        myBase.GetArrFF setDir, "localConfig.uCfg" 'загрузить в класс текущей конфигурации
    End If
End Sub
'#######################################################
'импорт файла локальной конфигурации
Private Sub cmbCfgImp_Click()
    Dim sTempPath As String
    sTempPath = GetFileFPath(9, "C:\Users\" & Environ("USERNAME") & "\Desktop\")   'получить полный путь к файлу
    
    If sTempPath <> "NoPath" Then
        With myBase 'работа с классом настроек
            '#######################################################
            .GetArrFF sTempPath 'загрузить в класс файл новой конфигурации
            '#######################################################
            .FillValues 'обязательно: найти значения выходных параметров по ключам
        End With
          
        WorkClsm.FillProperties myBase.Parameters, myBase.values 'передать извлечённые из настроек параметры в специальный класс
        
        With myBase 'работа с классом настроек
            '#######################################################
            .GetArrFF setDir, "localConfig.uCfg" 'загрузить в класс файл локальной конфигурации
            '#######################################################
        End With
        
        Me.cmbCfgImp.BackColor = &HFFFFFF    'белый
        
        UpdateDBLabels 'обновить метки выбранных файлов согласно загруженным данным
        UpdateEmployees bolUF_Set_Load 'обновить данные списка исполнителей
        PreSaveSettings 'изменить заголовок кнопки Готово / Сохранить
    End If
End Sub
