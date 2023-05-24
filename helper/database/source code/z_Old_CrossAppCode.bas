Attribute VB_Name = "z_Old_CrossAppCode"
'в модуле собран код для разделения функций Excel & Word

#Const CurrentAppCode = "xl" 'текущая директива выполнения кода
'wd - работа в MsWord. ширина = х, высота = у
'xl - работа в MsExcel. ширина = х - 5, высота = у - 7
Option Explicit 'запрет на использование неявных переменных
'###################################################################
'процедура изменяет размер формы поиска по БД
Sub Set_Z_UF_Search_Size()

    Select Case True  'параметр загрузки типа
        
        Case UMenu.typе = organisationsOLD
            SearchBy1
            
        Case UMenu.typе = instrumentsOLD
            SearchBy2
            
        Case UMenu.typе = etalonsOLD
            SearchBy3
            
        Case UMenu.typе = personsOLD
        
            With Z_UF_Search
            
                .caption = "Сведения о ФИО и должностях сотрудников"
                .cmb1.caption = "Фамилия И.О.": .cmb2.Visible = False
                .cmb4.caption = "«Поверитель»"
                .chbFullName.Visible = True
                
                With .cmb3
                    .caption = "Должность"
                    .Enabled = True
                    
                End With
                
            End With
            
        Case UMenu.typе = archiveOLD
            SearchBy14
            
    End Select
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 13:46
' Назначение:
' ----------------------------------------------------------------
Sub SearchBy1()
    With Z_UF_Search
        
        .caption = "Контрагенты ВНИИМ" 'заголовок
        #If CurrentAppCode = "wd" Then 'только для приложения Word
            .cmb1.caption = "Передать сведения" & vbNewLine & "в позицию курсора"
        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
            .cmb1.caption = "Передать" & vbNewLine & "на лист"
        #End If
        
        .cmb1.Height = 48
        .cmb1.Width = .cmb1.Width / 2
        
        With .cmb2
            .caption = "Наименование": .Top = 88
        End With
        
        With .cmb3
            .Top = 112: .Width = 57: .caption = "ИНН":
        End With
        
        With .cmb4
            .Left = 63: .Width = 57: .caption = "Адрес"
        End With
        
        .btnOpenFolder.Visible = True
        .btnOpenFolder.Left = .cmb4.Left
        .btnOpenFolder.Top = .cmb1.Top
        .btnOpenFolder.Width = .cmb1.Width
        .btnOpenFolder.Height = .cmb1.Height
        
    End With
End Sub
' ----------------------------------------------------------------
' Дата: 25.02.2023 13:46
' Назначение:
' ----------------------------------------------------------------
Sub SearchBy2()

    With Z_UF_Search
    
        .caption = "Средства измерений" 'заголовок
        
        With .cmbProtSv 'кнопка загрузки шаблона свидетельства или протокола
            
            .Visible = True
            .Height = 48
            .Width = .Width / 2
            .Left = 6
            .caption = "Загрузить" & vbNewLine & "шаблон"
            
        End With
        
        With .cmbDescription 'описание типа СИ
            .Visible = True: .Top = 88: .Left = 6: .Width = 57
        End With

        With .cmbMetodic 'методика поверки
            .Visible = True: .Left = 63: .Width = 57: .ControlTipText = "Открыть методику поверки для данного СИ"
        End With
        
        With .cmb1 'наименование СИ
            .Width = 29: .Top = 112: .caption = "Н": .ControlTipText = "Передать наименование СИ"
        End With
        
        With .cmb2 'Типовой состав СИ
            .Width = 28: .Top = 112: .Left = 35: .caption = "Т": .ControlTipText = "Передать типовой состав СИ"
        End With
        
        With .cmb3 'Номер в фиф
            .Width = 28: .Top = 112: .Left = 64: .caption = "№": .ControlTipText = "Передать номер в ФИФ"
        End With

        With .cmb4 'Методика поверки
            .Width = 28: .Left = 92: .caption = "М": .ControlTipText = "Передать наименование МП"
        End With
        
        .btnOpenFolder.Visible = True
        .btnOpenFolder.Left = .cmbMetodic.Left
        .btnOpenFolder.Top = .cmbProtSv.Top
        .btnOpenFolder.Width = .cmbProtSv.Width
        .btnOpenFolder.Height = .cmbProtSv.Height
        
    End With
End Sub
'###################################################################
'размеры формы поиска эталонов
Sub SearchBy3()
     With Z_UF_Search
     
        .caption = "Эталоны и вспомогательное оборудование" 'заголовок
                
        With .cmbProtSv
            .Visible = True: .Left = 6: .Height = 48
        
            #If CurrentAppCode = "wd" Then 'только для приложения Word
                .caption = "Передать сведения" & vbNewLine & "в позицию курсора"
            #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                .caption = "Поиск и заполнение" & vbNewLine & "полей оборудования"
            #End If
        End With
        
        With .cmb1 'Наименование оборудования
            .Top = 88: .Width = 57: .caption = "Наим.": .ControlTipText = "Наименование эталонного оборудования"
        End With
        
        With .cmb2 'Тип оборудования
            .Top = 88: .Width = 57: .caption = "Тип": .Left = 63: .ControlTipText = "Тип эталонного оборудования"
        End With
        
        With .cmb3 'номер оборудования
            .Top = 112: .Width = 57: .caption = "Номер": .ControlTipText = "Закреплённый номер"
        End With
        
        With .cmb4 'сведения о поверке
            .Left = 63: .Width = 57: .caption = "Прим.": .ControlTipText = "Примечание (сведения о поверке / калибровке)"
        End With
    End With
End Sub
'###################################################################
'размеры формы поиска по архиву выполненных работ
Sub SearchBy14()
    With Z_UF_Search
        .cmbUpdate.Left = 336: .caption = "Архив" 'заголовок
        
        With .btnOpenFolder
            .Visible = True: .Enabled = True: .ControlTipText = "Открыть каталог архива работ"
        End With
        
        With .cmb1
            .Height = 48: .caption = "Открыть каталог" & vbNewLine & "архива работы"
        End With
        
        With .cmb3
            #If CurrentAppCode = "wd" Then 'только для приложения Word
                If ActiveDocument.path <> "" Then .Enabled = True
            #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                If ActiveWorkbook.path <> "" Then .Enabled = True
            #End If
            
            .Height = 48: .caption = "Создать каталог" & vbNewLine & "архива работы"
        End With
        .cmb2.Visible = False: .cmb4.Visible = False
        
        With .chbFullName
            .caption = "запретить изменение"
            .ControlTipText = "запретить редактирование текущего протокола"
            '.value = True
            .Visible = True
            .Width = .Width + 20
            .Top = Z_UF_Search.btnOpenFolder.Top + 3
            .Left = Z_UF_Search.cmbUpdate.Left - Z_UF_Search.cmbUpdate.Width - 20
        End With
    End With
End Sub
'###################################################################
'процедура добавляет параметр для загрузки
Sub AddInvertParameter(objClsm As Z_clsmBase)
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        objClsm.AddP "wdFullFirstName"
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        objClsm.AddP "xlFullFirstName"
    #End If
End Sub
'###################################################################
'функция определяет дальшейнее выполнение кода
Function myWdDoc() As Boolean
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        myWdDoc = True
    #End If
End Function
'###################################################################
'процедура передаёт данные в документ в зависимости от типа приложения MS Office
Function Z_UF_Search_Cmb2(ByRef DataBase() As String, ByRef dbInd As Integer) As Boolean

    #If CurrentAppCode = "wd" Then 'только для приложения Word
    
        If UMenu.typе = organisationsOLD Then _
            DataTransfer "ИНН " & DataBase(1, dbInd), True: Exit Function 'только для формы заказчиков
        DataTransfer DataBase(1, dbInd), True
        
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
    
        If InStr(ActiveCell.Offset(0, -1), "ИНН") > 0 Then _
            DataTransfer DataBase(1, dbInd)
        
        If InStr(ActiveCell.Offset(-1, 3), "Юр. адрес:") > 0 Then _
            ActiveCell.Offset(-1, 2).Select: Z_UF_Search_Cmb2 = True: Exit Function
    #End If
End Function
'###################################################################
'процедура правильно активирует кнопку передачи сведений в документ
Sub EtalonsSearchButton()
    With Z_UF_Search
    
        #If CurrentAppCode = "wd" Then 'только для приложения Word
            .cmbProtSv.Enabled = .cmb1.Enabled
        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
            .cmbProtSv.Enabled = True
        #End If
    End With
End Sub
'###################################################################
'процедура изменяет размеры формы конструктора БД
Sub Set_UF_Cunstructor_Properties()
    
    Select Case True   'параметр загрузки элементов формы
        
        Case UMenu.typе = organisationsOLD  'заказчики
            ConstructorBy1
        
        Case UMenu.typе = instrumentsOLD  'средства измерений
            ConstructorBy2
        
        Case UMenu.typе = etalonsOLD   'эталоны
            ConstructorBy3
        
        Case UMenu.typе = personsOLD  'фамилии
            With Z_UF_Constructor
                .chbVerRefer.Visible = True
                .caption = "Конструктор сведений о ФИО и должностях сотрудников"
                
                .Label2.caption = "Фамилия, имя и отчество полностью:"
                .Label2.Width = .Label4.Width: .TextBox1.Width = .TextBox2.Width
                
                .Label3.Visible = False: .TextBox2.Visible = False
                
                .Label4.caption = "Должность:": .Label4.Top = .Label2.Top
                .TextBox3.Top = .TextBox1.Top: .TextBox3.BackColor = .TextBox1.BackColor
                
                .Label5.Visible = False: .TextBox4.Visible = False
                    
                .cmbImport.Top = .Label3.Top: .cmbDelete.Top = .Label3.Top
                .cmbAdd.Top = .Label3.Top: .cmbReady.Top = .Label3.Top
                .LabelInfo.Top = .Label3.Top + 5: .Label3.Top = .Label2.Top
                
                '.cmbOpenTemplateFolder.Visible = False
                Select Case Application.Version
                    Case "16.0" 'Office 2016
                        .Height = 290
                    Case "15.0" 'Office 2013
                        #If CurrentAppCode = "wd" Then 'только для приложения Word
                            .Height = 290
                        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                            .Height = 284
                        #End If
                    Case Else '2010, 2007
                        .Height = 284
                End Select
            End With
    End Select
End Sub
'###################################################################
'процедура корректирует размеры формы
Sub ConstructorBy1()
    With Z_UF_Constructor
    
        .caption = "Конструктор сведений о заказчиках" 'заголовок
        .Label2.caption = "Полное наименование заказчика:": .Label3.caption = "Индивидуальный налоговый номер:"
        .Label4.caption = "Сокращение наименования для архива:": .Label5.caption = "Адрес заказчика:"
        
        With .TextBox2
            .maxLength = 13: .BackColor = &HC0FFFF    'жёлтый
        End With
        
        .TextBox3.maxLength = 27
        .cmbOpenTemplateFolder.Visible = False
    End With
End Sub
'###################################################################
'процедура корректирует размеры формы средств измерений
Sub ConstructorBy2()
    With Z_UF_Constructor
        
        .caption = "Конструктор сведений о средствах измерений" 'заголовок
                
         With .Label4
             .caption = "Номер в ФИФ:": .Left = 378: .Width = 84
         End With
         
         .Label4.Top = .Label2.Top: .Label4.Left = .Label2.Left
         
         With .TextBox3 'номер в фиф
             .Width = 84: .maxLength = 8: .BackColor = &HC0FFFF   'жёлтый
         End With
         .TextBox3.Top = .TextBox1.Top: .TextBox3.Left = .TextBox1.Left
         
         .TextBox2.Width = .TextBox3.Width  'поле ввода типового состава СИ
         .TextBox2.BackColor = &HC0FFFF   'жёлтый
         
         With .Label3
             .Width = 400: .caption = "Тип СИ:"
         End With
                            
         .TextBox1.Left = .TextBox1.Left + .TextBox3.Width
         .TextBox1.Width = .TextBox1.Width - .TextBox3.Width
         
         .Label2.caption = "Полное наименование средства измерений:"
         .Label2.Left = .TextBox1.Left
         
         .LabelInfo.Top = .Label5.Top + 4
         .cmbImport.Top = .Label5.Top
         
         .cmbDelete.Top = .Label5.Top
         .cmbAdd.Top = .Label5.Top
         .cmbReady.Top = .Label5.Top
         
         .Label5.Top = .Label3.Top: .Label5.Left = .TextBox1.Left: .Label5.caption = "Методика поверки:"

         .TextBox4.Top = .TextBox2.Top: .TextBox4.Left = .TextBox1.Left
         .TextBox4.Width = .TextBox1.Width: .TextBox4.BackColor = &HC0FFFF    'жёлтый
         .cmbOpenTemplateFolder.Visible = True
         
        With .chbVerRefer
            .Visible = True: .Width = 130: .Left = 340: .caption = "перекрёстная ссылка"
            .ControlTipText = "Вставить ссылку на загрузку шаблона СИ из директории другого номера ФИФ"
        End With
        
        .chbEtalon.Visible = True
        
        If myWdDoc = False Then .tboxSearchConstr.Width = 126: .cmbFillTempProp.Visible = True: .LabelGetTemplate.Visible = True
         
         Select Case Application.Version
             Case "16.0" 'Office 2016
                 .Height = 332
             Case "15.0" 'Office 2013
                 #If CurrentAppCode = "wd" Then 'только для приложения Word
                     .Height = 332
                 #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                     .Height = 324
                 #End If
             Case Else '2010, 2007
                 .Height = 324
         End Select
    End With
End Sub
'###################################################################
'процедура корректирует размеры формы эталонов
Sub ConstructorBy3()
    With Z_UF_Constructor
    
        .caption = "Конструктор сведений об эталонном и вспомогательном оборудовании" 'заголовок
        .Label2.caption = "Наименование оборудования:"
        
        With .Label6
            .Visible = True: .Left = 6: .Width = 150: .Top = 234: .caption = "Ключевая фраза для поиска:"
        End With
        With .TextBox5 'ключевая фраза для поиска
            .Visible = True: .Left = 6: .Width = 170: .Top = 252: .BackColor = &HC0FFFF    'жёлтый
        End With
        
        With .Label3
            .Left = 176: .Width = 130: .caption = "Тип оборудования:"
        End With
        With .TextBox2 'Тип эталона
            .Left = 176: .Width = 130: .BackColor = &HC0FFFF    'жёлтый
        End With
        
        With .Label4
           .Left = 306: .caption = "Закреплённый номер:"
        End With
        With .TextBox3 'заводской номер
            .Left = 306: .Width = 156: .maxLength = 25: .BackColor = &HC0FFFF   'жёлтый
        End With
    
        .Label5.caption = "Примечание (сведения о поверке):"
        .TextBox4.BackColor = &HC0FFFF    'жёлтый
        
        .cmbOpenTemplateFolder.Visible = False
    End With
End Sub
'###################################################################
'процедура передаёт в открытую книгу xl фамилию исполнителя
'todo: [+] FillLastName -- рефакторить и перенести в LoadXlTemplate
Sub FillLastName( _
    sName As String, _
    sNameState As String, _
    Optional bolNameHead As Boolean, _
    Optional bolNameSecond As Boolean, _
    Optional objWs As Object _
    )
    '##################################################
    'ВРЕМЕННОЕ РЕШЕНИЕ
    If ActiveWorkbook.name Like "jr*" Then Exit Sub
    
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        
        If objWs Is Nothing Then Set objWs = ActiveWorkbook.ActiveSheet
        With objWs
        
            Dim ilastRow As Integer, sCellStr As String
            ilastRow = .Cells(Rows.count, 1).End(xlUp).Row 'последняя заполненная строка 1 столбца снизу
            
            Do While ilastRow > 24 '25 строка - последняя для поиска
                
                sCellStr = LCase(CStr(.Cells(ilastRow, 1)))
                                
                If bolNameHead And sCellStr Like "*утвердившего протокол*" Then Exit Do  'строка руководителя
                If bolNameSecond And sCellStr Like "*торой*" And sCellStr Like "*исполнит*" Then Exit Do 'строка исполнителя
                If sCellStr Like "*оверку*" Or sCellStr Like "*ыполнил*" Or sCellStr Like "*произв*" Then Exit Do 'строка исполнителя
                
                ilastRow = .Cells(ilastRow, 1).End(xlUp).Row 'перейти к следующей строке выше
            Loop
            
            If ilastRow = 25 Then Exit Sub
            
            If sName = "недоступно" Or sName = vbNullString Then
            
                If bolNameSecond Then _
                    .Cells(ilastRow - 2, 1).Resize(4, 1).EntireRow.Delete
                
                Exit Sub
                
            End If
            
            
            If InStr(CStr(.Cells(ilastRow + 1, 9)), "фамилия") > 0 Then _
                .Cells(ilastRow, 9) = sName 'передать имя
            
            If InStr(CStr(.Cells(ilastRow + 1, 3)), "должность") > 0 Then _
                .Cells(ilastRow, 3) = sNameState 'передать должность
                
            If .Cells(ilastRow, 3) <> "" Then _
                SetPrintArea ilastRow, objWs 'заполнена должность
        End With

        
    #End If
    
End Sub
Private Sub SetPrintArea(ilastRow As Integer, Optional objWs As Object)

    If objWs Is Nothing Then Set objWs = ActiveSheet
    
    Dim prArea As String, printMaxRow As Integer
    prArea = ActiveSheet.PageSetup.PrintArea
    printMaxRow = Right(prArea, InStr(StrReverse(prArea), "$") - 1)
    
    'ActiveSheet.PageSetup.PrintArea = "a1:j136"
    
    If ilastRow + 1 > printMaxRow Then _
        prArea = Left(prArea, Len(prArea) - Len(printMaxRow) - 1) & ilastRow + 1

    objWs.PageSetup.PrintArea = prArea
    Set objWs = Nothing
End Sub
'###################################################################
'процедура открывает pdf-файл
Sub OpenPDF(sourcePath As String, sFifNum As String, pdfMAsk As String)
    
    Dim sFileName As String, sTempDir As String
    sTempDir = Dir(sourcePath & "\instruments\" & "*" & sFifNum & "*", vbDirectory)   'каталог СИ
    sFileName = Dir(sourcePath & "\instruments\" & sTempDir & "\" & pdfMAsk & "*.pdf") 'имя файла
    
    If sTempDir <> "" Then _
        sFileName = sourcePath & "\instruments\" & sTempDir & "\" & sFileName 'каталог + имя файла
    
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        ActiveDocument.FollowHyperlink sFileName
        
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        ActiveWorkbook.FollowHyperlink sFileName
    #End If
End Sub
'###################################################################
'процедура переводит фокус на каталог
Sub openPath(sPath As String, Optional sFName As String)

    #If CurrentAppCode = "wd" Then 'только для приложения Word
        ActiveDocument.FollowHyperlink sPath & sFName
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        ActiveWorkbook.FollowHyperlink sPath & sFName
    #End If
End Sub
'###################################################################
'промежуточная процедура выбора ячейки только для xl
Function TrueFindCell() As Boolean

    TrueFindCell = False
    
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        
        Dim rTempCell As Range
        TrueFindCell = FindCellRight("Заказчик", rTempCell, , , True)
    #End If
End Function
'###################################################################
'процедура передаёт данные в документ / файл
Sub DataTransfer( _
    ByVal sData As String, _
    Optional insertEndSpace As Boolean, _
    Optional useLongName As Boolean, _
    Optional clearCellColors As Boolean = True, _
    Optional currAddress As String)
    
    Application.ScreenUpdating = False 'отключить обновление экрана
    
    If UMenu.typе = personsOLD And useLongName = False Then 'передать неполную форму ФИО
    
        Dim sArrTemp() As String, K As Byte, sTempStr As String
        sArrTemp = Split(sData, " ")
        sTempStr = sArrTemp(0) & " "
        
        For K = LBound(sArrTemp) + 1 To UBound(sArrTemp)
            sTempStr = sTempStr & Left(sArrTemp(K), 1) & "."
        Next
        
        sData = sTempStr
    End If

    #If CurrentAppCode = "wd" Then 'только для приложения Word
        Application.DisplayAlerts = word.wdAlertsNone 'отключить уведомления на экране
        
        With Selection
            If insertEndSpace Then sData = sData & " " 'добавить пробел после передаваемых данных

            .text = sData  'передать данные в позицию курсора
            .Range.HighlightColorIndex = wdNoHighlight 'очистить заливку цветом
            .Range.Font.ColorIndex = wdBlack 'черный цвет шрифта
            .MoveRight Unit:=wdCharacter 'переместиться на один элемент вправо
        End With
        
        Application.DisplayAlerts = word.wdAlertsAll
        
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        
        If currAddress <> vbNullString Then _
            ActiveSheet.Range(currAddress).Select
        
        With ActiveCell
        
            If clearCellColors Then
                .Font.ColorIndex = xlAutomatic 'чёрный цвет шрифта
               ' .Interior.Pattern = xlNone 'очистить заливку
            End If
            
            .numberFormat = "@"
            .value = sData
            
            Select Case True
                
                Case UMenu.typе = etalonsOLD
                    .Offset(0, 1).Select 'только для эталонного оборудования
                    
                Case Else:
                    .Offset(1, 0).Select
            End Select

        End With
    #End If
    
    Application.ScreenUpdating = True
End Sub
'###################################################################
'процедура передаёт наименование заказчика в свойство документа
Sub SetBuiltInProperty( _
    sProperty As String, _
    Optional sData As String, _
    Optional ClearProp As Boolean _
)
    
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            If ClearProp Then .BuiltinDocumentProperties(sProperty) = vbNullString: Exit Sub
            
'            Dim currProperty As String
'            currProperty = .BuiltinDocumentProperties(sProperty)
'
'            If InStr(currProperty, "--") > 0 Then
'                currProperty = Right(currProperty, Len(currProperty) - InStr(currProperty, "--") + 2)
'                sData = sData & currProperty
'
'            End If
                
            .BuiltinDocumentProperties(sProperty) = sData
        End With
End Sub
'###################################################################
'функция извлекает встроенное свойство
Function GetBuiltInProperty(sProperty As String) As String
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            GetBuiltInProperty = .BuiltinDocumentProperties(sProperty)
        End With
End Function
'###################################################################
'процедура сохраняет документ в зависимости от выбранных опций - переименовать открытый документ
Sub SaveNewName(SaveName As String, Optional saveAsCopy As Boolean, _
    Optional bolShortCut As Boolean, Optional NotInsNumber As Boolean, Optional sSavingDir As String)
    
'    SaveName - конечное имя файла без пути и расширения pr_210_2104_0000_18
'    saveAsCopy - сохранть как копию (не перезаписывать текущий документ)
'    bolShortCut - создать ярлык на рабочем столе
'    NotInsNumber - не передавать в комментарий заводской номер (при создании шаблона)
    #If CurrentAppCode = "wd" Then 'только для приложения Word
    
        With ActiveDocument
            If .name = ThisDocument.name Then MsgBox "Сейчас невозможно": Exit Sub  'защита от закрытия документа при разработке
            
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
            If .name = ThisWorkbook.name Then MsgBox "Сейчас невозможно": Exit Sub 'защита от закрытия документа при разработке
            Commit_factory_number
    #End If
            
            Dim sFinalDir As String, _
                sExt As String
            
            sFinalDir = .path & "\"
            If .path = vbNullString Then sFinalDir = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
            If sSavingDir <> vbNullString Then sFinalDir = sSavingDir  'конечная директория
            
            sExt = GetExt(.fullName) 'текущее расширение файла
            
            If sExt = vbNullString Then
                saveAsCopy = True 'текущий файл является временным и не сохранён
                
                #If CurrentAppCode = "wd" Then 'только для приложения Word
                    sExt = ".docx"
                #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                    sExt = ".xlsx"
                #End If
            End If
            
            Dim sNewName As String, sPrevName As String
            sPrevName = .fullName 'текущее имя книги
            sNewName = ReturnNotExistingName(sFinalDir, ReplaceBadSymbols(SaveName), sExt)  'новое имя файла
            
            #If CurrentAppCode = "wd" Then 'только для приложения Word
                .SaveAs sNewName, wdFormatDocumentDefault 'почему-то иначе не работает
            #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                .SaveAs sNewName, 51
            #End If
            
            If saveAsCopy = False Then Kill sPrevName 'если необходимо переименовать
            If bolShortCut Then _
                CreateShortcut sNewName, SaveName
        End With
End Sub
'###################################################################
'процедура создаёт ярлык текущего документа на рабочем столе
Sub CreateShortcut( _
    path As String, _
    Optional shcutName As String _
    )
    
    If shcutName = vbNullString Then _
        shcutName = Right(path, InStr(StrReverse(path), Application.PathSeparator) - 1)
        
    With CreateObject("WScript.Shell")
    
        If Dir(.SpecialFolders("Desktop") & "\" & shcutName & ".lnk") = vbNullString Then 'если ярлыка не существует
        
            With .CreateShortcut(.SpecialFolders("Desktop") & "\" & shcutName & ".lnk") 'создать ярлык
                .targetPath = path 'полный путь к файлу
                .description = "Перейти в каталог"
                .Save
            End With
        End If
    End With
End Sub
'###################################################################
 'функция извлекает встроенное свойство из открытого документа
Function TakeCategoryProperty(property As String)
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            TakeCategoryProperty = .BuiltinDocumentProperties(property) 'текущее значение параметра в свойствах книги
        End With
End Function
'###################################################################
'функция задаёт формат отображения даты в листбоксах
Function ListBoxDateFormatChange(myDate As Date, Optional bolAnotherDateFormat As Boolean) As String
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        
        ListBoxDateFormatChange = Format(myDate, "dd mmmm yyyy г.") 'текстовый формат
        If bolAnotherDateFormat Then ListBoxDateFormatChange = Format(myDate, "dd.mm.yyyy"): Exit Function 'числовой формат
            
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        
        ListBoxDateFormatChange = Format(myDate, "dd.mm.yyyy") 'числовой формат
        If bolAnotherDateFormat Then ListBoxDateFormatChange = Format(myDate, "dd mmmm yyyy г."): Exit Function 'текстовый формат
        
    #End If
End Function
'###################################################################
'процедура передаёт в документ дату выполнения работы
Sub TrueDateTransfer(myDate As String, myInterval As Integer, FirstDateSelected As Boolean, SecondDateSelected As Boolean, AdditionDateFormat As Boolean)
    Application.ScreenUpdating = False
    
    Dim DateToTransfer As String
    DateToTransfer = DateAdd("yyyy", myInterval, CDate(myDate)) - 1 'передаётся дата действительно до
    
    If FirstDateSelected = True And SecondDateSelected = False Then DateToTransfer = myDate 'дополнительная проверка, что передаётся дата выполнения работы

    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With Selection 'работа с выделенным текстом
        
            If AdditionDateFormat Then 'если выбран числовой формат "ДД.ММММ.ГГГГ"
                .text = DateToTransfer
            Else 'формат в виде "ДД мммм ГГГГ г."
                .text = Format(CDate(DateToTransfer), "dd mmmm yyyy г.")
            End If
            
            .Range.HighlightColorIndex = wdNoHighlight 'очистить заливку цветом
            .Range.Font.ColorIndex = wdBlack 'черный цвет шрифта
            .MoveRight Unit:=wdCharacter 'переместиться на один элемент вправо
            
        End With
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveCell
            .Font.ColorIndex = xlAutomatic 'чёрный цвет шрифта
            .Interior.Pattern = xlNone 'очистить заливку
            
            If AdditionDateFormat Then 'если выбран текстовый формат "ДД мммм ГГГГ г."
                .numberFormat = "@": .value = Format(CDate(DateToTransfer), "dd mmmm yyyy г.")
            Else 'формат в виде "ДД.ММММ.ГГГГ"
                .value = CDate(DateToTransfer)
            End If
            
            .Offset(1, 0).Select 'выделить ячейку ниже
        End With
    #End If
    Application.ScreenUpdating = True
End Sub
'###################################################################
'функция импортирует в данные и удаляет пустые символы в начале и конце данных
Function DataImport() As String 'импортировать данные из документа
    Application.ScreenUpdating = False 'отключить обновление экрана
    
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        Application.DisplayAlerts = word.wdAlertsNone
        
        DataImport = DeleteSpaceStEnd(Selection)  'передать данные из выделенной области
        
        Application.DisplayAlerts = word.wdAlertsAll
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        
        DataImport = DeleteSpaceStEnd(ActiveCell)
    #End If
    
    Application.ScreenUpdating = True
End Function
'###################################################################
'процедура реализует импорт данных в буфер обмена по нажатию на Num +
'Sub NumPlusImport()
'    With F_UF_Buffer
'
'        #If CurrentAppCode = "wd" Then 'только для приложения Word
'            If Selection.Characters.Count > 1 Then
'                .tboxBuffer = DataImport
'        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
'            If ActiveCell <> "" Then
'
'                .tboxBuffer = DataImport & "@" & ActiveCell.address
'        #End If
'                .Label2.Caption = "импорт выполнен"
'                TrueElementForeColor .Label2: OntimeUnload 'если кнопка импорта доступна
'            Else
'                Dim sLabelText As String
'                sLabelText = .Label2.Caption: .Label2.Caption = "выберите данные для импорта": .Label2.Forecolor = &HFF&  'красный
'
'                Application.OnTime Now + TimeValue("00:00:02"), "LabelBack"
'            End If
'    End With
'End Sub
'###################################################################
'метод правильно передаёт цвет текста контрола
Sub TrueElementForeColor(objName As Object, Optional clrBlack As Boolean, Optional clrWDXL As Byte)
    With objName
    
        If clrBlack Then .foreColor = &H80000008: Exit Sub 'чёрный цвет объекта
        
        #If CurrentAppCode = "wd" Then 'только для приложения Word
            .foreColor = &HFF0000    'синий цвет
        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
            .foreColor = &H8000& 'зелёный цвет шрифта
        #End If
        
        Select Case clrWDXL 'принудительная окраска независимо от приложения
            Case 1: .foreColor = &HFF0000   'синий цвет
            Case 2: .foreColor = &H8000& 'зелёный цвет шрифта
            Case 3: .foreColor = &HFF&       'красный цвет шрифта
            Case 4: .foreColor = &H808000    'бирюзовый
        End Select
    End With
End Sub


'######################################################
'Процедура выделяет ячейку для передачи даных
Sub TrueCellSelInNum()

    #If CurrentAppCode = "xl" Then 'только для приложения Excel

        With ActiveWorkbook
            If .Worksheets.count > 1 Then .Worksheets(1).Select 'если листов больше 1, то выбрать первый
            
            With .ActiveSheet
            
                Dim rTempCell As Range, i As Byte, j As Byte
                If FindCellRight("состав", rTempCell) Then If rTempCell = "" Then rTempCell = Chr(150) 'заменить двойной дефис на тире
                Set rTempCell = Nothing
                            
                Dim sTempStr As String
                sTempStr = ActiveSheet.PageSetup.PrintArea
                If InStr(sTempStr, ":") > 0 Then ' область печати
                    sTempStr = Right(sTempStr, Len(sTempStr) - InStr(sTempStr, ":")): sTempStr = Replace(sTempStr, "$", "")
                    
                    Dim bColumnIndex As Byte
                    Select Case Left(sTempStr, 1)
                        Case "J": bColumnIndex = 10 'для моих протоколов
                        Case Else: bColumnIndex = 16 'для других протоколов (дозиметристы, Гриша)
                    End Select
                    
                    For i = 4 To 11
                        For j = 1 To 4
                            If InStr(CStr(.Cells(i, bColumnIndex + j)), "номер протокола") > 0 Then .Cells(i - 1, bColumnIndex + j).Select: Exit Sub
                        Next j
                    Next i
                    
                End If
            End With
        End With
    #End If
End Sub


'###################################################################
'процедура определяет параметры экспорта документа и вызывает процедуру экспорта по параметрам
Sub ExportCopyOfDocument(objClsmExport As Object, SaveToSomnium As Boolean, _
        SaveToDesktop As Boolean, CreatePDF As Boolean, GroupExport As Boolean)
    
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            Dim sCurrentName As String
            sCurrentName = .name: SaveMeIfTemp sCurrentName   'если файл не был сохранён ранее, то сохранить его перед проведением манипуляций

            If GroupExport Then 'отправить все файлы текущего каталога
                GroupFileExport objClsmExport, .path, SaveToSomnium, SaveToDesktop ' конечная процедура группового экспорта
            Else 'отправить только открытый файл
                
                Dim sDestinationPath As String
                If SaveToSomnium Or SaveToDesktop Then 'выбрана опция отправки на сервер и/или на рабочий стол
                    
                    If SaveToSomnium Then 'отправка на сервер
                        sDestinationPath = GetDestinationToServer(objClsmExport, sCurrentName, .path)
                        ExportToDestPath sCurrentName, .path, sDestinationPath, CreatePDF
                    End If
                                        
                    If SaveToDesktop Then 'отправка на рабочий стол
                        sDestinationPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
                        ExportToDestPath sCurrentName, .path, sDestinationPath, CreatePDF
                    End If
                Else 'по умолчанию копия направляется в текущий каталог
                    sDestinationPath = .path & "\"
                    ExportToDestPath sCurrentName, .path, sDestinationPath, CreatePDF
                End If
            End If
        End With
End Sub
'###################################################################
'организационная процедура направления файлов экспорта
Sub ExportToDestPath(sCurrentName As String, sCurrentPath As String, _
    sDestinationPath As String, Optional CreatePDF As Boolean, Optional ReplaceExistingName As Boolean)
    
    If CreatePDF Then 'сохранение файла в формате pdf
        sCurrentName = GetFileNameWithOutExt(sCurrentName) 'выделить имя файла без расширения: pr_210_2104_222_18
        TrueExportWithPDF ReturnNotExistingName(sDestinationPath, sCurrentName, ".pdf")
    Else
        If ReplaceExistingName Then 'перезаписывать имеющееся имя
            TrueExportWithPDF sDestinationPath & "\" & sCurrentName, sCurrentPath & "\" & sCurrentName
        Else
            TrueExportWithPDF ReturnNotExistingName(sDestinationPath, sCurrentName), sCurrentPath & "\" & sCurrentName
        End If
    End If
    
    If InStr(sDestinationPath, "\Desktop\") = 0 Then _
        Explorer.OpenFolder sDestinationPath, True   'открыть каталог в фоне
End Sub
'###################################################################
'функция возвращает каталог назначения
Function GetDestinationToServer(objClsmExport As Object, sCurrentName As String, sCurrentPath As String) As String
    
    Dim sCurrFileMask As String, sDestinationSecond As String, sSecondFileName As String, sDestinationPath  As String
    sCurrFileMask = GetFileNameWithOutExt(sCurrentName) 'выделить имя файла без расширения: pr_210_2104_222_18
    
    If InStr(sCurrFileMask, "TEMP") = 0 Then  'не экспортировать на сервер файлы без номера
        If InStr(sCurrFileMask, "_") > 0 Then sCurrFileMask = Right(sCurrFileMask, InStrRev(sCurrFileMask, "_") - 1)
        
        Select Case Right(objClsmExport.sendType, 1)
            Case 1 'протокол / свидетельство
                #If CurrentAppCode = "wd" Then 'только для приложения Word
                    sDestinationPath = objClsmExport.wdSvPath: sDestinationSecond = objClsmExport.xlPrPath
                    sSecondFileName = Dir(sCurrentPath & "\*" & sCurrFileMask & "*.xls*")
                #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                    sDestinationPath = objClsmExport.xlPrPath: sDestinationSecond = objClsmExport.wdSvPath
                    sSecondFileName = Dir(sCurrentPath & "\*" & sCurrFileMask & "*.doc*")
                #End If
                
            Case 2 'протокол / сертификат
                #If CurrentAppCode = "wd" Then 'только для приложения Word
                    sDestinationPath = objClsmExport.wdSrtPath: sDestinationSecond = objClsmExport.xlPrcPath
                    sSecondFileName = Dir(sCurrentPath & "\*" & sCurrFileMask & "*.xls*")
                #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                    sDestinationPath = objClsmExport.xlPrcPath: sDestinationSecond = objClsmExport.wdSrtPath
                    sSecondFileName = Dir(sCurrentPath & "\*" & sCurrFileMask & "*.doc*")
                #End If
            Case 3 'извещение
                #If CurrentAppCode = "wd" Then 'только для приложения Word
                    sDestinationPath = objClsmExport.wdInPath
                #End If
        End Select
    End If

    If sSecondFileName <> "" Then
        TrueExportWithPDF ReturnNotExistingName(sDestinationSecond, sSecondFileName), sCurrentPath & "\" & sSecondFileName
        If InStr(sDestinationSecond, "\Desktop\") = 0 Then _
            Explorer.OpenFolder sDestinationSecond 'открыть каталог в фоне
    End If
        
    GetDestinationToServer = sDestinationPath
End Function
'###################################################################
'процедура сохраняет файл, если у него отсутствует расширение - файл временный
Sub SaveMeIfTemp(sCurrentName As String)
    If Not sCurrentName Like "*.*" Then 'файл НЕ был сохранён ранее - сохранить его на рабочем столе, чтобы впоследствии выполнять с ним манипуляции
        
        #If CurrentAppCode = "wd" Then 'только для приложения Word
            ActiveDocument.SaveAs "C:\Users\" & Environ("USERNAME") & "\Desktop\" & sCurrentName, wdFormatDocumentDefault
        #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
            
            Dim myAppVer As String
            myAppVer = Application.Version
            
            If myAppVer = "11.0" Then 'только для 2003
                ActiveWorkbook.SaveAs "C:\Users\" & Environ("USERNAME") & "\Desktop\" & sCurrentName, xlWorkbook
            Else
                ActiveWorkbook.SaveAs "C:\Users\" & Environ("USERNAME") & "\Desktop\" & sCurrentName, xlOpenXMLWorkbook
            End If
        #End If
    End If
End Sub
'###################################################################
'подпроцедура экспорта группы файлов
Sub GroupFileExport(objClsmExport As Object, sCurrPath As String, SaveToSomnium As Boolean, SaveToDesktop As Boolean)
    
    Dim sArrNamesFiles() As String, collDestIsOpen As New Collection, sCurGroupName As String, sTempFMask As String, i As Integer, sDestinationPath As String
    ReDim sArrNamesFiles(0)
    
    sCurGroupName = Dir(sCurrPath & "\*.xls*")
    Do While sCurGroupName <> "" 'заполнить массив именами файлов xl
        If sArrNamesFiles(UBound(sArrNamesFiles)) <> "" Then ReDim Preserve sArrNamesFiles(UBound(sArrNamesFiles) + 1)
        sArrNamesFiles(UBound(sArrNamesFiles)) = sCurGroupName: sCurGroupName = Dir
    Loop
    
    sCurGroupName = Dir(sCurrPath & "\*.doc*")
    Do While sCurGroupName <> "" 'заполнить массив именами файлов wd
        If sArrNamesFiles(UBound(sArrNamesFiles)) <> "" Then ReDim Preserve sArrNamesFiles(UBound(sArrNamesFiles) + 1)
        sArrNamesFiles(UBound(sArrNamesFiles)) = sCurGroupName: sCurGroupName = Dir
    Loop
    
'    sCurGroupName = Dir(sCurrPath & "\*.pdf")
'    Do While sCurGroupName <> "" 'заполнить массив именами файлов pdf
'        If sArrNamesFiles(UBound(sArrNamesFiles)) <> "" Then ReDim Preserve sArrNamesFiles(UBound(sArrNamesFiles) + 1)
'        sArrNamesFiles(UBound(sArrNamesFiles)) = sCurGroupName: sCurGroupName = Dir
'    Loop
    
    SortMassOne sArrNamesFiles, True: ReplaceRepeateInArrOne sArrNamesFiles, "", True
    SortMassOne sArrNamesFiles: ReduceArrOne sArrNamesFiles
    
    With collDestIsOpen 'создать коллекцию для открытия каталогов
        .Add "False", "pr_": .Add "False", "prc_": .Add "False", "sv_": .Add "False", "srt_": .Add "False", "in_"
    End With
    
    If SaveToSomnium Then 'стоит отметка отправки файлов в каталоги экспорта
        For i = LBound(sArrNamesFiles) To UBound(sArrNamesFiles)
            
            If InStr(sArrNamesFiles(i), "_") > 0 And InStr(sArrNamesFiles(i), "TEMP") = 0 Then 'не экспортировать на сервер файлы без номера
                sTempFMask = Left(sArrNamesFiles(i), InStr(sArrNamesFiles(i), "_"))

                Select Case sTempFMask
                    Case "pr_": sDestinationPath = objClsmExport.xlPrPath 'протоколы поверки
                    Case "prc_": sDestinationPath = objClsmExport.xlPrcPath 'протоколы калибровки
                    Case "sv_": sDestinationPath = objClsmExport.wdSvPath 'свидетельства о поверке
                    Case "srt_": sDestinationPath = objClsmExport.wdSrtPath 'сертификаты калибровки
                    Case "in_": sDestinationPath = objClsmExport.wdInPath 'извещения о непригодности
                End Select
                
                TrueExportWithPDF ReturnNotExistingName(sDestinationPath, sArrNamesFiles(i)), sCurrPath & "\" & sArrNamesFiles(i)
                
                With collDestIsOpen
                    If InStr(sDestinationPath, "\Desktop\") = 0 Then _
                        If .item(sTempFMask) = "False" Then .Remove (sTempFMask): .Add "True", sTempFMask: Explorer.OpenFolder sDestinationPath   'открыть каталог в фоне
                End With
            End If
        Next i
    End If
    
    If SaveToDesktop Then 'стоит отметка отправки файлов на рабочий стол
        sDestinationPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
        
        For i = LBound(sArrNamesFiles) To UBound(sArrNamesFiles)
            TrueExportWithPDF ReturnNotExistingName(sDestinationPath, sArrNamesFiles(i)), sCurrPath & "\" & sArrNamesFiles(i)
        Next i
    End If
    
    Set collDestIsOpen = Nothing
End Sub
'###################################################################
'функция экспортирует текущий документ во вне
Sub TrueExportWithPDF(sNewName As String, Optional sPrevName As String)

    If sNewName <> "отсутствует_номер" Then 'только если у текущего файла в имени TEMP
        If sPrevName = "" Then 'экспорт в pdf
            
            #If CurrentAppCode = "wd" Then 'только для приложения Word
                ActiveDocument.ExportAsFixedFormat sNewName, wdExportFormatPDF, False, wdExportOptimizeForPrint
            #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
                If SheetIsEmpty(ActiveWorkbook.ActiveSheet) Then MsgBox "Невозможно экспортировать PDF из пустого листа.", vbInformation: Exit Sub
                ActiveWorkbook.ExportAsFixedFormat xlTypePDF, sNewName, xlQualityStandard
            #End If
            
        Else 'сохранять как просто копию
            CreateObject("Scripting.FileSystemObject").CopyFile sPrevName, sNewName
        End If
    Else
        MsgBox "Текущему файлу присвоен временный индекс, отправка на сервер отменена.", vbInformation
    End If
End Sub
'#################################################
'функция проверяет, существует ли в данной директории файл в таким именем, и, если существует, прибавляет к нему индекс
Function ReturnNotExistingName(sFinalDir As String, sFileName As String, Optional sExt As String) As String
    Dim i As Integer, sName As String, sTempName As String, sNewName As String
    
    If Right(sFinalDir, 1) <> "\" Then sFinalDir = sFinalDir & "\"
    
    If sExt = "" Then 'если в функцию было передано имя файла совместно с вводом расширения
        sName = GetFileNameWithOutExt(sFileName) 'выделить имя документа без расширения
        sExt = GetExt(sFileName) 'выделить расширение файла
    Else
        sName = sFileName
    End If
    
    sTempName = sName:    i = 1
    
    Do 'повторять, пока файл с указанным именем существует в директории
        sNewName = sFinalDir & sTempName & sExt
        If Dir(sNewName) = "" Then Exit Do
        sTempName = sName & " (" & (i) & ")"
        i = i + 1
    Loop
    
    ReturnNotExistingName = sNewName
End Function
'###################################################################
'процедура открывает каталог текущего документа
Sub OpenMyFolder()
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        Explorer.OpenFolder ActiveDocument.path & "\", True
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        Explorer.OpenFolder ActiveWorkbook.path & "\", True
    #End If
End Sub
'#######################################################
'процедура выгружает форму, если работа ведётся только в xl
Sub TrueUnloadCreateProjectXl()
'    #If CurrentAppCode = "xl" Then 'только для приложения Excel
'        VBA.Unload H_UF_Create_Project_Xl
'    #End If
End Sub
'###################################################################
'функция копирует шаблон в рабочую директорию, а затем открывает шаблон
Function SaveAsTemplate( _
    ByRef objZ_clsmSearch As Z_clsmSearch, _
    myMIPackage As MeasInstrument _
    ) As Boolean
    
    SaveAsTemplate = False
    
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        SaveAsTemplate = LoadXlTemplate(objZ_clsmSearch, myMIPackage)
    #End If
End Function
'#######################################################
'процедура вызывает меню экспорта в word документ из xl протокола
Sub ExportToSv()
'    #If CurrentAppCode = "xl" Then 'только для приложения Word
'        H_UF_Create_Project_Xl.Show False
'    #End If
End Sub
'#######################################################
'функция возвращает расширение в имени файла
Function GetExt(sFileName As String) As String
    If InStr(sFileName, ".") = 0 Then Exit Function
    GetExt = Right(sFileName, InStr(StrReverse(sFileName), "."))
End Function
'#######################################################
'функция возвращает имя файла без расширения
Function GetFileNameWithOutExt(Optional sFileName As String) As String

    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            If sFileName = "" Then sFileName = .name
            
            GetFileNameWithOutExt = sFileName: If InStr(sFileName, ".") = 0 Then Exit Function
            GetFileNameWithOutExt = Left(sFileName, Len(sFileName) - InStr(StrReverse(sFileName), "."))
        End With
End Function
'###################################################################
'промежуточная процедура для устранения ошибки между приложениями
Sub PasteEtalons(sArrDataBase() As String)
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        PasteEtalonsData sArrDataBase
    #End If
End Sub


'#################################################
'процедура задаёт свойства при сохранении
Function SaveNewVersion(sStartDir As String) As String
    Dim sCurrVer As String, sNewVer As String, sDestinationPath As String, sCurVerDate As String, _
        sUpdateFilePath As String, sArrTemp() As String
        
    sDestinationPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
    sUpdateFilePath = sDestinationPath & "\update.ini" 'файл сведений версии обновления
    
    sCurrVer = GetCDProp("Version")
    If FileExist(sUpdateFilePath) Then 'если опознан ранее использованный файл версии
        
        Open sUpdateFilePath For Input As #1 'записать сведения обновления в файл
            sArrTemp = Split(Input$(LOF(1), #1), vbNewLine) 'поместить в массив содержимое всего файла
        Close
        
        sCurrVer = sArrTemp(LBound(sArrTemp) + 1)
    End If

    sNewVer = InputBox("Ввод номера версии", "Сохранение программы", sCurrVer): If sNewVer = "" Then sNewVer = sCurrVer
    LetCDProp "Version", sNewVer: LetCDProp "VersionDate", Date
    LetBIDProp "Comments", sNewVer: LetBIDProp "Category", Date
    
    Dim sCurrentName As String, sCurrentPath As String
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ThisDocument
            .Saved = False: .Save: sCurrentName = .name: sCurrentPath = .path
        End With
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ThisWorkbook
            .Save: sCurrentName = .name: sCurrentPath = .path
        End With
    #End If
    
    ExportToDestPath sCurrentName, sCurrentPath, sDestinationPath, , True 'экспорт на рабочий стол
    
    Open sUpdateFilePath For Output As #1 'записать сведения обновления в файл
        Print #1, Date & vbNewLine & sNewVer
    Close
    
    Explorer.OpenFolder sStartDir, True
End Function
'#################################################
'функция возвращает значение свойства из документа
Function GetCDProp(sPropName As String) As String
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ThisDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ThisWorkbook
    #End If
            GetCDProp = .CustomDocumentProperties(sPropName)
        End With
End Function
'#################################################
'функция присваивает значение свойства из документа
Sub LetCDProp(sPropName As String, sPropText As String)
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ThisDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ThisWorkbook
    #End If
            .CustomDocumentProperties(sPropName) = sPropText
        End With
End Sub
'#################################################
'функция присваивает значение свойства из документа
Sub LetBIDProp(sPropName As String, sPropText As String)
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ThisDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ThisWorkbook
    #End If
            .BuiltinDocumentProperties(sPropName) = sPropText
        End With
End Sub
'#################################################
'промежуточная функция для проверки заполнения листа
Sub FillIfXl(objClsm As Z_clsmConstructor, sArrDataBase() As String, iListIndex As Integer)
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
                
        Dim myInstr As MeasInstrument
        myInstr.sName = sArrDataBase(0, iListIndex) 'наименование
        myInstr.sType = sArrDataBase(1, iListIndex) 'тип
        myInstr.sFif = sArrDataBase(2, iListIndex) 'номер в фиф
        
        myInstr.sModification = InputBox("Введите модификацию СИ" & vbNewLine & "или отмените команду выполнения", , GetDefaultInput(myInstr)) ' "без модификации")
        If myInstr.sModification = "" Then Z_UF_Constructor.cmbFillTempProp.Enabled = True: Exit Sub
        
        Dim sTempStr As String
        sTempStr = objClsm.templatesDir & myInstr.sType & "_" & myInstr.sFif & "\"
        If FolderNotExist(sTempStr) Then MkDir sTempStr 'создать новый каталог
        
        Explorer.OpenFolder sTempStr: FillTempProperties myInstr, sTempStr
    #End If
End Sub
'#########################################################
'процедура создаёт каталог в архиве для помещения туда файлов
Sub TrueMkDir(objZ_clsmBase As Z_clsmBase, objClsmSrch As Z_clsmSearch, sArrDataBase() As String)
    Dim sFullCusName As String, sType As String, sLastSaveDate As String, sDocumentName As String
    
    #If CurrentAppCode = "wd" Then 'только для приложения Word
        With ActiveDocument
    #ElseIf CurrentAppCode = "xl" Then 'только для приложения Excel
        With ActiveWorkbook
    #End If
            sFullCusName = .BuiltinDocumentProperties("Company")
            sType = .BuiltinDocumentProperties("Comments")
            
            sDocumentName = GetFileNameWithOutExt(.name)
            sLastSaveDate = .BuiltinDocumentProperties("Last Save Time")
            
            If InStr(sDocumentName, "_") > 0 Then 'защита от ошибки
                Dim sArrTemp() As String
                sArrTemp = Split(sDocumentName, "_") 'получить отдельные часты наименования
                        
                sDocumentName = sArrTemp(LBound(sArrTemp))
                If UBound(sArrTemp) > 0 Then sDocumentName = sDocumentName & "_" & sArrTemp(UBound(sArrTemp) - 1) & "_" & sArrTemp(UBound(sArrTemp))
                
            End If
        End With
        
    With objClsmSrch
    
        Dim sShortCusName As String, sArrData() As String, i As Integer
        If sFullCusName <> "" Then 'свойство файла было получено
            sArrData = .FillDataBase(objZ_clsmBase.GetArrFF(.startDir, .DbName), True) 'преобразовать массив файла в массив базы данных
        
            For i = LBound(sArrData, 2) To UBound(sArrData, 2) 'получить краткое наименование организации
                If InStr(sArrData(0, i), sFullCusName) > 0 Then sShortCusName = sArrData(2, i): Exit For
            Next
            If sShortCusName = "" Then MsgBox "Заказчик не найден в базе данных. Выполнение остановлено.": Exit Sub
            
        End If
        
        If sFullCusName = "" Then _
            sShortCusName = InputBox("Сведения заказчика не опознаны. Введите ключевое слово заказчика:", "Ввод данных"): If sShortCusName = "" Then Exit Sub
        
        Dim sDirName As String
        
        If sType = "" Then _
            sType = InputBox("Сведения СИ не опознаны. Введите ключевое слово СИ:", "Ввод данных"): If sType = "" Then Exit Sub
        
        sShortCusName = ReplaceBadSymbols(sShortCusName): sType = ReplaceBadSymbols(sType) 'убрать запрещённые в имени файла символы
'        sDocumentName = Replace(sDocumentName, "pr_", "sv_")
'        sDocumentName = Replace(sDocumentName, "jr_", "sv_")
'        sDocumentName = Replace(sDocumentName, "prc_", "srt_")
'        sDocumentName = Replace(sDocumentName, "jrc_", "srt_")
        
        If InStr(sDocumentName, "prm_") > 0 Then sDocumentName = "ИЗМЕРЕНИЕ"

        sDirName = Format(sLastSaveDate, "yyyy.mm") & " " & _
                    sType & " -- " & _
                    sShortCusName & " - " & _
                    sDocumentName
        
        If Not sDocumentName = "ИЗМЕРЕНИЕ" Then
            
            If InStr(sDocumentName, "prc_") > 0 Then _
                sDirName = sDirName & " srt_"
                
            If InStr(sDocumentName, "pr_") > 0 Then _
                sDirName = sDirName & " sv_"
        End If
        
        Dim currfgisSvNum As String
        currfgisSvNum = ActiveWorkbook.BuiltinDocumentProperties("Subject")
        
        If currfgisSvNum <> vbNullString Then _
        currfgisSvNum = Right(currfgisSvNum, InStr(StrReverse(currfgisSvNum), "/") - 1)
        
        If currfgisSvNum <> vbNullString Then _
            sDirName = sDirName & currfgisSvNum
        
        If Dir(.ArchivePath & sDirName, vbDirectory) <> "" Then
            MsgBox sDirName & vbNewLine & vbNewLine & "Каталог уже существует."
        Else
            sDirName = InputBox("Имя создаваемого каталога:", "Подтвердите ввод", sDirName)
            
            If sDirName <> "" Then
                sDirName = ReplaceBadSymbols(sDirName): sDirName = .ArchivePath & sDirName
                
                If Dir(sDirName, vbDirectory) = "" Then
                    MkSubDir .ArchivePath, sLastSaveDate, sArrDataBase(0, 0)
                    
                    MkDir sDirName
                  '  CreateShortcut sDirName
                End If
                
                Explorer.OpenFolder sDirName, True

            End If
        End If
    End With
    
    DataBase.ReCacheData
    
End Sub
'#########################################################
'процедура создаёт дополнительный каталог для удобочитаемости
Sub MkSubDir(sArchvePath As String, sLastSaveDate As String, sFirstListDirDate As String)
    
    If sFirstListDirDate = "" Then Exit Sub 'если первый элемент БД пустой
    sLastSaveDate = Format(sLastSaveDate, "yyyy.mm")
    sFirstListDirDate = Left(sFirstListDirDate, InStr(sFirstListDirDate, " ") - 1) 'выделить метку год + месяц
    
    Dim sTempDirName As String
    If Left(sLastSaveDate, 4) > Left(sFirstListDirDate, 4) Then 'если год создаваемого каталога новее
        
        sTempDirName = sArchvePath & Left(sLastSaveDate, 4) & " +++++++++++++++++++++++++++++++"
        
        If FolderNotExist(sTempDirName) Then _
            MkDir sTempDirName 'создать каталог, если его нет
        
    ElseIf Left(sLastSaveDate, 4) = Left(sFirstListDirDate, 4) Then 'один и тот же год
        
        If Right(sLastSaveDate, 2) <> Right(sFirstListDirDate, 2) Then 'месяц отличается
            sTempDirName = sArchvePath & sLastSaveDate & "  ============================="
            
            If FolderNotExist(sTempDirName) Then _
                MkDir sTempDirName 'создать каталог, если его нет
        End If
    End If
End Sub
'#########################################################
'процедура заполняет данные средства измерений на всех листах
Sub FillNameInstrument(miData As MeasInstrument)
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        
        Application.ScreenUpdating = False
        Dim ws As Worksheet, rTempCell As Range
        
        With miData
        
            For Each ws In ActiveWorkbook.Worksheets
                If FindCellRight("Наименование", rTempCell, , , , ws) Then rTempCell = .sName & " " & .sType  'наименование СИ
                If FindCellRight("ФИФ", rTempCell, , , , ws) Then rTempCell = .sFif  'ФИФ
                If FindCellRight("Методика", rTempCell, , , , ws) Then rTempCell = .sMetodic 'Методика
            Next
        End With
        Set rTempCell = Nothing: Application.ScreenUpdating = True
        
        ActiveWorkbook.Save
    #End If
End Sub
'#########################################################
'функция извлекает заводской номер из ячейки и передаёт его в обработку
Function GetSerialNumFromCell() As String
    #If CurrentAppCode = "xl" Then 'только для приложения Excel
        
        Dim rTempCell As Range, sTempStr As String
        If FindCellRight("Заводской / серийный номер:", rTempCell) Then _
            sTempStr = CStr(rTempCell): DeleteSpaceStEnd sTempStr
        
        GetSerialNumFromCell = "prm_TEMP_PROTOCOL"
        If sTempStr <> "" Then GetSerialNumFromCell = "prm_" & sTempStr
        Set rTempCell = Nothing
    #End If
End Function

'#########################################################
'дополнительная функция для связи word и excel
Sub SubExportFormatting()
'    #If CurrentAppCode = "xl" Then 'только для приложения Excel
'        ExportFormatting True
'    #End If
End Sub
'#########################################################
Sub ChangeXlPropertyComment(myInstrument As MeasInstrument)

    #If CurrentAppCode = "xl" Then 'только для приложения Excel
    
        
        With ActiveWorkbook
            If .Sheets.count = 1 Then
                If InStr(.ActiveSheet.name, myInstrument.sType) = 0 Then 'если имя модификации не включает в себя имя типа
                    SetBuiltInProperty "Comments", myInstrument.sType & " (" & .ActiveSheet.name & ")" 'имя текущего листа в книгу
                Else
                    SetBuiltInProperty "Comments", .ActiveSheet.name 'имя текущего листа в книгу
                End If
            Else
                SetBuiltInProperty "Comments", myInstrument.sType 'несколько рабочих листов
            End If
        End With
    #End If
End Sub
