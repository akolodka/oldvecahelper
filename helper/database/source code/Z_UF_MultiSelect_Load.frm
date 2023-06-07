VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_UF_MultiSelect_Load 
   Caption         =   "Доступно несколько шаблонов"
   ClientHeight    =   3555
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   5610
   OleObjectBlob   =   "Z_UF_MultiSelect_Load.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Z_UF_MultiSelect_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const sKeywordPrimary = "первичная" 'какое слово должно присутствовать в наименовании файла
Const FIF_PR_KEYWORD = "fif_"

Private myBase As New Z_clsmBase, WorkClsm As New Z_clsmSearch
Dim dbInstumentsDir As String, sArrKeyCode() As String
Private isFormLoaded As Boolean

Private Sub UserForm_Initialize()
    
    Dim verificationType As String, _
        rTempCell As Range, _
        TemplateFileName As String
    
    If FindCellRight("Вид поверки", rTempCell) Then _
        verificationType = CStr(rTempCell): _
        DeleteSpaceStEnd verificationType 'тип поверки по открытому протоколу

    With myMultiSel 'пакет переменных мультивыбора
        dbInstumentsDir = Dir(.sMSBaseDir & "*" & .sMSfifNum & "*", vbDirectory)  'директория для выбора файлов
        
        If UMenu.typе <> instrumentsOLD Then 'если идёт работа с формой создания свидетельства
            
            Me.chbPrimary.Enabled = True
            If verificationType = "первичная" Then _
                Me.chbPrimary = True
                
            .sMSAdditionalText = GetAdditionTextFromComment
            
        Else 'загрузка шаблона протокола
            
            If .bolThisWBSheetLoad Then 'если загрузка идёт в текущую книгу
            
                If verificationType = "первичная" Then _
                    Me.chbPrimary = True 'если опознано, что ныне загружен протокол первичной поверки
                    
            Else
                If dbInstumentsDir <> vbNullString Then 'директория с шаблонами была опознана
                    TemplateFileName = .sMSBaseDir & dbInstumentsDir & "\*" & .sMSfifNum & "*первичная*" & "*.xls*"
            
                    If Dir(TemplateFileName) <> vbNullString Then _
                        Me.chbPrimary.Enabled = True  'если найден хоть один файл первичной поверки
                        
                End If
            End If
            
        End If
        
    End With

    FillArrKeycodeFromFile 'заполнить массив кейкодов
    UpdateListBoxResults 'обновить содержимое листбокса
    
    isFormLoaded = True

End Sub
Private Sub UserForm_Terminate()
    myMultiSel.sMSAdditionalText = ""
    
    Set myBase = Nothing
    Set WorkClsm = Nothing
    
    If UMenu.typе = instrumentsOLD Then _
        Z_UF_Search.tboxSearch.SetFocus
        
End Sub
Private Sub UserForm_Activate()
    
    If UMenu.typе <> instrumentsOLD Then
        
        With Me.tboxSearch
            
            .text = GetAdditionTextFromComment & " " 'только для свидетельства
            
            If Me.listFiles.List(0) = vbNullString Then _
                .text = vbNullString
        End With
    End If
    
    If Me.listFiles.ListCount = 1 Then
        
        If Me.listFiles.List(0) <> vbNullString Then _
            InitiateLoad
    
    End If
        

End Sub
'#########################################################
'функция обращается к свойству документа для извлечения комментария
Function GetAdditionTextFromComment() As String
    
    Dim ws As Worksheet, _
        sAdditionalText As String, _
        sSearchArray() As String, _
        stringToSearch As String
        
    For Each ws In ActiveWorkbook.Worksheets
    
        sAdditionalText = ws.name
        sSearchArray = Split(sAdditionalText, "-")
        
        If stringToSearch <> vbNullString Then _
            stringToSearch = stringToSearch & " "
        
        Dim i As Byte
        For i = LBound(sSearchArray) To UBound(sSearchArray)
            stringToSearch = stringToSearch & sSearchArray(i)
            If i < UBound(sSearchArray) Then stringToSearch = stringToSearch & " -"
        Next i
    Next ws
    
    GetAdditionTextFromComment = stringToSearch 'дополнительная строка поиска
    Set ws = Nothing
End Function
'#########################################################
'процедура заполняет массив кейкодов
Sub FillArrKeycodeFromFile()
    
    Const fileName As String = "keycode.npDb"
     
    Dim fso As New FileSystemObject, _
        charTablePath As String
        
    charTablePath = fso.BuildPath(Config.sourceDataPath, fileName)
    
    If fso.FileExists(charTablePath) Then 'если опознано наличие в каталоге надстройки файла кейкодов
        
        sArrKeyCode = WorkClsm.FillDataBase( _
            myBase.GetArrFF(charTablePath))  'получить массив кейкодов(если файл обнаружен)

    Else 'если файл не был обнаружен и загружен
        ReDim sArrKeyCode(0, 1)
    End If
    
End Sub
'##################################################
'Переключатель выбора первичной поверки
Private Sub chbPrimary_Change()
    If isFormLoaded Then UpdateListBoxResults: Me.tboxSearch.SetFocus
    
    TrueElementForeColor Me.chbPrimary, True 'вернуть чёрный цвет
    If Me.chbPrimary Then TrueElementForeColor Me.chbPrimary 'когда выбран переключатель
End Sub
Private Sub tboxSearch_Change()
    
    If isFormLoaded Then _
        UpdateListBoxResults

    If Me.listFiles.List(0) = vbNullString Then
        
        With Me.tboxSearch
            .SelStart = 0
            .SelLength = Len(.text)
        End With
    End If
End Sub
'##################################################
'Процедура обновляет список для загрузки данных
Sub UpdateListBoxResults()

    Dim sSearchFileMask As String, _
        TemplateFileName As String, _
        sArrTemp() As String, _
        i As Byte, _
        sArrFind() As String
        
    Me.listFiles.Clear
    Me.cmbLoad.Enabled = False
    
    sSearchFileMask = myMultiSel.sMSBaseDir _
                    & dbInstumentsDir & "\*" & myMultiSel.sMSfifNum & "*"  'маска файлов для поиска

    TemplateFileName = sSearchFileMask & "*.xls*" 'загрузка протокола поверки
   ' If UMenu.typе = wdProject Then TemplateFileName = sSearchFileMask & "*.doc*"  'загрузка проекта свидетельства
    
    TemplateFileName = Dir(TemplateFileName)
    
'    If Dir(templateFileName) = vbNullString Then _
'        templateFileName = vbNullString
'
'    If Dir(templateFileName) <> vbNullString Then _
'        templateFileName = Dir(templateFileName)
    
    ReDim sArrFind(1, 0) 'получить массив всех найденных файлов
    Do While TemplateFileName <> vbNullString
        
        If Not TemplateFileName Like FIF_PR_KEYWORD & "*" Then
        
            If Me.chbPrimary Then
                If InStr(TemplateFileName, sKeywordPrimary) > 0 Then _
                    FillsArrTemp sArrFind, sArrTemp, TemplateFileName
            Else
                If InStr(TemplateFileName, sKeywordPrimary) = 0 Then _
                    FillsArrTemp sArrFind, sArrTemp, TemplateFileName
            End If
            
        End If
        
        TemplateFileName = Dir
    Loop
    SortMassBiv sArrFind 'сортировать одномерный массив методом вставки
    
    If Me.tboxSearch = vbNullString Then 'заполнить листбокс найденными значениями
        
        For i = LBound(sArrFind, 2) To UBound(sArrFind, 2)
            Me.listFiles.AddItem sArrFind(UBound(sArrFind), i)
        Next
        
    Else
        Dim sTempArr() As String
        ReDim sTempArr(UBound(sArrFind), 0) 'выделить память под временный массив
        
        If sArrKeyCode(0, 0) = vbNullString Then 'массив кейкодов не был загружен
            FindInBivArr sArrFind, sTempArr, Me.tboxSearch  'получить массив поисковых совпадений
        Else ' массив кейкодов был загружен
            Dim sInputRus As String, sInputEng As String, sInputString As String
            FillInputData sArrKeyCode, Me.tboxSearch, sInputRus, sInputEng  'получить значения для поиска по массиву
            sInputString = sInputRus & " " & sInputEng: FindInBivArr sArrFind, sTempArr, sInputString  'получить массив поисковых совпадений значениями на русском языке
        End If
        
        For i = LBound(sTempArr, 2) To UBound(sTempArr, 2) 'заполнить листбокс найденными значениями
            Me.listFiles.AddItem sTempArr(UBound(sArrFind), i)
        Next
    End If
    
    If UBound(Me.listFiles.List) >= 0 Then
    
        If Me.listFiles.List(0) <> vbNullString Then _
            Me.listFiles.Selected(0) = True: _
            Me.cmbLoad.Enabled = True
        
    End If
End Sub

'##################################################
'Процедура добавляет элемент в массив
Sub FillsArrTemp(ByRef sArrFind() As String, sArrTemp() As String, TemplateFileName As String)
    sArrTemp = Split(TemplateFileName, myMultiSel.sMSfifNum) 'разбить строку на номере в фиф
    
    If sArrFind(LBound(sArrFind), UBound(sArrFind, 2)) <> "" Then ReDim Preserve sArrFind(UBound(sArrFind), UBound(sArrFind, 2) + 1) 'расширить массив
    
    DeleteSpaceStEnd sArrTemp(UBound(sArrTemp))
    sArrFind(LBound(sArrFind), UBound(sArrFind, 2)) = sArrTemp(UBound(sArrTemp)) 'поместить полное наименовение в первую ячейку массива
    If InStr(sArrTemp(UBound(sArrTemp)), "#!") > 0 Then sArrTemp = Split(sArrTemp(UBound(sArrTemp)), "#!")
    sArrFind(UBound(sArrFind), UBound(sArrFind, 2)) = sArrTemp(UBound(sArrTemp)) 'сокращённое наименование, если имеется
End Sub
'##################################################
'передать полное наименование выбранного файла в переменную
Private Sub cmbLoad_Click()
    
    InitiateLoad

End Sub
    Private Sub InitiateLoad()
    
        With myMultiSel
            
            Dim TemplateFileName As String
            
            If Me.listFiles.ListIndex >= 0 Then _
                TemplateFileName = Dir(.sMSBaseDir & dbInstumentsDir & "\*" & .sMSfifNum & _
                    "*" & Me.listFiles.List(Me.listFiles.ListIndex))
            
            If UMenu.typе <> instrumentsOLD Then 'загрузка шаблона word
            
                If Me.chbPrimary And InStr(TemplateFileName, sKeywordPrimary) = 0 Then
                    
                    If TemplateFileName <> vbNullString Then _
                        TemplateFileName = Dir 'поиск шаблона word для первичной поверки
                End If
                
            End If
            
            If TemplateFileName <> vbNullString Then _
                .sMSFullFileName = .sMSBaseDir & dbInstumentsDir & "\" & TemplateFileName
                
        End With
        
        VBA.Unload Me
    
    End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then VBA.Unload Me 'esc
    If KeyCode = 13 Then cmbLoad_Click 'enter
End Sub
Private Sub listFiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then VBA.Unload Me 'esc
    If KeyCode = 13 Then cmbLoad_Click 'enter
    
    If KeyCode = 38 Then
        If Me.listFiles.Selected(0) = True Then
            
            With Me.tboxSearch
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        End If
    End If
End Sub
Private Sub chbPrimary_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then VBA.Unload Me 'esc
    If KeyCode = 13 Then cmbLoad_Click 'enter
End Sub
Private Sub tboxSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then VBA.Unload Me 'esc
    If KeyCode = 13 Then cmbLoad_Click 'enter
    
    If KeyCode = 40 Then
        If Me.listFiles.ListCount > 0 Then Me.listFiles.Selected(1) = True: Me.listFiles.SetFocus
    End If
End Sub
Private Sub cmbLoad_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then VBA.Unload Me 'esc
    If KeyCode = 13 Then cmbLoad_Click 'enter
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
