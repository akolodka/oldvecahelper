VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatabaseConfig 
   Caption         =   "Настройки источников данных"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5805
   OleObjectBlob   =   "frmDatabaseConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatabaseConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SHIFT_KEY = 1
Const MAX_LENGTH = 42
Const SHORT_STRING_PREFIX = " "

Const COMBOX_DEFAULT = "< выбрать >"

Const SAVE_CAPTION = "Сохранить"
Const SAVE_FONTSIZE = 11

Const READY_CAPTION = "Готово"
Const READY_FONTSIZE = 12

Private fso As New FileSystemObject, _
        isSaved As Boolean, _
        status As Boolean
        
Private Sub UserForm_Initialize()
    
    SetEventControls Me
    
    CheckSourcePath
    CheckSandboxPath
    CheckArchivePath
    
    CheckCharTable
    FillComboxes
    
    Me.labelVersion = Addin.VersionCaption
    
    isSaved = True
    
End Sub
    Private Sub CheckSourcePath( _
        )
        
        Me.btnOpenSourcePath.Enabled = False
        Me.btnChooseSourcePath.BackColor = Colors.yellowPastel
        
        If Config.sourceDataPath = Base.defaultValue Then _
            Exit Sub
            
        FillPathLabel Me.labelBasePath.name, Config.sourceDataPath
        
        If fso.FolderExists(Config.sourceDataPath) Then
        
            Me.btnChooseSourcePath.BackColor = Colors.white
            Me.btnOpenSourcePath.Enabled = True
            
        End If
        
        PaintApplyServerButton
        
        If UMenu.isLoaded Then _
            NotifyAboutSaving
        
    End Sub
        Private Sub PaintApplyServerButton()
        
            Me.btnApplyServerConfig.Enabled = True
            Me.btnApplyServerConfig.BackColor = Colors.yellowPastel
            
            'buildPath, потому что один путь имеет на конце слеш
            If fso.BuildPath(Config.sourceDataPath, 1) = fso.BuildPath(Base.serverProgramDataPath, 1) Then
            
                Me.btnApplyServerConfig.Enabled = False
                Me.btnApplyServerConfig.BackColor = Colors.white
                
            End If
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 07.03.2023 22:20
        ' Назначение: обработка поведения контрола
        ' ----------------------------------------------------------------
        Private Sub FillPathLabel( _
            controlName As String, _
            path As String _
            )
            
            Me.Controls(controlName).TextAlign = fmTextAlignCenter
            
            If path = Base.defaultValue Then _
                Exit Sub
                
    '        Me.Controls(controlName).ForeColor = Colors.red
    '        Me.Controls(controlName).Caption = Base.defaultValuePrinted
            
            Me.Controls(controlName).caption = ShortedString(path, MAX_LENGTH)
            Me.Controls(controlName).TextAlign = fmTextAlignLeft
            
            Dim targetColor As Long
            targetColor = Colors.turquoise
            
            If Not fso.FolderExists(path) Then _
                targetColor = Colors.red
                
            Me.Controls(controlName).foreColor = targetColor
        
        End Sub
            ' ----------------------------------------------------------------
            ' Дата: 07.03.2023 22:21
            ' Назначение: Умная обрезка строки
            ' ----------------------------------------------------------------
            Private Function ShortedString( _
                stringData As String, _
                maxLength As Byte _
                ) As String
                
                ShortedString = SHORT_STRING_PREFIX & stringData ' по умолчанию возвращать всю строку 'название пути полностью умещается
                If Len(stringData) <= maxLength Then _
                    Exit Function
                
                Dim instrSeparator As Integer
                instrSeparator = InStr(stringData, Application.PathSeparator)
                
                If Not CBool(instrSeparator) Then _
                    Exit Function 'если не обнаружен разделитель каталога
                    
                Dim leftPart As String
                If instrSeparator + 2 <= Len(stringData) Then _
                    leftPart = Left(stringData, InStr(instrSeparator + 2, stringData, Application.PathSeparator)) 'обязательная левая часть строки
                
                Dim rightPart As String
                If maxLength - Len(leftPart) - 5 >= 0 Then _
                    rightPart = Right(stringData, maxLength - Len(leftPart) - 5) ' чтобы не было ошибки в случае короткой строки
                
                ShortedString = SHORT_STRING_PREFIX & leftPart & " ... " & rightPart
                
    
            End Function
        ' ----------------------------------------------------------------
        Private Sub NotifyAboutSaving()
            
            If Not UMenu.isLoaded Then _
                Exit Sub
            
            ShowSaveButton Me.btnSaveReady
            isSaved = False
            
        End Sub
            Private Sub ShowSaveButton( _
                btnName As Object _
                )
                
                btnName.Font.Size = SAVE_FONTSIZE
                btnName.caption = SAVE_CAPTION
                btnName.BackColor = Colors.yellowPastel
                
            End Sub
    ' ----------------------------------------------------------------
    Private Sub CheckSandboxPath()
        
        Me.btnOpenSandboxPath.Enabled = False
        Me.btnChooseSandboxPath.BackColor = Colors.yellowPastel
        
        If Config.sandboxPath = Base.defaultValue Then _
            Exit Sub
            
        FillPathLabel Me.labelSandboxPath.name, Config.sandboxPath
        
        If fso.FolderExists(Config.sandboxPath) Then
            
            Me.btnChooseSandboxPath.BackColor = Colors.white
            Me.btnOpenSandboxPath.Enabled = True
            
        End If
   
        If UMenu.isLoaded Then _
            NotifyAboutSaving
            
    End Sub
    ' ----------------------------------------------------------------
    Private Sub CheckArchivePath()
        
        Me.btnOpenArchiveLocalPath.Enabled = False
        Me.btnChooseArchiveLocalPath.BackColor = Colors.yellowPastel
        
        If Config.archiveLocalPath = Base.defaultValue Then _
            Exit Sub
            
        FillPathLabel Me.labelArchiveLocalPath.name, Config.archiveLocalPath
        
        If fso.FolderExists(Config.archiveLocalPath) Then
        
            Me.btnChooseArchiveLocalPath.BackColor = Colors.white
            Me.btnOpenArchiveLocalPath.Enabled = True
        
        End If
        
        If UMenu.isLoaded Then _
            NotifyAboutSaving
        
    End Sub
    ' ----------------------------------------------------------------
    Private Sub CheckCharTable()
    
        Me.labelCharTable.caption = "не работает"
        Me.labelCharTable.foreColor = Colors.red
        
        Dim ct As New cCharTable
        If Not ct.isAvaiable Then _
            Exit Sub
        
        Me.labelCharTable.caption = "работает"
        Me.labelCharTable.foreColor = Colors.green

    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 13.03.2023 20:50, 16.04.2023 17:13
    ' ----------------------------------------------------------------
    Private Sub FillComboxes()
        
        Dim personsData As New Collection
        Set personsData = DataBase.VniimPersons

        Dim isFilled As Boolean
        ' ----------------------------------------------------------------
        isFilled = FillComboxList(personsData, Me.comboxVerifier)

        If isFilled Then
        
            Me.labelVerifier.Enabled = True
            
            FillComboxValue _
                value:=Config.verifierKey, _
                combox:=Me.comboxVerifier
            
        End If
        ' ----------------------------------------------------------------
        isFilled = FillComboxList(personsData, Me.comboxExecutor)

        If isFilled Then _
            Me.labelExecutor.Enabled = True

        FillComboxValue _
            value:=Config.executorKey, _
            combox:=Me.comboxExecutor
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 15.03.2023 09:37
        ' ----------------------------------------------------------------
        Private Function FillComboxList( _
            personsData As Collection, _
            combox As ComboBox _
            ) As Boolean

            FillComboxList = False

            If Not CBool(personsData.count) Then _
                Exit Function

            combox.Enabled = True
            combox.AddItem COMBOX_DEFAULT
            combox.text = combox.List(LBound(combox.List))

            Dim i As Integer
            For i = 1 To personsData.count
                combox.AddItem Trim(personsData(i))
            Next

            FillComboxList = True

        End Function
        ' ----------------------------------------------------------------
        ' Дата: 16.04.2023 18:20
        ' ----------------------------------------------------------------
        Private Sub FillComboxValue( _
            value As String, _
            combox As ComboBox _
            )
        
            Dim result As String
            result = Replace(value, Base.defaultValue, COMBOX_DEFAULT)

            If result = COMBOX_DEFAULT Then _
                Exit Sub
            ' ----------------------------------------------------------------
            DataBase.FilterCache result
            If Not CBool(DataBase.ResultFilterCount) Then _
                Exit Sub
            ' ----------------------------------------------------------------
            Dim item As New cItemPerson
            Set item = DataBase.GetItem(DataBase.ResultFilter(1).key)
            
            Dim resultValue As String
            resultValue = item.lastName & " " & item.firstName & " " & item.middleName
            
            combox.value = resultValue
        
        End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If Not isSaved Then
        
        Dim ans As Integer
        ans = Handler.ask("Сохранить изменения перед выходом?")
        
        If ans = vbYes Then _
            Config.Save
    
    End If
    
End Sub
Private Sub UserForm_Terminate()
    UMenu.isLoaded = False
End Sub
' ----------------------------------------------------------------
Private Sub btnOpenSourcePath_Click()
    Explorer.OpenFolder Config.sourceDataPath, True
End Sub
Private Sub btnChooseSourcePath_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    status = Config.ChooseFolderPath( _
        folderType:=sourceDir, _
        initialSomniumPath:=Shift = SHIFT_KEY _
        )
    
    If status Then _
        CheckSourcePath

End Sub
Private Sub btnChooseSourcePath_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    If Not KeyCode = vbKeyReturn Then _
        Exit Sub
    
    status = Config.ChooseFolderPath( _
        folderType:=sourceDir, _
        initialSomniumPath:=Shift = SHIFT_KEY _
        )
    
    If status Then _
        CheckSourcePath

End Sub
' ----------------------------------------------------------------
Private Sub btnOpenSandboxPath_Click()
    Explorer.OpenFolder Config.sandboxPath, True
End Sub
Private Sub btnChooseSandboxPath_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    status = Config.ChooseFolderPath( _
        folderType:=sandboxDir, _
        initialSomniumPath:=False _
        )
        
    If status Then _
        CheckSandboxPath

End Sub
Private Sub btnChooseSandboxPath_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    If Not KeyCode = vbKeyReturn Then _
        Exit Sub

    status = Config.ChooseFolderPath( _
        folderType:=sandboxDir, _
        initialSomniumPath:=False _
        )
        
    If status Then _
        CheckSandboxPath

End Sub
' ----------------------------------------------------------------
Private Sub btnOpenArchiveLocalPath_Click()
    Explorer.OpenFolder Config.archiveLocalPath, True
End Sub

Private Sub btnChooseArchiveLocalPath_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    status = Config.ChooseFolderPath( _
        folderType:=archiveLocalDir, _
        initialSomniumPath:=False _
        )
    
    If status Then _
        CheckArchivePath
    
End Sub
Private Sub btnChooseArchiveLocalPath_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    If Not KeyCode = vbKeyReturn Then _
        Exit Sub
    
    status = Config.ChooseFolderPath( _
        folderType:=archiveLocalDir, _
        initialSomniumPath:=False _
        )
    
    If status Then _
        CheckArchivePath

End Sub
' ----------------------------------------------------------------
Private Sub comboxVerifier_Change()
        
    PaintCombobox Me.comboxVerifier
    
    If UMenu.isLoaded Then
        
        Config.verifierKey = Base.defaultValue
        ' ----------------------------------------------------------------
        DataBase.FilterCache Me.comboxVerifier
        If Not CBool(DataBase.ResultFilterCount) Then _
            Exit Sub
        ' ----------------------------------------------------------------
        Config.verifierKey = DataBase.ResultFilter(1).key
        NotifyAboutSaving

    End If
    
End Sub
    Private Sub PaintCombobox( _
        cb As ComboBox _
        )
        
        cb.BackColor = Colors.yellowPastel
        If cb.value <> COMBOX_DEFAULT Then _
            cb.BackColor = Colors.white
    
    End Sub
' ----------------------------------------------------------------
Private Sub comboxExecutor_Change()

    PaintCombobox Me.comboxExecutor
    
    If UMenu.isLoaded Then
        
        Config.executorKey = Base.defaultValue
        ' ----------------------------------------------------------------
        DataBase.FilterCache Me.comboxExecutor
        If Not CBool(DataBase.ResultFilterCount) Then _
            Exit Sub
        ' ----------------------------------------------------------------
        Config.executorKey = DataBase.ResultFilter(1).key
        NotifyAboutSaving
        
    End If
End Sub
Private Sub btnApplyServerConfig_Click()
    
    Dim ans As Integer
    ans = Handler.ask("Применить настройки для работы с Somnium?")
    
    If ans = vbNo Then _
        Exit Sub
    
    Config.sourceDataPath = Base.serverProgramDataPath
    Config.sandboxPath = Base.desktopPath
    
    isSaved = False
    SaveData
    
    DataBase.Initialize persons
    UserForm_Initialize
    
    ShowReadyButton Me.btnSaveReady
    
End Sub
' ----------------------------------------------------------------
Private Sub btnSaveReady_Click()
    
    If isSaved Then _
        UMenu.Unload
        
    SaveData
    
End Sub
    Private Sub SaveData( _
        )
        
        If isSaved Then _
            Exit Sub
        
        Config.Save
        
        ShowReadyButton Me.btnSaveReady
        isSaved = True
            
    End Sub
        Private Sub ShowReadyButton( _
            btnName As Object _
            )
            
            btnName.Font.Size = READY_FONTSIZE
            btnName.caption = READY_CAPTION
            btnName.BackColor = Colors.white
            
        End Sub
