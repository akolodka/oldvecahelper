VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_NumName_Config 
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3840
   OleObjectBlob   =   "UF_NumName_Config.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_NumName_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SELSTART_POSITION = 0
Const SHIFT_KEY = 1

Const ENRIVON_KEY = "USERNAME"
Const DEFAULT_DEPARTMENT = 210

Const LABEL_LENGTH = 32

Private isEntered As Boolean, _
        isSaved As Boolean
        
Private fso As New FileSystemObject

Private Sub UserForm_Initialize()

    Me.tboxDepartment = Replace(Config.numDepartment, Base.defaultValue, DEFAULT_DEPARTMENT)
    Me.tboxLaboratory = Replace(Config.numLaboratory, Base.defaultValue, DEFAULT_DEPARTMENT)

    Me.labelJournalPath = Replace(Config.journalPath, Base.defaultValue, Base.defaultValuePrinted)

    Me.tboxEmployeeMajor = Replace(Config.empMajor, Base.defaultValue, Environ(ENRIVON_KEY))
    Me.tboxEmployeeSecond = Replace(Config.empSecond, Base.defaultValue, Base.defaultValuePrinted)
    Me.tboxEmployeeThird = Replace(Config.empThird, Base.defaultValue, Base.defaultValuePrinted)
    
    SetEventControls Me
    CheckJournalPath
    
    Me.versionLabel.Caption = Addin.VersionCaption
    
    isSaved = True
    
End Sub
    Private Sub CheckJournalPath( _
        )
    
        Me.btnOpenJournalDir.Enabled = False
        ColorText Me.labelJournalPath, Colors.red
        Me.btnChooseJournal.BackColor = Colors.yellow
        Me.labelJournalPath.TextAlign = fmTextAlignCenter
    
        If Config.journalPath = Base.defaultValue Then _
            Exit Sub
        
        Me.labelJournalPath = ShortedString( _
                                            strData:=Config.journalPath, _
                                            maxLength:=LABEL_LENGTH _
                                            )
               
        If fso.FileExists(Config.journalPath) Then
               
            ColorText Me.labelJournalPath
            Me.btnChooseJournal.BackColor = Colors.white
            Me.labelJournalPath.ForeColor = Colors.green
            Me.labelJournalPath.TextAlign = fmTextAlignLeft
            
            Me.btnOpenJournalDir.Enabled = True
            
        End If

        If UMenu.isLoaded Then _
            NotifyAboutSaving
    
    End Sub
        '##########################################################################
        'функци€ возвращает текстовую строку нужной длины в зависимости от установленной границы
        Private Function ShortedString( _
            strData As String, _
            maxLength As Byte _
            ) As String
            
            ShortedString = " " & strData ' по умолчанию возвращать всю строку 'название пути полностью умещаетс€
            
            Dim leftPart As String, _
                rightPart As String, _
                iStrStart As Integer
                
            If Len(strData) > maxLength Then 'пусть не умещаетс€ полностью
                
                iStrStart = InStr(strData, "\")
                If iStrStart = 0 Then _
                    Exit Function
                
                If iStrStart + 2 <= Len(strData) Then _
                    leftPart = Left(strData, InStr(iStrStart + 2, strData, "\")) 'об€зательна€ лева€ часть строки
                
                If maxLength - Len(leftPart) - 5 >= 0 Then _
                    rightPart = Right(strData, maxLength - Len(leftPart) - 5) ' чтобы не было ошибки в случае короткой строки
        
                ShortedString = " " & leftPart & " ... " & rightPart
            End If
        End Function
        '=================================================
        Private Sub NotifyAboutSaving()
            
            If Not UMenu.isLoaded Then Exit Sub
            
            ShowSaveButton Me.cmbSaveReady
            isSaved = False
            
        End Sub
            Private Sub ShowSaveButton( _
                btnName As Object _
                )
                
                btnName.Font.Size = 11
                btnName.Caption = "—охранить"
                btnName.BackColor = Colors.yellow
                
            End Sub
Private Sub UserForm_Activate()
    Me.tboxLaboratory.SetFocus
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If Not isSaved Then
        
        Dim ans As Integer
        ans = Handler.ask("—охранить изменени€ перед выходом?")
        
        If ans = vbYes Then _
            Config.Save
    
    End If
    
End Sub
'=================================================

'Private Sub tboxLaboratory_Enter()
'
'    If Me.tboxLaboratory = Me.tboxDepartment Then _
'        Me.tboxLaboratory.SelStart = Len(Me.tboxLaboratory)
'
'End Sub
Private Sub tboxLaboratory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then
        
        SelectTextField _
                        oTextBox:=Me.tboxLaboratory, _
                        selectEnd:=True
                        
        isEntered = True
        
    End If
    
End Sub

Private Sub tboxLaboratory_Change()
    
    ColorText Me.tboxLaboratory
    If Me.tboxLaboratory = Me.tboxDepartment Then _
        ColorText Me.tboxLaboratory, Colors.red
    
    Config.numLaboratory = Me.tboxLaboratory
    NotifyAboutSaving
    
End Sub
    Private Sub ColorText( _
        ufObject As Object, _
        Optional color As Colors = Colors.turquoise _
        )
        
        ufObject.ForeColor = color
        
    End Sub
        
Private Sub tboxLaboratory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then _
        SaveData
    
End Sub
Private Sub tboxLaboratory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub tboxDepartment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then
        
        SelectTextField Me.tboxDepartment
        isEntered = True
        
    End If
    
End Sub
    Private Sub SelectTextField( _
        oTextBox As Object, _
        Optional selectEnd As Boolean = False _
        )
        
        oTextBox.selStart = SELSTART_POSITION
        
        If selectEnd Then _
            oTextBox.selStart = Len(oTextBox.value)
        
        oTextBox.SelLength = Len(oTextBox.value)

    End Sub
Private Sub tboxDepartment_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then _
        SaveData
    
End Sub
Private Sub tboxDepartment_Change()
    
    Config.numDepartment = Me.tboxDepartment
    
    ChangeLaboratoryNumber
    NotifyAboutSaving
    
End Sub
    Private Sub ChangeLaboratoryNumber()
        
        If Not UMenu.isLoaded Then Exit Sub
        Me.tboxLaboratory = Me.tboxDepartment
        
    End Sub
Private Sub tboxDepartment_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub tboxEmployeeMajor_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then
        
        SelectTextField Me.tboxEmployeeMajor
        isEntered = True
        
    End If
    
End Sub
' ----------------------------------------------------------------
' ƒата: 19.02.2023 15:34
' Ќазначение:
' ----------------------------------------------------------------
Private Sub tboxEmployeeMajor_Change()
    
    ColorText Me.tboxEmployeeMajor
    
    If Me.tboxEmployeeMajor = Base.defaultValuePrinted Then _
        ColorText Me.tboxEmployeeMajor, Colors.red
    
    Config.empMajor = Base.defaultValue
    
    If Me.tboxEmployeeMajor <> vbNullString Then _
        Config.empMajor = Me.tboxEmployeeMajor
        
    NotifyAboutSaving
    
End Sub
Private Sub tboxEmployeeMajor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then _
        SaveData
    
End Sub
Private Sub tboxEmployeeMajor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub tboxEmployeeSecond_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then
        
        SelectTextField Me.tboxEmployeeSecond
        isEntered = True
        
    End If
    
End Sub
Private Sub tboxEmployeeSecond_Change()
    
    ColorText Me.tboxEmployeeSecond
    
    If Me.tboxEmployeeSecond = Base.defaultValuePrinted Then _
        ColorText Me.tboxEmployeeSecond, Colors.red
    
    Config.empSecond = Base.defaultValue
    
    If Me.tboxEmployeeSecond <> vbNullString Then _
        Config.empSecond = Me.tboxEmployeeSecond
        
    NotifyAboutSaving
    
End Sub
Private Sub tboxEmployeeSecond_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then _
        SaveData
    
End Sub
Private Sub tboxEmployeeSecond_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub tboxEmployeeThird_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then
        
        SelectTextField Me.tboxEmployeeThird
        isEntered = True
        
    End If
    
End Sub
Private Sub tboxEmployeeThird_Change()

    ColorText Me.tboxEmployeeThird
    
    If Me.tboxEmployeeThird = Base.defaultValuePrinted Then _
        ColorText Me.tboxEmployeeThird, Colors.red
    
    Config.empThird = Base.defaultValue
    
    If Me.tboxEmployeeThird <> vbNullString Then _
        Config.empThird = Me.tboxEmployeeThird
        
    NotifyAboutSaving
    
End Sub
Private Sub tboxEmployeeThird_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then _
        SaveData
    
End Sub
Private Sub tboxEmployeeThird_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub frEmployee_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
Private Sub frPrefix_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub

'=================================================
Private Sub cmbSaveReady_Click()
    
    If isSaved Then UMenu.Unload
    SaveData
    
End Sub
    Private Sub SaveData( _
        )
        
        If isSaved Then Exit Sub
            
        ShowReadyButton Me.cmbSaveReady
        isSaved = True
        
        Config.Save
            
    End Sub
        Private Sub ShowReadyButton( _
            btnName As Object _
            )
            
            btnName.Font.Size = 12
            btnName.Caption = "√отово"
            btnName.BackColor = Colors.white
            
        End Sub
'=================================================
Private Sub btnChooseJournal_MouseUp( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)
        
    Dim status As Boolean
    
    status = Config.ChooseJournalPath( _
        initialSomniumPath:=Shift = SHIFT_KEY _
        )
    If status Then _
        CheckJournalPath
    
End Sub
'=================================================
Private Sub btnOpenJournalDir_Click()
    Explorer.OpenFolder _
        path:=fso.GetParentFolderName(Config.journalPath), _
        isFocusOnWindow:=True
End Sub

Private Sub UserForm_Terminate()
    UMenu.isLoaded = False
End Sub

