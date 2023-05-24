VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_NumName_Main 
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3510
   OleObjectBlob   =   "UF_NumName_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_NumName_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'запрет на использование неявных переменных

Const SELSTART_POSITION = 0
Const ASCII_LEFT = 48
Const ASCII_RIGHT = 57

Enum EHeigt_Shifts
    
    MS_2013 = 95
    MS_2016 = 95
    
End Enum

Private fso As New FileSystemObject
Private isEntered As Boolean
' ----------------------------------------------------------------
' Наименование: UserForm_Initialize (Private Constructor (Initialize))
' Назначение: Инициализация
' Дата: 17.02.2023 23:04
' ----------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    HideNonUsefulControls
    
    If IsPathAvailable Then _
        Me.cmbOpenThisPath.Enabled = True
    
    Me.tboxPrefix = NumName.numberPrefix
    Me.tboxNum = NumName.numberCore
                    
    Me.tboxNum.selStart = SELSTART_POSITION
    Me.tboxNum.SelLength = Len(Me.tboxNum)
            
    If NumName.numberCore <> vbNullString Then _
        Me.cmbSendToFile.Enabled = True
        
    Me.tboxYear = NumName.numberSuffix
    
    Me.tboxEmpOther = Base.defaultValuePrinted
    If Cache.empOther <> Base.defaultValue Then _
        Me.tboxEmpOther = Cache.empOther
    
    SetCheckBoxes
    SetEventControls Me 'инициировать групповые события для всех контролов
    
    Me.Caption = NumName.ProtocolFileMask
    
End Sub
    Private Sub SetCheckBoxes( _
        )
    
        Me.cheboxSaveAsCopy = Cache.saveAsCopy
        Me.cheboxEmpMajor.Caption = Config.empMajor
        Me.cheboxEmpOther = Cache.useEmpOther
        
        FillCheckBox Me.cheboxEmpSecond, Config.empSecond, Cache.useEmpSecond
        FillCheckBox Me.cheboxEmpThird, Config.empThird, Cache.useEmpThird
        
    End Sub
        Private Sub FillCheckBox( _
            objCheckBox As Object, _
            configValue As String, _
            configBool As Boolean _
            )
            
            objCheckBox.Caption = Base.defaultValuePrinted
            If configValue <> Base.defaultValue Then
            
                objCheckBox.Enabled = True
                objCheckBox.Caption = configValue
                objCheckBox = configBool
                
            End If
            
        End Sub
        Private Sub HideNonUsefulControls()
            
            If Protocol.typeDoc = verifying Then
            
                If fso.FileExists(Config.journalPath) Then _
                    Exit Sub
            
            End If
                       
            Me.Height = Me.Height - HeightShift
            
            Me.cmbGetNum.Visible = False
            Me.cmbOpenThisPath.Width = Me.cmbOpenThisPath.Width * 1.5
            
            Me.cmbSendToFile.Left = Me.cmbGetNum.Left + Me.cmbSendToFile.Width / 2
            Me.cmbSendToFile.Width = Me.cmbSendToFile.Width * 1.5
            
        End Sub
            Private Function HeightShift( _
                ) As EHeigt_Shifts
                
                HeightShift = EHeigt_Shifts.MS_2013
                
                If Application.Version = "16.0" Then _
                    HeightShift = MS_2016
            
            End Function
    Private Function IsPathAvailable( _
        ) As Boolean

        IsPathAvailable = Not ActiveWorkbook.path = vbNullString
        
    End Function
' ----------------------------------------------------------------
Private Sub UserForm_Terminate()
    Cache.Save
End Sub
'=================================================
Private Sub cmbOpenThisPath_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Explorer.OpenActiveWbPath Shift:=Shift
End Sub
Private Sub cmbSendToFile_Click()
    SendData
End Sub
    Private Sub SendData()

        Protocol.SendNumberToSheet (NumName.FullNumber)
        Protocol.SendFactoryNumberToComment
        Protocol.RenameFile NumName.ProtocolFileMask, Cache.saveAsCopy
        
        UMenu.Unload
    
    End Sub
'=================================================
Private Sub tboxPrefix_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then _
        SelectTextField Me.tboxPrefix, selStart:=3: _
        isEntered = True
    
End Sub
Private Sub tboxPrefix_Change()
    
    If Not UMenu.isLoaded Then Exit Sub
    NumName.numberPrefix = Me.tboxPrefix
    Me.Caption = NumName.ProtocolFileMask
    
End Sub
Private Sub tboxPrefix_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = PressedKeyAsc(KeyAscii)
End Sub
Private Sub tboxPrefix_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub tboxNum_Change()
    
    If Not UMenu.isLoaded Then Exit Sub
    
    NumName.numberCore = Me.tboxNum
    Me.cmbSendToFile.Enabled = Me.tboxNum <> vbNullString
    
    Me.Caption = NumName.ProtocolFileMask
    
End Sub
Private Sub tboxNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = PressedKeyAsc(KeyAscii)
End Sub
' ----------------------------------------------------------------
' Дата: 18.02.2023 15:10
' Назначение: обработка нажатий
'    параметр KeyCode: любое
'    параметр Shift: любое
' ----------------------------------------------------------------
Private Sub tboxNum_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
'    If KeyCode = vbKeyReturn Then _
'        KeyCode = SELSTART_POSITION: _
'        ReturnPressed Shift

    If KeyCode = vbKeyAdd Then _
        KeyCode = SELSTART_POSITION: _
        AddPressed
    
End Sub
'    ' ----------------------------------------------------------------
'    ' Дата: 18.02.2023 15:28
'    ' Назначение: обработка нажатия Ctrl+Enter или Shift+Enter
'    '    параметр Shift:
'    ' ----------------------------------------------------------------
'    Private Sub ReturnPressed( _
'        ByVal Shift As Integer _
'        )
'
'        If Shift = vbKeyRButton Or Shift = vbKeyLButton Then 'Ctrl + Enter or Shift + Enter
'
'            Dim statusOk As Boolean
'            statusOk = NumName.GetFromJournal
'
'            If Not statusOk Then _
'                Exit Sub
'
'        End If
'
'        SendData
'
'    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 18.02.2023 15:28
    ' Назначение: обработка нажатие Num+ (NumAdd)
    ' ----------------------------------------------------------------
    Private Sub AddPressed( _
        )
        
        If NumName.numberCore = vbNullString Then _
            NumName.numberCore = 0
        
        If NumName.numberCore = Base.defaultValue Then _
           NumName.numberCore = 0
            
        NumName.numberCore = CInt(NumName.numberCore) + 1
        SendData
        
    End Sub

Private Sub tboxNum_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then _
        SelectTextField Me.tboxNum: _
        isEntered = True
    
End Sub
    Private Sub SelectTextField( _
        oTextBox As Object, _
        Optional selStart As Integer = SELSTART_POSITION _
        )
        
        If selStart > Len(oTextBox.value) Then _
            selStart = SELSTART_POSITION

        oTextBox.selStart = selStart
        oTextBox.SelLength = Len(oTextBox.value)

    End Sub
Private Sub tboxNum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    UMenu.isLoaded = False
    Me.tboxNum = NumName.numberCore 'применение форматирования вида "0000"
    UMenu.isLoaded = True
    
    isEntered = False
    
End Sub
'=================================================
Private Sub tboxYear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then _
        SelectTextField Me.tboxYear, selStart:=2: _
        isEntered = True
    
End Sub
Private Sub tboxYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = PressedKeyAsc(KeyAscii)
End Sub
    Private Function PressedKeyAsc( _
        ByVal KeyAscii As Integer _
        ) As Integer
                    
        PressedKeyAsc = KeyAscii
        
        If KeyAscii < ASCII_LEFT Or KeyAscii > ASCII_RIGHT Then _
            PressedKeyAsc = SELSTART_POSITION
            
    End Function
Private Sub tboxYear_Change()
    
    If Not UMenu.isLoaded Then Exit Sub
    NumName.numberSuffix = Me.tboxYear
    Me.Caption = NumName.ProtocolFileMask

End Sub
Private Sub tboxYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
'=================================================
Private Sub cheboxEmpSecond_Change()
        
    ColorChebox cheboxEmpSecond
    
    If Not UMenu.isLoaded Then Exit Sub
    Cache.useEmpSecond = Me.cheboxEmpSecond
    
End Sub
    Private Sub ColorChebox( _
        oChebox As Object, _
        Optional color As Colors = Colors.blue _
        )
        
        oChebox.ForeColor = Colors.black
        
        If oChebox Then _
            oChebox.ForeColor = color
        
    End Sub
    
'=================================================
Private Sub cheboxEmpThird_Change()
    
    ColorChebox cheboxEmpThird
    
    If Not UMenu.isLoaded Then Exit Sub
    Cache.useEmpThird = Me.cheboxEmpThird
    
End Sub
'=================================================
Private Sub cheboxEmpOther_Change()
    
    ColorChebox cheboxEmpOther
    
    If Not UMenu.isLoaded Then Exit Sub
    Cache.useEmpOther = Me.cheboxEmpOther
    
End Sub
'=================================================

Private Sub cheboxSaveAsCopy_Change()
    
    ColorChebox cheboxSaveAsCopy
    
    If Not UMenu.isLoaded Then Exit Sub
    Cache.saveAsCopy = Me.cheboxSaveAsCopy
    
End Sub
'=================================================
Private Sub tboxEmpOther_Change()
    
    If Not UMenu.isLoaded Then Exit Sub
    Cache.empOther = Me.tboxEmpOther
    
End Sub
Private Sub tboxEmpOther_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Not isEntered Then _
        SelectTextField Me.tboxEmpOther: _
        isEntered = True
    
End Sub
Private Sub tboxEmpOther_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    isEntered = False
End Sub
' ----------------------------------------------------------------
' Дата: 19.02.2023 00:00
' Назначение:
' ----------------------------------------------------------------
Private Sub cmbGetNum_Click()
    
    Dim statusOk As Boolean
    statusOk = NumName.GetFromJournal
    
    If Not statusOk Then
    
        UF_NumName_Main.Hide
        UF_NumName_Main.Show False
        Exit Sub
        
    End If
    
    SendData

End Sub

