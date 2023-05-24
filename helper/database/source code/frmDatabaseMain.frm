VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatabaseMain 
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   OleObjectBlob   =   "frmDatabaseMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatabaseMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 15.03.2023 13:03
' ----------------------------------------------------------------
Option Explicit

Private findTimer As Single, _
        keyTimer As Single

Const TIME_SHIFT As Single = 0.5
Const LABEL_COOLDOWN As Single = 1
Const NOTIFY_MESSAGE = " Отправлено в книгу."

Const DEFAULT_SEARCH As String = "—"
Const RESULT_MAXSHOW_COUNT = 4

Const SHIFT_KEY = 1
Const CTRL_KEY = 2

Const TOP_INFO = 158
Const TOP_VERSION = 182

Private Enum ETargetResult
    
    noResult = 0
    firstResult = 1
    secondResult = 2
    thirdResult = 3
    fourthResult = 4

End Enum

Private Enum ETargetMoving
    
    defaultItem
    nextItem
    previousItem
    
End Enum

Private Enum EApplyTypes
    
    notype
    newItem
    UpdateKey
    
End Enum

Private targetResult As ETargetResult, _
        applyType As EApplyTypes
Private Sub UserForm_Initialize()
    
    SetEventControls Me
    
    Me.labelBackground = DataBase.SearchBackgroundText
    Me.labelBackground.Width = Len(Me.labelBackground.caption) * 7
    
    DefaultControls
    
    FillResults
    targetResult = GetTargetResult(defaultItem)
    SelectResult

    Me.LabelInfo = DataBase.LabelInfoText
    Me.LabelInfo.Top = TOP_INFO
    Me.cheboxSearch = True
    
    Me.labelVersion = Addin.VersionCaption
    Me.labelVersion.Top = TOP_VERSION
    
'    Dim i As Integer
'    For i = 1 To 15
'        Me.listArchiveItem.AddItem i
'    Next

    applyType = notype
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 19.03.2023 10:25
    ' ----------------------------------------------------------------
    Private Sub FillResults( _
        Optional searchKey As String _
        )
        
        DefaultKeysLabel Me.labelResultKey1, DEFAULT_SEARCH, True
        DefaultKeysLabel Me.labelResultKey2, DEFAULT_SEARCH, True
        DefaultKeysLabel Me.labelResultKey3, DEFAULT_SEARCH, True
        DefaultKeysLabel Me.labelResultKey4, DEFAULT_SEARCH, True
        
        DataBase.FilterCache searchKey
        ' ----------------------------------------------------------------
        If Not CBool(DataBase.ResultFilterCount) Then _
            Exit Sub
        ' ----------------------------------------------------------------
        DefaultKeysLabel Me.labelResultKey1, DataBase.ResultFilter(1).key
        FillLabelResult DataBase.ResultFilter(1), labelResultMain1, labelResultSecondary1
        ' ----------------------------------------------------------------
        If DataBase.ResultFilterCount > 1 Then
            DefaultKeysLabel Me.labelResultKey2, DataBase.ResultFilter(2).key
            FillLabelResult DataBase.ResultFilter(2), labelResultMain2, labelResultSecondary2
        End If
        ' ----------------------------------------------------------------
        If DataBase.ResultFilterCount > 2 Then
            DefaultKeysLabel Me.labelResultKey3, DataBase.ResultFilter(3).key
            FillLabelResult DataBase.ResultFilter(3), labelResultMain3, labelResultSecondary3
        End If
        ' ----------------------------------------------------------------
        If DataBase.ResultFilterCount > 3 Then
            DefaultKeysLabel Me.labelResultKey4, DataBase.ResultFilter(4).key
            FillLabelResult DataBase.ResultFilter(4), labelResultMain4, labelResultSecondary4
        End If
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 22.03.2023 20:19
        ' ----------------------------------------------------------------
        Private Sub FillLabelResult( _
            frData As cFilterResults, _
            LabelLarge As MSForms.label, _
            LabelSmall As MSForms.label _
            )
            
            LabelLarge.foreColor = Colors.black
            LabelSmall.foreColor = Colors.black
            
            LabelLarge.Visible = True
            LabelSmall.Visible = True
            
            LabelLarge = frData.LabelLargeCaption
            LabelSmall = frData.LabelSmallCaption
            
            LabelLarge.ControlTipText = Replace(frData.TipText, Base.defaultValue, vbNullString)
        
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.03.2023 10:27
        ' ----------------------------------------------------------------
        Private Function ClearString( _
            data As String _
            ) As String
            
            ClearString = Left(data, InStr(data, Base.delimiterKeyValue) - 1)
            
        End Function
    ' ----------------------------------------------------------------
    ' Дата: 01.04.2023 16:18
    ' ----------------------------------------------------------------
    Private Function GetTargetResult( _
        Optional moving As ETargetMoving = defaultItem _
        ) As ETargetResult
        
        GetTargetResult = noResult
        
        Dim dbCount As Integer
        dbCount = DataBase.ResultFilterCount
        
        If dbCount > RESULT_MAXSHOW_COUNT Then _
            dbCount = RESULT_MAXSHOW_COUNT
        
        If Not CBool(dbCount) Then _
            Exit Function
            
        If moving = defaultItem Then _
            GetTargetResult = firstResult: _
            Exit Function
                
        Dim result As ETargetResult
        If moving = nextItem Then
            
            result = firstResult
            If targetResult < dbCount Then _
                result = targetResult + 1
            
        End If
        
        If moving = previousItem Then

            result = dbCount
            If targetResult > 1 Then _
                result = targetResult - 1

        End If
        
        GetTargetResult = result
    
    End Function
Private Sub UserForm_Activate()
    Me.texboxSearch.SetFocus
End Sub
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Me.texboxSearch.SetFocus
'End Sub
Private Sub cheboxSearch_Change()
    ColorChebox Me.cheboxSearch, blueDark
End Sub
    Private Sub ColorChebox( _
        oChebox As Object, _
        Optional color As Colors = Colors.blue _
        )
        
        oChebox.foreColor = Colors.black
        
        If oChebox Then _
            oChebox.foreColor = color
        
    End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 20:21
' ----------------------------------------------------------------
Private Sub btnNewItem_Click()
    
    ResetButtonsAddColors
    ShowAdditionalControls

    applyType = newItem
    texboxKey_Change
    
    Me.texboxKey.SetFocus
    Me.btnNewItem.BackColor = Colors.yellowPastel
    
End Sub
' ----------------------------------------------------------------
' Дата: 28.03.2023 21:30
' ----------------------------------------------------------------
    Private Sub ResetButtonsAddColors()

        Me.btnNewItem.BackColor = Colors.white
        Me.btnUpdateKey.BackColor = Colors.white

    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 28.03.2023 21:17
    ' ----------------------------------------------------------------
    Private Sub ResetNewAndUpdateControls()
        
        Me.btnApply.Enabled = False
        Me.btnApply.BackColor = Colors.white
        
        Me.texboxKey.borderColor = Colors.grey
        Me.texboxKey.foreColor = Colors.black
    
    End Sub
    ' ----------------------------------------------------------------
    ' Дата: 26.03.2023 20:23
    ' ----------------------------------------------------------------
    Private Sub ShowAdditionalControls()
        
        Me.LabelInfo.Visible = False
        Me.labelVersion.Visible = False
        
        Me.texboxKey.Visible = True
        
        If Not CBool(Len(Me.texboxKey)) Then _
            Me.labelKey.Visible = True
        
        Me.btnCancel.Visible = True
        Me.btnApply.Visible = True
        
    End Sub
' ----------------------------------------------------------------
' Дата: 29.03.2023 09:17
' ----------------------------------------------------------------
Private Sub btnApply_Click()
    
    Select Case True
        
        Case applyType = newItem
        
            DataBase.TargetItemKey = DataBase.AddItem(Me.texboxKey)
            DataBase.Edit DataBase.TargetItemKey
            
            VBA.Unload Me
        
        Case applyType = UpdateKey
            
            Dim status As Boolean
            If DataBase.IsKeyExists(Me.texboxKey) Then  'слияние элементов
               
                status = DataBase.Merge( _
                    resultKey:=Me.texboxKey, _
                    wipeKey:=DataBase.TargetItemKey _
                    )
            Else 'переименование ключа
                
                status = DataBase.UpdateKey( _
                    sourceKey:=DataBase.TargetItemKey, _
                    newKey:=Me.texboxKey _
                    )

            End If
            
            If status Then _
                UserForm_Initialize
    End Select
    
End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 20:26
' ----------------------------------------------------------------
Private Sub btnCancel_Click()
        
    applyType = notype
    Me.texboxKey = vbNullString
    
    HideAdditionalControls
    ResetButtonsAddColors
    
    Me.texboxSearch.SetFocus

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 26.03.2023 21:18
    ' ----------------------------------------------------------------
    Private Sub HideAdditionalControls()
    
        Me.texboxKey.Visible = False
        Me.labelKey.Visible = False
        
        Me.btnCancel.Visible = False
        Me.btnApply.Visible = False
        
        Me.LabelInfo.Visible = True
        Me.labelVersion.Visible = True
    
    End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 20:25
' ----------------------------------------------------------------
Private Sub btnUpdateKey_Click()
    
    ResetButtonsAddColors
    ShowAdditionalControls
    
    applyType = UpdateKey
    
    Me.texboxKey = DataBase.TargetItemKey
    Me.labelKey.Visible = Me.texboxKey = vbNullString
    
    Me.texboxKey.SetFocus
    Me.btnUpdateKey.BackColor = Colors.yellowPastel

End Sub
' ----------------------------------------------------------------
' Дата: 19.03.2023 11:00, 24.03.2023 10:28
' ----------------------------------------------------------------
Private Sub labelInfo_Click()

    DataBase.OpenDatabaseDir
    Me.texboxSearch.SetFocus
    
End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 19:50
' ----------------------------------------------------------------
Private Sub btnRecache_Click()

    DataBase.ReCacheData
    UserForm_Initialize
    
    Me.texboxSearch = vbNullString
    Me.texboxSearch.SetFocus
    
End Sub
' ----------------------------------------------------------------
' Дата: 05.04.2023 22:13
' ----------------------------------------------------------------
Private Sub btnOpenFolder_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    Select Case True
    
        Case Shift = SHIFT_KEY, Shift = CTRL_KEY
            DataBase.Edit DataBase.TargetItemKey
            
        Case Else
            DataBase.OpenItemDir
            
    End Select
    
    UserForm_Initialize
    
End Sub
' ----------------------------------------------------------------
' Дата: 28.03.2023 15:12
' ----------------------------------------------------------------
Private Sub btnMainAction_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )

    DoAction
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 30.03.2023 15:38
    ' ----------------------------------------------------------------
    Private Sub DoAction()
    
        findTimer = Timer
            
        If Me.cheboxSearch Then _
            DataBase.MainAction
        
        If Not Me.cheboxSearch Then _
            DataBase.SpecialAction
        
        UserForm_Initialize
        NotifyUser
        
    End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 21:23
' ----------------------------------------------------------------
Private Sub texboxKey_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    keyTimer = Timer
    
    If KeyCode = vbKeyReturn Then
    
        If Me.btnApply.Enabled Then _
            btnApply_Click
        
    End If
    
End Sub
Private Sub texboxKey_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.texboxKey.SetFocus
End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 20:32
' ----------------------------------------------------------------
Private Sub texboxKey_Change()
    
    ResetNewAndUpdateControls
    Do While Timer - keyTimer < TIME_SHIFT
    
        Me.labelKey.Visible = Me.texboxKey = vbNullString
        DoEvents
        
    Loop
    
    ' ----------------------------------------------------------------
    If Me.texboxKey = vbNullString Then _
        Exit Sub
    
    Dim isKeyUnique As Boolean
    isKeyUnique = Not DataBase.IsKeyExists(Me.texboxKey)
    
    If isKeyUnique Then
    
        Me.btnApply.Enabled = True
        Me.btnApply.BackColor = Colors.yellowPastel
        Exit Sub
        
    End If
    ' ----------------------------------------------------------------
    Dim targetColor As Colors
    targetColor = Colors.redLips
    
    If applyType = UpdateKey Then
        
        If LCase(DataBase.TargetItemKey) <> LCase(Me.texboxKey) Then 'потому что Compare Text
        
            targetColor = Colors.yellowGold
            
            Me.btnApply.Enabled = True
            Me.btnApply.BackColor = Colors.yellowPastel
            
        End If
         
    End If
    
    Me.texboxKey.borderColor = targetColor
    Me.texboxKey.foreColor = targetColor
    
End Sub
' ----------------------------------------------------------------
' Дата: 30.03.2023 13:49
' ----------------------------------------------------------------
Private Sub texboxSearch_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    findTimer = Timer
    
    Select Case KeyCode
        
        Case vbKeyUp, vbKeyLeft
        
            KeyCode = vbEmpty
            targetResult = GetTargetResult(previousItem)
            SelectResult
            
    
        Case vbKeyDown, vbKeyRight
        
            KeyCode = vbEmpty
            targetResult = GetTargetResult(nextItem)
            SelectResult
            
        ' ----------------------------------------------------------------
        Case vbKeyReturn, vbKeyAdd
            
            If KeyCode = vbKeyAdd Then _
                Me.cheboxSearch = False
            
            KeyCode = vbEmpty
            DoAction
            
    End Select
    
End Sub
    Private Sub NotifyUser()
    
        Me.LabelInfo.caption = NOTIFY_MESSAGE
        Me.LabelInfo.foreColor = Colors.green
        
        Do While Timer - findTimer < LABEL_COOLDOWN
            DoEvents
        Loop
        
        Me.LabelInfo.foreColor = Colors.blue
        Me.LabelInfo.caption = DataBase.LabelInfoText
    
    End Sub
Private Sub labelBackground_Click()
    Me.texboxSearch.SetFocus
End Sub
Private Sub texboxSearch_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.texboxSearch.SetFocus
End Sub
Private Sub texboxFilter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.texboxFilter.SetFocus
End Sub
' ----------------------------------------------------------------
' Дата: 19.03.2023 11:30
' ----------------------------------------------------------------
Private Sub texboxSearch_Change()
    
    DefaultControls
    ' ----------------------------------------------------------------
    Do While Timer - findTimer < TIME_SHIFT
    
        Me.labelBackground.Visible = Me.texboxSearch = vbNullString
        DoEvents
        
    Loop
    ' ----------------------------------------------------------------
    
    FillResults Me.texboxSearch.text
    targetResult = GetTargetResult(defaultItem): _
    SelectResult
    
    If Me.texboxSearch = vbNullString Then _
        Exit Sub
    
    If Not CBool(DataBase.ResultFilterCount) Then: _
        Me.texboxSearch.borderColor = Colors.redLips: _
        Me.texboxSearch.foreColor = Colors.redLips
    ' ----------------------------------------------------------------
    Me.texboxSearch.SelStart = vbEmpty
    Me.texboxSearch.SelLength = Len(Me.texboxSearch)
    
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 23.03.2023 21:14
    ' ----------------------------------------------------------------
    Private Sub SelectResult( _
        )
        
        If targetResult = noResult Then _
            Exit Sub
        
        If applyType = newItem Then _
            btnCancel_Click
            
        Dim keyLabel As MSForms.label, _
            resultBigLabel As MSForms.label, _
            resultLittleLabel As MSForms.label
            
        Select Case True
        
            Case targetResult = firstResult
                Set keyLabel = Me.labelResultKey1
                Set resultBigLabel = Me.labelResultMain1
                Set resultLittleLabel = Me.labelResultSecondary1
            
            Case targetResult = secondResult
                Set keyLabel = Me.labelResultKey2
                Set resultBigLabel = Me.labelResultMain2
                Set resultLittleLabel = Me.labelResultSecondary2
            
            Case targetResult = thirdResult
                Set keyLabel = Me.labelResultKey3
                Set resultBigLabel = Me.labelResultMain3
                Set resultLittleLabel = Me.labelResultSecondary3
            
            Case targetResult = fourthResult
                Set keyLabel = Me.labelResultKey4
                Set resultBigLabel = Me.labelResultMain4
                Set resultLittleLabel = Me.labelResultSecondary4
            
        End Select
            
        If keyLabel = DataBase.TargetItemKey Then _
            Exit Sub
            
        If Not keyLabel.Enabled Then _
            Exit Sub
        
        PaintDefaultColors
            
        PaintLabel keyLabel, Colors.greenLight, Colors.greenLight
        PaintLabel resultBigLabel, Colors.greenLight, Colors.greenLight
        PaintLabel resultLittleLabel, Colors.greenLight, Colors.greenLight
        
        DataBase.TargetItemKey = keyLabel
        
        If applyType = UpdateKey Then _
            btnUpdateKey_Click
        
        Me.btnOpenFolder.Enabled = True
        Me.btnOpenFolder.BackColor = Colors.oragnePastel
        
        Me.btnMainAction.Enabled = True
        Me.btnMainAction.BackColor = Colors.greenPastel
        
        Me.btnUpdateKey.Enabled = True
        
    End Sub
        Private Sub PaintDefaultColors()
        
            PaintLabel Me.labelResultKey1, Colors.grey, Colors.black
            PaintLabel Me.labelResultKey2, Colors.grey, Colors.black
            PaintLabel Me.labelResultKey3, Colors.grey, Colors.black
            PaintLabel Me.labelResultKey4, Colors.grey, Colors.black
            
            PaintLabel Me.labelResultMain1, Colors.grey, Colors.black
            PaintLabel Me.labelResultMain2, Colors.grey, Colors.black
            PaintLabel Me.labelResultMain3, Colors.grey, Colors.black
            PaintLabel Me.labelResultMain4, Colors.grey, Colors.black
            
            Me.labelResultSecondary1.foreColor = Colors.black
            Me.labelResultSecondary2.foreColor = Colors.black
            Me.labelResultSecondary3.foreColor = Colors.black
            Me.labelResultSecondary4.foreColor = Colors.black
            
            Me.texboxSearch.borderColor = Colors.grey
            Me.texboxSearch.foreColor = Colors.black
            
        End Sub
    ' ----------------------------------------------------------------
    ' Дата: 19.03.2023 11:28,  30.03.2023 13:49
    ' ----------------------------------------------------------------
    Private Sub DefaultControls()

        Me.btnOpenFolder.Enabled = False
        Me.btnOpenFolder.BackColor = Colors.white
        
        Me.btnMainAction.Enabled = False
        Me.btnMainAction.BackColor = Colors.white
        
        Me.btnUpdateKey.Enabled = False
        
        PaintDefaultColors

        DefaultKeysLabel Me.labelResultKey1, "1", True
        DefaultKeysLabel Me.labelResultKey2, "2", True
        DefaultKeysLabel Me.labelResultKey3, "3", True
        DefaultKeysLabel Me.labelResultKey4, "4", True
        
        DefaultLabel Me.labelResultMain1
        DefaultLabel Me.labelResultMain2
        DefaultLabel Me.labelResultMain3
        DefaultLabel Me.labelResultMain4
        
        DefaultLabel Me.labelResultSecondary1
        DefaultLabel Me.labelResultSecondary2
        DefaultLabel Me.labelResultSecondary3
        DefaultLabel Me.labelResultSecondary4

        DataBase.TargetItemKey = vbNullString
        targetResult = noResult
        
        Me.texboxKey = vbNullString
        applyType = notype
        
        HideAdditionalControls
        ResetButtonsAddColors
        
    End Sub
        ' ----------------------------------------------------------------
        ' Дата: 29.03.2023 09:46
        ' ----------------------------------------------------------------
        Private Sub DefaultLabel( _
            label As MSForms.label, _
            Optional disable As Boolean = False _
            )
            
            label.Enabled = Not disable
            label.caption = vbNullString
            label.ControlTipText = vbNullString
            
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 23.03.2023 20:17
        ' ----------------------------------------------------------------
        Private Sub DefaultKeysLabel( _
            label As MSForms.label, _
            defaultText As String, _
            Optional italicFont As Boolean = False _
            )
            
            label.foreColor = Colors.grey
            label.Enabled = False
            
            If Not italicFont Then
            
                label.foreColor = Colors.black
                label.Enabled = True
                
            End If
                
            'label.Font.Italic = italicFont
            label.caption = defaultText
        
        End Sub
        ' ----------------------------------------------------------------
        ' Дата: 19.03.2023 13:19
        ' ----------------------------------------------------------------
        Private Sub PaintLabel( _
            label As MSForms.label, _
            borderColor As Colors, _
            foreColor As Colors _
            )
            
            label.borderColor = borderColor
            label.foreColor = foreColor
            
        End Sub
' ----------------------------------------------------------------
' Дата: 26.03.2023 19:30
' ----------------------------------------------------------------
Private Sub btnEdit_Click()
    DataBase.Edit key:=DataBase.TargetItemKey
End Sub
Private Sub togbtnChangeFilter_Change()
    'todo: ToggleButton Cache.subFilterType
End Sub
' ---------------------------------------------------------------
Private Sub labelResultKey1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = firstResult: SelectResult
End Sub
Private Sub labelResultMain1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = firstResult: SelectResult
End Sub
' ---------------------------------------------------------------
Private Sub labelResultKey2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = secondResult: SelectResult
End Sub
Private Sub labelResultMain2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = secondResult: SelectResult
End Sub
' ---------------------------------------------------------------
Private Sub labelResultKey3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = thirdResult: SelectResult
End Sub
Private Sub labelResultMain3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = thirdResult: SelectResult
End Sub
' ---------------------------------------------------------------
Private Sub labelResultKey4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = fourthResult: SelectResult
End Sub
Private Sub labelResultMain4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    targetResult = fourthResult: SelectResult
End Sub
