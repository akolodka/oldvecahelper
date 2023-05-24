VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateConfig 
   Caption         =   "Настройки обновления"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "frmUpdateConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const SOMNIUM_ADDRESS = "\\Somnium\irlab\Документы\Помощник ПКР\"
Const SHIFTKEY_CONTROL = 2

Private Enum e_pathMaxLength
    fileName = 48
    FolderName = 43
End Enum

Private Enum e_captionTypes
    sаve
    rеady
End Enum

Dim cmbSaveCaption As e_captionTypes
Private fso As New FileSystemObject


Private Sub UserForm_Initialize()
    
    CheckToggle
    UpdateLabelCaptions
    
    ChangeSaveReadyCaption rеady
    Me.cmbSaveReady.SetFocus
    
    Update.Check
    
    If Cache.IsUpdateAvailable Then
    
        Me.cmbInstall.visible = True
        Me.labelUpdate.visible = False
        
        Me.cmbInstall.SetFocus

    End If
    
    Me.laberVersion = Addin.VersionCaption
    
End Sub
    Private Sub UpdateLabelCaptions( _
        )
        
        If Not fso.FolderExists(Config.pathUpdateFolder) Then _
            Exit Sub
            
        Me.labelSourceDir = FormattedLabelText(Config.pathUpdateFolder, FolderName)
        Me.labelSourceDir.Forecolor = Colors.turquoise
        Me.cmbChooseSourceDir.BackColor = Colors.white
        Me.cmbOpenSourceDir.Enabled = True

    End Sub
        Private Function FormattedLabelText( _
            sourceText As String, _
            maxlength As e_pathMaxLength _
            ) As String
            
            Dim resultText As String
            resultText = Replace(sourceText, Base.defaultValue, Base.defaultValuePrinted)
            
            FormattedLabelText = resultText
            
            If resultText = Base.defaultValuePrinted Then _
                Exit Function
            
            If Len(resultText) <= maxlength Then _
                Exit Function
            
            Dim stringStart As Integer, _
                stringLeftPart As String, _
                stringRightPart As String
                
            stringStart = InStr(resultText, Application.PathSeparator)
            If stringStart = 0 Then _
                Exit Function
            
            Dim lenLeft As Integer
            lenLeft = stringStart + 2
            
            If lenLeft > Len(resultText) Then _
                Exit Function
            
            stringLeftPart = Left(resultText, InStr(lenLeft, resultText, Application.PathSeparator))
            FormattedLabelText = resultText & " ... "
            
            Dim lenRight As Integer
            lenRight = maxlength - Len(stringLeftPart) - 5
            
            If lenRight < 0 Then _
                Exit Function
            
            stringRightPart = Right(resultText, lenRight)
            FormattedLabelText = resultText & " ... " & stringRightPart
        End Function
    Private Sub CheckToggle( _
        )
        
        If Not fso.FolderExists(Config.pathUpdateFolder) Then _
            Exit Sub
            
        Me.tgbCheckAuto.Enabled = True
        Me.tgbCheckAuto = Config.isCheckAuto
        
        If Me.tgbCheckAuto Then
            Me.tgbInstall.Enabled = True
            Me.tgbInstall = Config.isInstallAuto
        End If
        
    End Sub
Private Sub cmbOpenSourceDir_Click( _
    )
    Explorer.OpenFolder Config.pathUpdateFolder, True
End Sub
    
Private Sub cmbOpenSourceDir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me

End Sub
' ----------------------------------------------------------------
Private Sub cmbChooseSourceDir_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    If Shift = SHIFTKEY_CONTROL Then 'выбрать каталог на Somniumn
        Config.pathUpdateFolder = FolderFromFileDialog(, , True)
    Else 'выбор на рабочем столе
        Config.pathUpdateFolder = FolderFromFileDialog()
    End If
    
    UpdateLabelCaptions
    CheckToggle
    
    If Config.pathUpdateFolder <> Base.defaultValue Then _
        ChangeSaveReadyCaption
    
End Sub
    Private Sub ChangeSaveReadyCaption( _
        Optional captionType As e_captionTypes = sаve _
        )
        
        Select Case captionType
            
            Case sаve
            
                With Me.cmbSaveReady 'кнопка готово
                    .Font.Size = Font.Size - 1
                    .caption = "Сохранить"
                    .BackColor = Colors.yellowPastel
                End With
                cmbSaveCaption = sаve
                
            Case rеady
                
                With Me.cmbSaveReady 'кнопка готово
                    .Font.Size = Font.Size + 1
                    .caption = "Готово"
                    .BackColor = Colors.white
                End With
                cmbSaveCaption = rеady
                
        End Select
        
    End Sub
Private Sub cmbChooseSourceDir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
    
End Sub
' ----------------------------------------------------------------
Private Sub cmbSaveReady_Click()
    
    Select Case True
        Case cmbSaveCaption = sаve
            Config.Save
            ChangeSaveReadyCaption rеady
            
        Case cmbSaveCaption = rеady
            VBA.Unload Me
    
    End Select
End Sub

Private Sub cmbSaveReady_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single _
    )
    
    If Shift = 1 Then _
        Addin.Download
End Sub


Private Sub cmbSaveReady_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
        
End Sub
    ' ----------------------------------------------------------------
    ' Наименование: FolderFromFileDialog (Private Function)
    ' Назначение: диалоговое окно выбора каталога
    '    параметр dialogTitle:
    '    параметр initialFileName:
    '    параметр isSomniumPath:
    ' Дата: 05.12.2021 16:05
    ' ----------------------------------------------------------------
    Private Function FolderFromFileDialog( _
        Optional dialogTitle As String = "Выбор каталога", _
        Optional initialFileName As String, _
        Optional isSomniumPath As Boolean _
        )
        FolderFromFileDialog = Base.defaultValue
        
        If initialFileName = Base.defaultValue Then _
            initialFileName = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
            
        With Application.FileDialog(msoFileDialogFolderPicker)
        
            .Filters.Clear
            .InitialView = msoFileDialogViewDetails
            .AllowMultiSelect = False
            
            .Title = dialogTitle
            .initialFileName = initialFileName
            
            If isSomniumPath Then _
                .initialFileName = SOMNIUM_ADDRESS
            
            If .Show = False Then _
                Exit Function
            
            FolderFromFileDialog = .SelectedItems(1) & Application.PathSeparator
            
        End With
        
    End Function
Private Sub tgbCheckAuto_Change()
    
    With Me.tgbCheckAuto
        Config.isCheckAuto = .value
        
        .BackColor = Colors.white
        If .value Then _
            .BackColor = Colors.orangePastel
            
        Me.tgbInstall.Enabled = Me.tgbCheckAuto
        If Not .value Then _
            Me.tgbInstall = False
        
    End With
    
    ChangeSaveReadyCaption
End Sub
Private Sub tgbCheckAuto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
    
End Sub
Private Sub tgbInstall_Change()

    With Me.tgbInstall
        Config.isInstallAuto = .value
        
        .BackColor = Colors.white
        If .value Then _
            .BackColor = Colors.orangePastel
    
    End With
    
    ChangeSaveReadyCaption
End Sub
Private Sub tgbInstall_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
    
End Sub
Private Sub cmbInstall_Click()

    Update.Install isReload:=True
  
    Me.cmbInstall.visible = False
    Me.labelUpdate.caption = " Модули программы актуальны."
    Me.cmbSaveReady.SetFocus
    
End Sub

Private Sub cmbInstall_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
        
End Sub
