Attribute VB_Name = "ControlsEvents"
Option Explicit

Const KEY_CHEBOX = "chebox"
Const KEY_TOGGLE = "togbtn"

Public EventControls As New Collection
Private oEvent As cEvents
' ----------------------------------------------------------------
' Наименование: SetEventControls (Public Sub)
' Назначение: Собрать все контролы формы для отслеживания событий
'    параметр UserForm:
' Дата: 09.10.2022 18:33
' ----------------------------------------------------------------
Public Sub SetEventControls( _
    UserForm As Object _
    )
    
    Set EventControls = Nothing
    Set oEvent = Nothing
    
    Dim ctrl As control, _
        isControlEvented As Boolean
        
    For Each ctrl In UserForm.Controls
        
        Set oEvent = New cEvents
        isControlEvented = True
        
        Select Case True

            Case TypeOf ctrl Is MSForms.CommandButton
                Set oEvent.oButton = ctrl

            Case TypeOf ctrl Is MSForms.CheckBox
            
                If IsControlCorrect(ctrl, KEY_CHEBOX) Then _
                    Set oEvent.oCheckBox = ctrl
                    
                If IsControlCorrect(ctrl, KEY_TOGGLE) Then _
                    Set oEvent.oToggleButton = ctrl
            
            Case TypeOf ctrl Is MSForms.ComboBox
                Set oEvent.oComboBox = ctrl
                
            Case TypeOf ctrl Is MSForms.ListBox
                Set oEvent.oListBox = ctrl
                
            Case TypeOf ctrl Is MSForms.TextBox
                Set oEvent.oTextBox = ctrl
                
'            Case TypeOf ctrl Is MSForms.label
'                Set oEvent.oLabel = ctrl
            
            Case Else
                isControlEvented = False 'если ни один контрол не был добавлен
            
        End Select
        
        If isControlEvented Then _
            EventControls.Add oEvent
            
    Next

End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.03.2023 14:47
    ' ----------------------------------------------------------------
    Private Function IsControlCorrect( _
        ctrl As control, _
        include As String _
        ) As Boolean
        
        IsControlCorrect = False
        
        Dim i As Integer: i = InStr(ctrl.name, include)
        If CBool(i) Then _
            IsControlCorrect = True 'если в имени контрола присутствует ключевые символы

    End Function

