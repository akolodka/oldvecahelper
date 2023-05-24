Attribute VB_Name = "ControlsEvents"
Option Explicit

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
                Set oEvent.oCheckBox = ctrl
                
            Case TypeOf ctrl Is MSForms.ComboBox
                Set oEvent.oComboBox = ctrl
                
            Case TypeOf ctrl Is MSForms.ListBox
                Set oEvent.oListBox = ctrl
                
            Case TypeOf ctrl Is MSForms.TextBox
                Set oEvent.oTextBox = ctrl
            
            Case Else
                isControlEvented = False 'если ни один контрол не был добавлен
            
        End Select
        
        If isControlEvented Then _
            EventControls.Add oEvent
            
    Next

End Sub
