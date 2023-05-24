VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oButton As MSForms.CommandButton
Attribute oButton.VB_VarHelpID = -1
Public WithEvents oCheckBox As MSForms.CheckBox
Attribute oCheckBox.VB_VarHelpID = -1
Public WithEvents oComboBox As MSForms.ComboBox
Attribute oComboBox.VB_VarHelpID = -1
Public WithEvents oListBox As MSForms.ListBox
Attribute oListBox.VB_VarHelpID = -1
Public WithEvents oTextBox As MSForms.TextBox
Attribute oTextBox.VB_VarHelpID = -1

Private Sub oButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyPressControl KeyCode, Shift
End Sub
Private Sub oCheckBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyPressControl KeyCode, Shift
End Sub
Private Sub oComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyPressControl KeyCode, Shift
End Sub
Private Sub oListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyPressControl KeyCode, Shift
End Sub
Private Sub oTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyPressControl KeyCode, Shift
End Sub

Private Sub KeyPressControl( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
    )
    
    If KeyCode = vbKeyEscape Then _
        UMenu.Unload
    
    If Not KeyCode = vbKeyReturn Then _
        Exit Sub

    If Not UMenu.typå = mainNumName Then _
        Exit Sub
        
    If Shift = vbKeyRButton Or Shift = vbKeyLButton Then 'Ctrl + Enter or Shift + Enter
        
        Dim statusOk As Boolean
        statusOk = NumName.GetFromJournal
        
        If Not statusOk Then
                   
            UF_NumName_Main.Hide
            UF_NumName_Main.Show False
            Exit Sub
            
        End If
        
    End If
    
    Protocol.SendNumberToSheet (NumName.FullNumber)
    Protocol.SendFactoryNumberToComment
    Protocol.RenameFile NumName.ProtocolFileMask, Cache.saveAsCopy
    
    UMenu.Unload

End Sub



