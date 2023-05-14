Attribute VB_Name = "UnloadUserForms"
Option Explicit
'Const MODULE_KEYWORD = "UnLoadMenus"

Const ADDIN_KEYWORD = "helper_"
Const PROC_KEYWORD = "UnloadAll"

Public Sub UnloadAll( _
    )
    
    If UserForms.count = vbEmpty Then _
        Exit Sub
    
    Dim uForm As Object
    For Each uForm In UserForms
    
        VBA.Unload uForm
    
    Next
    
End Sub

Public Sub InitiateUnload()
    
    Dim LoadedAddin As Object
    For Each LoadedAddin In Application.AddIns2
        
          If IsHelperAddin(LoadedAddin.name) Then _
            RunUnload LoadedAddin.name
    Next
    
End Sub
    Private Function IsHelperAddin( _
        addinName As String _
        ) As Boolean
        
        IsHelperAddin = False
        
        If InStr(addinName, ADDIN_KEYWORD) > 0 Then _
            IsHelperAddin = True
        
    End Function
    Private Sub RunUnload( _
        addinName As String _
        )
        
        Run "'" & addinName & "'!" & PROC_KEYWORD
        
    End Sub

'            Dim isExists As Boolean
'            isExists = isModuleExist(LoadedAddin.Name)
'
'            If isExists Then _
'
'
'        End If


'    Private Function isModuleExist( _
'        wbName As String _
'        ) As Boolean
'
'        Dim wb As Workbook
'        Set wb = Workbooks(wbName)
'
'        Dim i As Integer
'        For i = 1 To wb.VBProject.VBComponents.Count
'
'            If wb.VBProject.VBComponents.Item(i).Name = MODULE_KEYWORD Then
'
'                isModuleExist = True
'                Exit Function
'
'            End If
'
'        Next
'
'        Set wb = Nothing
'
'    End Function


