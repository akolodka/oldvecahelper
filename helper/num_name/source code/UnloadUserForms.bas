Attribute VB_Name = "UnloadUserForms"
Option Explicit
'Const MODULE_KEYWORD = "UnloadUserForms"

Const ADDIN_KEYWORD = "helper_"
Const PROC_KEYWORD = "UnloadAll"

Public Sub UnloadAll( _
    )
    
    If UserForms.count = 0 Then _
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
