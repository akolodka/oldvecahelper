VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave( _
    ByVal SaveAsUI As Boolean, _
    Cancel As Boolean _
    )
    
    Addin.IncreaseBuild
    
End Sub

Private Sub Workbook_Open()

    If Not Config.isCheckAuto Then _
        Exit Sub
        
    If Cache.dateOfCheck = Date Then _
        Exit Sub 'автопроверка не чаще раза в день
        
    Update.Check

End Sub

Private Sub Workbook_BeforeClose( _
    Cancel As Boolean _
    )

    If Not Config.isInstallAuto Then _
        Exit Sub

    Update.Install silentMode

End Sub
