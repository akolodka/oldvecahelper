Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private Sub Ribbon_NumName(control As IRibbonControl)
    LoadMenu mainNumName
End Sub

Private Sub Ribbon_NumName_Config(control As IRibbonControl)
    LoadMenu configNumName
End Sub

Private Sub ribbon_RenamePDF(control As IRibbonControl)
    LoadMenu mainRenamePDF
End Sub

' ----------------------------------------------------------------
' Наименование: LoadMenu (Public Sub)
' Назначение: Передача управления с ribbon-эклемента
'    параметр menuType:
'    параметр configMenu:
' Дата: 08.10.2022 12:37
' ----------------------------------------------------------------
Public Sub LoadMenu( _
    menuType As eMenuTypes _
    )
    
    InitiateUnload
    ClearSingletone
    
    Dim isConfigCorrect As Boolean
    isConfigCorrect = Config.IsCorrect
    
    Select Case True
            
        Case Not isConfigCorrect, menuType = configNumName
            UF_NumName_Config.Show False
        
        Case menuType = mainNumName
            UF_NumName_Main.Show False
        
        Case menuType = mainRenamePDF
            UF_RenamePDF.Show False
            
    End Select
    
    UMenu.typе = menuType
    UMenu.isLoaded = True
    
End Sub

