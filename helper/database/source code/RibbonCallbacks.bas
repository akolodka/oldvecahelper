Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Const WIDTH_SHORTED = 500

Private Sub ribbon_DatabaseConfig(control As IRibbonControl)
    LoadMenu configDataBase
End Sub

Private Sub ribbon_Instruments(control As IRibbonControl)
    LoadMenu instrumentsOLD, instruments
End Sub

Private Sub ribbon_Customers(control As IRibbonControl)
    LoadMenu organisationsOLD, organisations
End Sub

Private Sub ribbon_Archive(control As IRibbonControl)
    LoadMenu archiveOLD, archiveLocal
End Sub

Private Sub ribbon_Names(control As IRibbonControl)
    LoadMenu personsOLD, persons
End Sub

Private Sub ribbon_Etalons(control As IRibbonControl)
    LoadMenu etalonsOLD, etalons
End Sub
' ----------------------------------------------------------------
' Наименование: LoadMenu (Public Sub)
' Назначение: Передача управления с ribbon-эклемента
'    параметр menuType:
'    параметр configMenu:
' Дата: 08.10.2022 12:37
' ----------------------------------------------------------------
Public Sub LoadMenu( _
    typeMenu As eMenuTypes, _
    Optional typeDatabase As EDadaBaseTypes _
    )
    
    InitiateUnload
    ClearSingletone
    ' ----------------------------------------------------------------
    UMenu.typе = typeMenu 'todo: [-] отказаться от старого типа меню
    ' ----------------------------------------------------------------
    Dim isConfigCorrect As Boolean
    isConfigCorrect = Config.IsCorrect
    ' ----------------------------------------------------------------
    
    If Not isConfigCorrect Or typeMenu = configDataBase Then
        
        DataBase.Initialize persons
        frmDatabaseConfig.Show False
        
    Else
    
        DataBase.Initialize typeDatabase
        'todo: [+] выключить здесь instuments
        If typeDatabase = archiveLocal Or typeDatabase = instruments Then
            Z_UF_Search.Show False
        Else
            frmDatabaseMain.Show False
        End If

    End If
    
    UMenu.isLoaded = True
    
End Sub

