Attribute VB_Name = "RibbonCallbacks"
Option Explicit 'Потребовать явного объявления всех переменных в файле

'contractsConfig (элемент: button, атрибут: onAction), 2010+
Private Sub ContractsConfig(control As IRibbonControl)
    OpenConfig
End Sub

'buttonContracts (элемент: button, атрибут: onAction), 2010+
Private Sub ContractsMain(control As IRibbonControl)

    InitiateUnload
    frmContractsMain.Show False
End Sub

