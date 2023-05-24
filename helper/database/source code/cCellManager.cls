VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCellManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const KEY_ORGANISATION = "аказчик"
Const KEY_TAX = "ИНН"
Const KEY_ADDRESS = "дрес:"
Const KEY_MANUFACTURER = "зготовител"

' ----------------------------------------------------------------
' Дата: 27.03.2023 20:28
' ----------------------------------------------------------------
Public Sub FillManufacturer( _
    key As String, _
    org As cItemOrganisation _
    )
    
    Dim isSended As Boolean
    isSended = SendToSheet( _
        key:=KEY_MANUFACTURER, _
        data:=org.shortName _
        )
    
    If Not isSended Then _
        ActiveCell = org.shortName: _
    
End Sub
'' ----------------------------------------------------------------
'' Дата: 30.03.2023 15:40
'' ----------------------------------------------------------------
'Public Sub SendToActiveCell( _
'    org As cItemOrganisation _
'    )
'
'    ActiveCell = org.shortName
'
'End Sub
Public Sub FillOrganisation( _
    key As String, _
    org As cItemOrganisation _
    )
    ' ----------------------------------------------------------------
    Properties.Company = key
    ' ----------------------------------------------------------------
    Dim isSended As Boolean
    isSended = SendToSheet( _
        key:=KEY_ORGANISATION, _
        data:=org.shortName _
        )
    
    If Not isSended Then _
        ActiveCell = org.shortName: _
        Exit Sub
    ' ----------------------------------------------------------------
    SendToSheet _
        key:=KEY_TAX, _
        data:=org.taxNumber
    ' ----------------------------------------------------------------
    SendToSheet _
        key:=KEY_ADDRESS, _
        data:=org.legalAddress
End Sub
    ' ----------------------------------------------------------------
    ' Дата: 25.03.2023 08:39
    ' ----------------------------------------------------------------
    Private Function SendToSheet( _
        key As String, _
        data As String _
        ) As Boolean
    
        SendToSheet = False
        
        Dim findedCell As Range
        Set findedCell = FindCell( _
            findKey:=key, _
            searchAtPrintArea:=True _
            )
    
        Dim isFinded As Boolean
        isFinded = Not findedCell Is Nothing
        
        If Not isFinded Then _
            Exit Function
            
        Dim targetCell As Range
        Set targetCell = findedCell.Offset(0, 1)
        
        targetCell = data
        
        SendToSheet = True
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 18.02.2023 17:36
    ' Назначение: поиск ячейки в области активного листа
    '    параметр findKey: поисковый запрос -- любой
    '    параметр lookType: xlPart или xlWhole
    '    параметр searchAtPrintArea: True или False
    ' Возвращаемый тип: Range
    ' ----------------------------------------------------------------
    Private Function FindCell( _
        findKey As String, _
        Optional lookType = xlPart, _
        Optional searchAtPrintArea = False _
        ) As Range
        
        Set FindCell = Nothing
    
        Dim searchRange As Range
        Set searchRange = ActiveWorkbook.ActiveSheet.UsedRange
        
        If searchAtPrintArea Then
            
            If ActiveWorkbook.ActiveSheet.PageSetup.PrintArea <> vbNullString Then _
                Set searchRange = Range(ActiveWorkbook.ActiveSheet.PageSetup.PrintArea)
            
        End If
        
        Dim result As Range
        Set result = searchRange.Find(findKey, LookAt:=lookType, LookIn:=xlValues)
        
        Set FindCell = result
        
    End Function

