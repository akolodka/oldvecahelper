VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContractsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PAYMENT_MASK = "*счета*"
Const CONTRACTS_MASK = "*по хд-*"
Const TRAVEL_CONTRACTS_MASK = "*хд с командир*"

Public Enum DocumentType

    PaymentInvoices
    Contracts
    TravelContracts
    
End Enum

Private fso As New FileSystemObject
Private cfg As New cConfigManagerContracts
    
Public Sub GetDocumentBy( _
    docType As DocumentType _
    )

    Dim mask As String
    mask = GetMaskFromDocType(docType)
    
    Dim copyStatus As Boolean
    copyStatus = CopyWorkBook(mask)
    
    If Not copyStatus Then _
        Exit Sub
    
    Dim destionationPath As String
    destionationPath = fso.BuildPath(cfg.destinationFolder, GetFileName(mask))

    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(destionationPath)
    
    FilterSheetByDocumentType wb, docType
    
    wb.Activate
    ActiveWindow.Zoom = GetZoom(docType)
    ActiveWindow.WindowState = xlMaximized
    
    wb.Save
    
End Sub
    Private Function CopyWorkBook( _
        mask As String _
        ) As Boolean
        
        CopyWorkBook = False
        
        Dim fileName As String
        fileName = GetFileName(mask)
        
        Dim sourcePath As String
        sourcePath = fso.BuildPath(cfg.sourceFolder, fileName)
        
        Dim destionationPath As String
        destionationPath = fso.BuildPath(cfg.destinationFolder, fileName)
        
        If Not fso.FileExists(sourcePath) Then _
            MsgBox "Файл """ & sourcePath & """ не существует. Проверьте настройки": _
            Exit Function
            
        If Not fso.FolderExists(cfg.destinationFolder) Then _
            MsgBox "Каталог """ & cfg.destinationFolder & """ не существует. Проверьте настройки": _
            Exit Function
            
        fso.CopyFile sourcePath, destionationPath
        CopyWorkBook = True
    
    End Function
        Private Function GetFileName( _
            mask As String _
            ) As String
    
            Dim maskPath As String
            maskPath = fso.BuildPath(cfg.sourceFolder, mask)
            
            Dim fileName As String
            fileName = Dir(maskPath)
            
            If fileName Like "~$*" Then _
                fileName = Dir
                
            If fileName <> vbNullString Then _
                GetFileName = fileName
        
        End Function
    Private Function GetMaskFromDocType( _
        docType As DocumentType _
        ) As String
        
        Select Case True
        
            Case docType = PaymentInvoices
                GetMaskFromDocType = PAYMENT_MASK
            
            Case docType = Contracts
                GetMaskFromDocType = CONTRACTS_MASK
                
            Case docType = TravelContracts
                GetMaskFromDocType = TRAVEL_CONTRACTS_MASK
        
        End Select
    
    End Function
    Private Sub FilterSheetByDocumentType( _
        wb As Workbook, _
        docType As DocumentType _
        )
        
        Select Case True
        
            Case docType = PaymentInvoices
                wb.ActiveSheet.UsedRange.AutoFilter Field:=11, _
                    Criteria1:="=*" & cfg.filterCriteria & "*", _
                    Operator:=xlFilterValues
                    
                wb.ActiveSheet.UsedRange.AutoFilter Field:=12, Criteria1:="="
            
            Case docType = Contracts
                wb.ActiveSheet.Rows("1:2").EntireRow.Delete
    
                wb.ActiveSheet.UsedRange.AutoFilter Field:=12, _
                    Criteria1:="=*" & cfg.filterCriteria & "*", _
                    Operator:=xlFilterValues
                    
                wb.ActiveSheet.UsedRange.AutoFilter Field:=9, Criteria1:="<>"
                wb.ActiveSheet.UsedRange.AutoFilter Field:=11, Criteria1:="="
                
            Case docType = TravelContracts
                wb.ActiveSheet.Rows("1:2").EntireRow.Delete

                wb.ActiveSheet.UsedRange.AutoFilter Field:=10, _
                    Criteria1:="=*" & cfg.filterCriteria & "*", _
                    Operator:=xlFilterValues
            
                wb.ActiveSheet.UsedRange.AutoFilter Field:=7, _
                    Criteria1:=RGB(255, 0, 0), _
                    Operator:=xlFilterCellColor
        
        End Select
        
    End Sub
    Private Function GetZoom( _
        docType As DocumentType _
        ) As Integer
        
        If docType = PaymentInvoices Then _
            GetZoom = 75: _
            Exit Function
        
        GetZoom = 100
        
    End Function

