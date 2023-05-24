Attribute VB_Name = "DEBUG_MODULE"
Option Explicit

Sub RunTest()
    
    Dim workPath As String
    workPath = ActiveWorkbook.path
    
    Dim mask As String
    mask = "Õ»À_2104"
    
    Dim targetFileName As String
    targetFileName = Dir(workPath & Application.PathSeparator & "*" & mask & "*", vbHidden)

    '=======================
    Dim fso As New FileSystemObject
    Dim sourcePath As String: sourcePath = fso.BuildPath(workPath, targetFileName)
    Dim destinationPath As String: destinationPath = fso.BuildPath(Base.configDataPath, targetFileName)
    
    fso.CopyFile sourcePath, destinationPath
    
    
    Dim fileData As String
    fileData = Base.ContentFromFile(destinationPath)
    
    If InStr(fileData, ",") Then
    
        Dim newArr() As String
        newArr = Split(fileData, ",")
        Debug.Print newArr(LBound(newArr) + 2)
    Else
        
        newArr = Split(fileData, " ")
        Debug.Print Mid(newArr(LBound(newArr)), 2)
        
    End If

    '=======================
    fso.DeleteFile destinationPath
    
End Sub










'Sub F_010_VISALLBOOK()
'    Dim xlwbT As Workbook
'    For Each xlwbT In Application.Workbooks
'        xlwbT.Windows(1).Visible = 1
'    Next
'End Sub
