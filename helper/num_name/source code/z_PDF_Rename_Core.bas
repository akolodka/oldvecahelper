Attribute VB_Name = "z_PDF_Rename_Core"
Option Explicit
Private fso As New FileSystemObject

'todo: перебрать этот модуль
Public Sub RenamePDF( _
    Optional renamePassport = False _
    )
    
    Dim pdfName As String, _
        pdfCount As Integer
        
    pdfName = Dir(ActiveWorkbook.path & Application.PathSeparator & "*.pdf")
    
    Do While pdfName <> vbNullString
        pdfCount = pdfCount + 1
        pdfName = Dir
    Loop
    
    Select Case pdfCount
        Case 0
            Exit Sub
        Case 1
            pdfName = Dir(ActiveWorkbook.path & Application.PathSeparator & "*.pdf")
            pdfName = ActiveWorkbook.path & Application.PathSeparator & pdfName
        Case Else
            pdfName = PathFromUserChoose
    End Select
    
    Dim newPDFname As String
    newPDFname = PDFpsNewName
    
    If Not renamePassport Then _
        newPDFname = PDFsvNewName

    If pdfName <> vbNullString Then _
        Name pdfName As newPDFname
        
End Sub
    '##########################################################################
    'функция предоставляет меню выбора файла
    Function PathFromUserChoose( _
        )    'функция выбора файла
        
        With Application.fileDialog(msoFileDialogFilePicker)
        
            .Filters.Clear
            .InitialView = msoFileDialogViewDetails
            .Title = "Выбор источника данных"
            .Filters.Add "файлы pdf", "*.pdf; *.pdf", 1
            
            .AllowMultiSelect = False:
            .InitialFileName = ActiveWorkbook.path & Application.PathSeparator & Dir(ActiveWorkbook.path & Application.PathSeparator & "*.pdf")
            
            If .Show = 0 Then _
                Exit Function
                
            PathFromUserChoose = .SelectedItems(1) 'полный путь к файлу
            
        End With
        
    End Function
    ' ----------------------------------------------------------------
    ' Дата: 17.04.2023 11:49
    ' ----------------------------------------------------------------
    Private Function PDFsvNewName( _
        ) As String
        
        Dim docName As String
        docName = ActiveWorkbook.name
        
        Dim maskName As String
        maskName = fso.GetBaseName(docName)
        
        Dim filePrefix As String
        filePrefix = "sv_"
        
        If InStr(maskName, "rc_") > vbEmpty Then _
            filePrefix = "srt_"
                    
        Dim arrSplit() As String
        arrSplit = Split(maskName, "_")
        
        Dim numberMask As String
        numberMask = arrSplit(LBound(arrSplit) + 1) & "_" & arrSplit(LBound(arrSplit) + 2) & "_" & arrSplit(LBound(arrSplit) + 3)
        
        Dim newName As String
        newName = filePrefix & numberMask & " -- " & ActiveWorkbook.BuiltinDocumentProperties("Comments") & ".pdf"
        newName = ReplaceBadSymbols(newName)
'        Dim splitArr() As String
'        splitArr = Split(ActiveWorkbook.BuiltinDocumentProperties("Comments"), " -- ")
'
'        Dim currComment As String
'        currComment = ActiveWorkbook.ActiveSheet.Name & " -- " & splitArr(UBound(splitArr))
        
'        newName = newName & currComment & ".pdf"
        PDFsvNewName = ActiveWorkbook.path & Application.PathSeparator & newName
    
    End Function
        '#########################################################
        'функция устраняет в строке запрещённые в имени файла символы
        Private Function ReplaceBadSymbols(sTempStr As String) As String
            
            sTempStr = Replace(sTempStr, "\", "_"): sTempStr = Replace(sTempStr, "/", "_")
            sTempStr = Replace(sTempStr, ":", "_"): sTempStr = Replace(sTempStr, "*", "_")
            sTempStr = Replace(sTempStr, "?", "_"): sTempStr = Replace(sTempStr, "<", "_")
            sTempStr = Replace(sTempStr, ">", "_"): sTempStr = Replace(sTempStr, "|", "_")
            sTempStr = Replace(sTempStr, """", "_")
            
            ReplaceBadSymbols = sTempStr
        End Function
    Private Function PDFpsNewName( _
        ) As String
        
        Dim splitArr() As String
        splitArr = Split(ActiveWorkbook.BuiltinDocumentProperties("Comments"), " -- ")
        
        Dim currComment As String
        currComment = "пс " & NameSheet & " -- " & splitArr(UBound(splitArr))
        
        Dim newName As String
        newName = newName & currComment & ".pdf"
        newName = ReplaceBadSymbols(newName)
        
        PDFpsNewName = ActiveWorkbook.path & Application.PathSeparator & newName
    
    End Function
        Private Function NameSheet( _
            ) As String
            
            Dim currName As String
            currName = ActiveWorkbook.ActiveSheet.name
            
            If currName Like "*-*" And Len(currName) <= 6 And _
                InStr(currName, "-") <= 3 Then
            
                Dim arrNum() As String
                arrNum = Split(currName, "-")
                
                currName = "#" & Format(arrNum(UBound(arrNum)), "000") & " -- " & currName
            End If
            
            NameSheet = currName
        End Function

