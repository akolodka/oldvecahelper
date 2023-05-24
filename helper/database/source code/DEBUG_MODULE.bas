Attribute VB_Name = "DEBUG_MODULE"
Option Explicit


Sub RunRun()

    Dim count As Long
    count = 500
    
    Dim pb As New cProgressBar
    pb.Initialize count
    
    Dim i As Integer
    For i = 1 To count
    
        pb.ShowProgress
        
    Next i
    
    pb.Terminate

End Sub

Sub test2()
    Debug.Print Application.OperatingSystem
End Sub

Sub RunTest()

'    Dim ch As New cCharTable
'    ch.RefactorKeycodeFile

    ClearSingletone
    frmDatabaseMain.Show False

'    DataBase.Initialize instruments
'
'    Dim mi As New cItemInstrument
'    mi.typeDevice = "Alpha"
'    mi.regFgisNum = "46163-99"
'    mi.name = "альфа-спектрометр с полупроводниковым детектором многоканальный"
'    mi.methodic = "МИ 1798-87 «Альфа-спектрометры с полупроводниковыми детекторами. Методика поверки»"
'
'    DataBase.Add mi.typeDevice, mi
    
'    Dim et As New cItemEtalon
'    et.factoryNumber = "ГЭТ 6–2016"
'    et.dateExpire = "31.12.2025"
'
'    DataBase.Add "ГЭТ 6", et
'    DataBase.ReCacheData










'    DataBase.Initialize organisations
'
'    Dim result As New Collection
'    Set result = DataBase.FilterCache("ybb мен пет")
'
''    Dim ct As New cCharTable, source As String, result As String
''    source = "ybb дух"
''    result = ct.Reverse(source)
''
'
'    Debug.Print result.Count
    

    
    
'    DataBase.Merge "paper", "paper2"
'
    
    
    
'    frmDatabaseConfig.Show False
'     UMenu.isLoaded = True

End Sub

Private Sub test233()

    Dim str As String
    str = "Я люблю порно"
    
    Dim key As String
    key = " "
    
    Debug.Print InStr(str, key) 'первое вхождение ключа слева, считая СЛЕВА
    Debug.Print InStrRev(str, key) 'первое вхождение ключа справа, считая СЛЕВА
    
    Debug.Print InStr(StrReverse(str), key) 'первое вхождение ключа справа, считая СПРАВА
    Debug.Print InStrRev(StrReverse(str), key) 'первое вхождение ключа справа, считая СЛЕВА
    
    
End Sub

Sub tests()
    ThisWorkbook.IsAddin = False
    
End Sub

'Sub F_010_VISALLBOOK()
'    Dim xlwbT As Workbook
'    For Each xlwbT In Application.Workbooks
'        xlwbT.Windows(1).Visible = 1
'    Next
'End Sub


' ----------------------------------------------------------------
'     Наименование: SortArray (Private Function)
'     Назначение: сортировка одномерного массива методом вставки
'        параметр arrToSort: исходный массив
'        параметр isIncreaseSorting:
'     Возвращаемый тип: String()
'     Дата: 23.10.2022 19:00
'     ----------------------------------------------------------------

