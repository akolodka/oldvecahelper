Attribute VB_Name = "DEBUG_MODULE"
Option Explicit

Private fso As New FileSystemObject

Private Sub test()
    
    Dim md As Addin
    For Each md In Application.AddIns2
    
        If md.name <> ThisWorkbook.name Then _
            Debug.Print md.FullName
        
    Next
    
End Sub

Private Sub test2()
    
    Dim mdFolder As Folder
    Set mdFolder = fso.GetFolder(Base.excelAddinsDir)
    
    Dim md As file
    For Each md In mdFolder.Files
    
        Debug.Print md.path

    
    Next
    
End Sub

