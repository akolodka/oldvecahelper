VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUserFormVisualisator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SUFFIX_ONE = "элемент"
Const SUFFIX_TWO = "элемента"
Const SUFFIX_THREE = "элементов"

Const LABEL_LENGTH = 50

' ----------------------------------------------------------------
' Дата: 19.03.2023 10:48
' ----------------------------------------------------------------
Public Function ItemTextSuffix( _
    value As String _
    ) As String
    
    Dim rightChar As String
    rightChar = Right(value, 1)
    
    Dim result As String
    Select Case rightChar
    
        Case 1
        
            result = value & " " & SUFFIX_ONE
            
            If Len(value) >= 2 Then
            
                If Right(value, 2) = 11 Then _
                    result = value & " " & SUFFIX_THREE
                
            End If
            
        Case 2 To 4
            result = value & " " & SUFFIX_TWO
            
            If Len(value) >= 2 Then
            
                If Right(value, 2) >= 12 And Right(value, 2) <= 14 Then _
                    result = value & " " & SUFFIX_THREE
                
            End If
            
        Case Else
            result = value & " " & SUFFIX_THREE
            
            
    End Select
    
    ItemTextSuffix = result
    
End Function
' ----------------------------------------------------------------
' Дата: 19.03.2023 10:55
' ----------------------------------------------------------------
Public Function ShortedString( _
    strData As String, _
    Optional maxLength As Integer = LABEL_LENGTH _
    ) As String
    
    ShortedString = strData ' по умолчанию возвращать всю строку 'название пути полностью умещается
    
    Dim leftPart As String, _
        rightPart As String, _
        iStrStart As Integer
        
    If Len(strData) > maxLength Then 'пусть не умещается полностью
        
        iStrStart = InStr(strData, Application.PathSeparator)
        
        If Not (CBool(iStrStart)) Then _
            Exit Function
        
        If iStrStart + 2 <= Len(strData) Then _
            leftPart = Left(strData, InStr(iStrStart + 2, strData, Application.PathSeparator)) 'обязательная левая часть строки
        
        If maxLength - Len(leftPart) - 5 >= 0 Then _
            rightPart = Right(strData, maxLength - Len(leftPart) - 5) ' чтобы не было ошибки в случае короткой строки

        ShortedString = leftPart & " ... " & rightPart
    End If
    
End Function

' ----------------------------------------------------------------
' Дата: 21.03.2023 21:28
' ----------------------------------------------------------------
Public Function LabelBackgroundText( _
    typeDb As EDadaBaseTypes _
    ) As String
    
    Dim result As String
    Select Case True
    
        Case typeDb = organisations
            result = "Поиск контрагента..."
        
        Case typeDb = instruments
            result = "Поиск средства измерений..."
        
        Case typeDb = persons
            result = "Поиск в адресной книге..."

        Case typeDb = etalons
            result = "Поиск сведений эталона..."

        Case typeDb = archiveLocal
            result = "Поиск в архиве..."
            
    End Select
    
    LabelBackgroundText = result
End Function

