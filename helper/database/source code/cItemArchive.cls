VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 19.03.2023 09:51
' Автор: akolodka
' ----------------------------------------------------------------
Option Explicit
Implements IDatabaseItem

Private Type TArchiveFields
    
    beginDate As String
    endDate As String
    typeSerial As String
    organisation As String
    comment As String
    
End Type

Private this As TArchiveFields
Private Sub Class_Initialize()
    
    this.beginDate = Base.defaultValue
    this.endDate = Base.defaultValue
    
    this.typeSerial = Base.defaultValue
    this.organisation = Base.defaultValue
    this.comment = Base.defaultValue
    
End Sub
Public Property Get beginDate() As String
    beginDate = this.beginDate
End Property
Public Property Let beginDate(ByVal NewValue As String)
    this.beginDate = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get endDate() As String
    endDate = this.endDate
End Property
Public Property Let endDate(ByVal NewValue As String)
    this.endDate = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get typeSerial() As String
    typeSerial = this.typeSerial
End Property
Public Property Let typeSerial(ByVal NewValue As String)
    this.typeSerial = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get organisation() As String
    organisation = this.organisation
End Property
Public Property Let organisation(ByVal NewValue As String)
    this.organisation = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get comment() As String
    comment = this.comment
End Property
Public Property Let comment(ByVal NewValue As String)
    this.comment = NewValue
End Property
