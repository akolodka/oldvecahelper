VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFilterResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 24.03.2023 09:11
' Автор: akolodka
' Назначение:
' ----------------------------------------------------------------
Option Explicit

Private Type TFilterData
    
    key As String
    LabelLargeCaption As String
    LabelSmallCaption As String
    TipText As String

End Type

Private this As TFilterData

Private Sub Class_Initialize()
    
    this.key = Base.defaultValue
    this.LabelLargeCaption = Base.defaultValue
    this.LabelSmallCaption = Base.defaultValue
    this.TipText = Base.defaultValue
    
End Sub
' ----------------------------------------------------------------
Public Property Get key() As String
    key = this.key
End Property
Public Property Let key(ByVal NewValue As String)
    this.key = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get LabelLargeCaption() As String
    LabelLargeCaption = this.LabelLargeCaption
End Property
Public Property Let LabelLargeCaption(ByVal NewValue As String)
    this.LabelLargeCaption = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get LabelSmallCaption() As String
    LabelSmallCaption = this.LabelSmallCaption
End Property
Public Property Let LabelSmallCaption(ByVal NewValue As String)
    this.LabelSmallCaption = NewValue
End Property
' ----------------------------------------------------------------
Public Property Get TipText() As String
    TipText = this.TipText
End Property
Public Property Let TipText(ByVal NewValue As String)
    this.TipText = NewValue
End Property
