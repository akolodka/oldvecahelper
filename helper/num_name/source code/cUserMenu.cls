VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUserMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Наименование: cUserMenu (Class Module)
' Назначение: Управление пользовательскими меню
' Автор: akolodka
' Date: 23.10.2022 18:16
' ------------------------------------------------------

Option Explicit

Public Enum eMenuTypes
    
    unknownType
    
    mainNumName
    configNumName
    mainRenamePDF

End Enum

Private Type TUmeny
    
    isLoaded As Boolean
    type As eMenuTypes
    
End Type

Private this As TUmeny

Friend Property Get isLoaded() As Boolean
    isLoaded = this.isLoaded
End Property
Friend Property Let isLoaded(ByVal NewValue As Boolean)
    this.isLoaded = NewValue
End Property

Public Property Get typе() As eMenuTypes
    typе = this.type
End Property
Public Property Let typе(ByVal NewValue As eMenuTypes)
    this.type = NewValue
End Property


Friend Sub Unload( _
    )
    
    UnloadAll
    this.isLoaded = False

End Sub

