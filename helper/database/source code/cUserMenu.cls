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
' Автор: akolodka
' Date: 20.02.2023 22:28
' ------------------------------------------------------
Option Explicit

Public Enum eMenuTypes
    
    'UnknownType
    configDataBase
    mainDatabase
    
    
    
    
    organisationsOLD
    instrumentsOLD
    etalonsOLD
    
    archiveOLD
    personsOLD
    
End Enum


Private Type TUmeny
    
    isLoaded As Boolean
    type As eMenuTypes
    
End Type

Private this As TUmeny
' ------------------------------------------------------
Friend Property Get isLoaded() As Boolean
    isLoaded = this.isLoaded
End Property
Friend Property Let isLoaded(ByVal NewValue As Boolean)
    this.isLoaded = NewValue
End Property
' ------------------------------------------------------
Public Property Get typе() As eMenuTypes
    typе = this.type
End Property
Public Property Let typе(ByVal NewValue As eMenuTypes)
    this.type = NewValue
End Property
' ------------------------------------------------------
Friend Sub Unload( _
    )
    
    UnloadAll
    this.isLoaded = False

End Sub
