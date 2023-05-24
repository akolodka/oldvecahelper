VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatusBar 
   Caption         =   "Выполнение..."
   ClientHeight    =   450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "frmStatusBar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------------------------
' Дата: 05.04.2023 22:03
' ----------------------------------------------------------------
Option Explicit

Const DEFAULT_WIDTH As Single = 0.0001

Private Sub UserForm_Initialize()
    frameProgress.Width = DEFAULT_WIDTH
End Sub
