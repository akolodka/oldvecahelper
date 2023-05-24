VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_RenamePDF 
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3360
   OleObjectBlob   =   "UF_RenamePDF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_RenamePDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const RENAME_PASSPORT = True

Private Sub cmbPassport_Click()
    RenamePDF RENAME_PASSPORT
End Sub

Private Sub cmbPassport_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
End Sub

Private Sub cmbSvid_Click()

    RenamePDF
    Explorer.OpenActiveWbPath
    VBA.Unload Me
    
End Sub

Private Sub cmbSvid_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
End Sub


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
End Sub
