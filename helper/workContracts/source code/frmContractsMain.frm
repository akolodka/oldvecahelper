VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContractsMain 
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2550
   OleObjectBlob   =   "frmContractsMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContractsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cm As New cContractsManager

Private Sub btnContract_Click()
    cm.GetDocumentBy Contracts
End Sub

Private Sub btnPayment_Click()
    cm.GetDocumentBy PaymentInvoices
End Sub

Private Sub btnTravelContracts_Click()
    cm.GetDocumentBy TravelContracts
End Sub

Private Sub btnPayment_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
        
End Sub

Private Sub btnContract_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
    
End Sub
Private Sub btnTravelContracts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then _
        VBA.Unload Me
    
End Sub

