VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBASpecForm 
   Caption         =   "VBASpec"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   OleObjectBlob   =   "VBASpecForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "VBASpecForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private oOwner As IVBASpecOutput

Friend Property Get Owner() As IVBASpecOutput
    Set Owner = oOwner
End Property
Friend Property Set Owner(oValue As IVBASpecOutput)
    Set oOwner = oValue
End Property

'========================================================================================================================================'

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii.Value = vbKeyEscape Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Or CloseMode = vbFormCode Then
        If oOwner.Running Then
            oOwner.Running = False
            Cancel = 1
        Else
            oOwner.Done = True
        End If
    Else
        oOwner.Running = False
    End If
End Sub
