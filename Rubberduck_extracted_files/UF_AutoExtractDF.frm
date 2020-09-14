VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AutoExtractDF 
   Caption         =   "Auto Extract Data Fields"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   OleObjectBlob   =   "UF_AutoExtractDF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AutoExtractDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    SDSR.Value = Selection.Address

End Sub

Private Sub CancelButton_Click()

    UF_AutoExtractDF.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub

Private Sub GoButton_Click()

    Dim TheDFMan As New clsDataFieldsManager
    Call TheDFMan.m_AutoExtractDataFields(Range(SDSR.Value))

End Sub
