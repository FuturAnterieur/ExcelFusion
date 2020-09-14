VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_DPSys 
   Caption         =   "Data Processing System"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6450
   OleObjectBlob   =   "UF_DPSys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_DPSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    TextBoxDir.Value = ActiveSheet.Name
    TextBoxOutput.Value = ""

End Sub

Private Sub CancelButton_Click()

    UF_DPSys.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub

Private Sub GoButton_Click()

    Dim TheDPSys As New clsDPSystem
    Call TheDPSys.m_DoInstructionsOnChartsSheet(TextBoxDir.Value, TextBoxOutput.Value)

End Sub

