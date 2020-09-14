VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ExtTerms 
   Caption         =   "ExtractTerms"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   OleObjectBlob   =   "UF_ExtTerms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ExtTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()

UF_ExtTerms.Hide

End Sub



Private Sub UserForm_Initialize()


ChkBoxCase.Value = False
ChkBoxSpecialChars.Value = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub

Private Sub GoButton_Click()

Dim DummySeparatorArray() As Variant
DummySeparatorArray = Array("+")

Dim Separators() As String

Separators = Split(SepsBox.Value, "_or_", , vbTextCompare)

Call ExtractTermsMk3(InputField.Value, OutputCell.Value, Separators, ChkBoxKeepUnsepVersions.Value, ChkBoxCase.Value, ChkBoxSpecialChars.Value)

End Sub
