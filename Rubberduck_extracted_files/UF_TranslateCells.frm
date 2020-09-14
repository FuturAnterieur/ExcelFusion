VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TranslateCells 
   Caption         =   "Translate Or Convert Cells"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   OleObjectBlob   =   "UF_TranslateCells.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TranslateCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()

UF_TranslateCells.Hide

End Sub



Private Sub CheckBoxCellOrText_Click()

End Sub

Private Sub UserForm_Initialize()


CheckBoxCase.Value = False
CheckBoxSpecialChars.Value = False
CheckBoxQuickOrPatient.Value = False
CheckBoxTextOrCell.Value = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub

Private Sub GoButton_Click()

    Dim TheTC As New clsTransChart

    Call TheTC.m_BuildDictionary(Range(TCI.Value), Range(TCO.Value), CheckBoxCase, CheckBoxSpecialChars)
    
    If CheckBoxQuickOrPatient.Value = False Then
        Call TheTC.m_QuickTranslateCells(Range(InputField.Value), Range(OutputField.Value), CheckBoxTextOrCell.Value)
    Else
        Call TheTC.m_TranslateCells(Range(InputField.Value), Range(OutputField.Value), CheckBoxTextOrCell)
    End If

End Sub

