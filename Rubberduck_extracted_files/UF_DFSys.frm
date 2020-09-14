VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_DFSys 
   Caption         =   "Data fusion system"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   OleObjectBlob   =   "UF_DFSys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_DFSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    TextBoxDir.Value = ActiveSheet.Name
    TextBoxOutput.Value = ""

End Sub

Private Sub CancelButton_Click()

    UF_DFSys.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub

Private Sub GoButton_Click()
'On Error GoTo ExitCode:

    Dim TheDFSys As New clsDFSystem
    Call TheDFSys.m_EstablishEssentialData(TextBoxDir.Value)

    Call TheDFSys.m_BuildEntryUIDs

    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet(TextBoxOutput.Value)
    
ExitCode:
    Call TheDFSys.m_CleanupTempColumns

End Sub
