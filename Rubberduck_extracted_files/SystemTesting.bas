Attribute VB_Name = "SystemTesting"
Sub TestRelFilForAllInterv()
    
    'On Error GoTo ExitCode
    Dim TheDFSys As New clsDFSystem
    'TheDFSys.m_AutoExtractDataFields (Worksheets("ECMNodeTest").Range("F7:H7"))
    Call TheDFSys.m_EstablishEssentialData("DirectivesSheet")

    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("RFOutputTest", False)

'    Dim CGAcondeval As New clsCallFunc
'    Dim CGBcondeval As New clsCallFunc
'    Call CGAcondeval.m_SetOrigFuncStr("If(_OpType = " & Chr(34) & "Double pneumonectomy" & Chr(34) & ",1,0)")
'    Call CGAcondeval.m_AddNewArg("_OpType")
'    Call CGBcondeval.m_SetOrigFuncStr("If(_OpType = " & Chr(34) & "Lymph node biopsy" & Chr(34) & ",1,0)")
'    Call CGBcondeval.m_AddNewArg("_OpType")
'
'    Dim RelFilAllInterv As New clsRelationalFilter
'    RelFilAllInterv.m_SetParentFormat "PNum"
'    RelFilAllInterv.m_SetParentDFS TheDFSys
'    Call RelFilAllInterv.m_AddGroup("A", "PNum & OpDate & OpType & Surgeon", CGAcondeval)
'    Call RelFilAllInterv.m_AddGroup("B", "PNum & OpDate & OpType & Surgeon", CGBcondeval)
'
'    Dim IGCcondeval As New clsCallFunc
'    Call IGCcondeval.m_SetOrigFuncStr("If(A_OpDate > B_OpDate, 1, 0)")
'    Call IGCcondeval.m_AddNewArg("A_OpDate")
'    Call IGCcondeval.m_AddNewArg("B_OpDate")
'    Set RelFilAllInterv.m_InterGrpFilter = IGCcondeval

    Dim CGAcondeval As New clsCallFunc
    'Call CGAcondeval.m_SetOrigFuncStr("1") 'not necessary if the

    Dim RelFilAllInterv As New clsRelationalFilter
    RelFilAllInterv.m_SetParentFormat "Patient #"
    RelFilAllInterv.m_SetParentDFS TheDFSys
    Call RelFilAllInterv.m_AddGroup("A", "Patient # & Operation date", CGAcondeval)

    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If(A_NbInGrp > 1, 1, 0)")
    Call IGCcondeval.m_AddNewArg("A_NbInGrp")
    Set RelFilAllInterv.m_InterGrpFilter = IGCcondeval

    Call RelFilAllInterv.m_ApplyToEntries
    Call RelFilAllInterv.m_PrintFilteredEntries("RFOutputTestFiltered")

    Call TheDFSys.m_CleanupTempColumns
Exit Sub
ExitCode:
    Call TheDFSys.m_CleanupTempColumns

End Sub


Sub TestDPSystemSub()

    Dim TheDPSys As New clsDPSystem
    Call TheDPSys.m_DoInstructionsOnChartsSheet("TestDataProcess", "TestDPOut")
    
End Sub

Sub TestTC()

Dim Cobaye As New clsTransChart
Dim TCI As Range
Dim TCO As Range

Sheets("TestTC").Select

Set TCI = Range("F15:F17")
Set TCO = Range("G15:G17")

'Set TCI = Range("L21:L62")
'Set TCO = Range("M21:M62")

'Call Cobaye.m_BuildDictionary(Sheets("TestsRelFil").Cells(17, 3), Sheets("TestsRelFil").Cells(17, 4))

Call Cobaye.m_BuildDictionary(TCI, TCO)

Dim InputRange As Range, OutputRange As Range


Set InputRange = Range("L8")
Set OutputRange = InputRange.Offset(0, 1)

'Sheets("SinglePort2017").Select
'Set InputRange = Range("F2:F122")
'Set OutputRange = InputRange.Offset(0, 1)

'Call Cobaye.m_QuickTranslateCells(InputRange, OutputRange) 'much faster indeed
Call Cobaye.m_TranslateCells(InputRange, OutputRange, 0, 1)

'Sheets("Feuille1 - Tabela 1").Select
'Call Unit01.FixAllDatesInRange(Sheets("Feuille1 - Tabela 1").Range("D4:D341"), Sheets("Feuille1 - Tabela 1").Range("E4:E341"))
'Call Unit01.FixAllDatesInRange(Sheets("Feuille1 - Tabela 1").Range("H4:H341"), Sheets("Feuille1 - Tabela 1").Range("I4:I341"))

End Sub

Sub TestTCMultiChoices()

    Dim TheTC As New clsTransChart
    
    Call TheTC.m_BuildDictionary(Sheets("SPToRedCapV1").Range("N44:N48"), Sheets("SPToRedCapV1").Range("O44:O48"))
    
    Dim Arr(0 To 1) As Variant
    Arr(0) = "SPTestImport!I2:I21": Arr(1) = "RedCappedSPMod!O3:O22"
    
    Call TheTC.m_ConvertTextToMCRForColumn(Arr)
    


End Sub


'beware : code below may be old
Sub TestMatchMode1()
    'this is making me realize why I originally designed this with
    'the possibility of using different Entry Specifiers for each sheet
    'so that e.g. only the sheets-with-Patient #-as sole-ES get matched
    'to sheets with sheets-that-have-also-another-ES, such as Opdate.
    
    Dim TheDFSys As New clsDFSystem
    Call TheDFSys.m_EstablishEssentialData("ECMNodeTest")

    Call TheDFSys.m_BuildEntryUIDs

    Dim EntryObj As clsEntry

    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("ECMNodeTreeOutputTesting")

    Dim CGAcondeval As New clsCallFunc
    Call CGAcondeval.m_SetOrigFuncStr("If(_DataC =" & Chr(34) & "fi" & Chr(34) & ", 1, 0)", 1)
    Call CGAcondeval.m_AddNewArg("_DataC", True)
    
    Dim CGBcondeval As New clsCallFunc
    Call CGBcondeval.m_SetOrigFuncStr("If(_DataC =" & Chr(34) & "sd" & Chr(34) & ", 1, 0)", 1)
    Call CGBcondeval.m_AddNewArg("_DataC", True)
    
    Dim CGCcondeval As New clsCallFunc
    Call CGCcondeval.m_SetOrigFuncStr("If(_DataC =" & Chr(34) & "as" & Chr(34) & ", 1, 0)", 1)
    Call CGCcondeval.m_AddNewArg("_DataC", True)
    
    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If(And(A_DoA > C_DoA, B_NbInGrp > 0), 1, 0)")
    Call IGCcondeval.m_AddNewArg("A_DoA", , True)
    Call IGCcondeval.m_AddNewArg("C_DoA", , True)
    Call IGCcondeval.m_AddNewArg("B_NbInGrp")
    
    Dim RelFil As New clsRelationalFilter
    Call RelFil.m_SetParentFormat("Pnum")
    Call RelFil.m_SetParentDFS(TheDFSys)
    
    Call RelFil.m_AddGroup("A", "Pnum & DoA", CGAcondeval)
    Call RelFil.m_AddGroup("B", "Pnum & DoA", CGBcondeval)
    Call RelFil.m_AddGroup("C", "Pnum & DoA", CGCcondeval)
    Set RelFil.m_InterGrpFilter = IGCcondeval
    
    Call RelFil.m_ApplyToEntries
    Call RelFil.m_PrintFilteredEntries("ECMNodeTestFiltered")

End Sub


