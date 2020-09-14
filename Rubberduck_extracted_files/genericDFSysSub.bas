Attribute VB_Name = "genericDFSysSub"
Sub GenericDFSys()


    Dim TheDFSys As New clsDFSystem
    'TheDFSys.m_AutoExtractDataFields (Sheets("FusionSPDemo").Range("J7:K7"))
    Call TheDFSys.m_EstablishEssentialData("FusionPourSurgSite")

    Call TheDFSys.m_BuildEntryUIDs

    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("FSSRecentSheetsOnly")

    Call TheDFSys.m_CleanupTempColumns



'    Dim TheRF As New clsRelationalFilter
'
'    Dim CGAcondeval As New clsCallFunc
'    Dim CGBcondeval As New clsCallFunc
'    Call CGAcondeval.m_SetOrigFuncStr("If(_IsLC = " & Chr(34) & "Yes" & Chr(34) & ",1,0)")
'    Call CGAcondeval.m_AddNewArg("_IsLC", True)
'
''    Call CGAcondeval.m_SetOrigFuncStr("If(_IsLC = 1,1,0)")
''    Call CGAcondeval.m_AddNewArg("_IsLC")
'
'    TheRF.m_SetParentFormat "PNum"
'    TheRF.m_SetParentDFS TheDFSys
'    Call TheRF.m_AddGroup("A", "PNum & OpDate", CGAcondeval)
'
'    Call TheRF.m_ApplyToEntries
'    Call TheRF.m_PrintFilteredEntries("LCFilter1")

End Sub
