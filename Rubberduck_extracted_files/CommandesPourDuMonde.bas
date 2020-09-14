Attribute VB_Name = "CommandesPourDuMonde"
Sub FilterSDSPOnAllInterv()

    Dim TheDFSys As New clsDFSystem
    'TheDFSys.m_AutoExtractDataFields (Sheets("ECMNodeTest").Range("F7:H7"))
    Call TheDFSys.m_EstablishEssentialData("AllOpsExporter")

    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("AllOpsExported", False)
    
    Call TheDFSys.m_CleanupTempColumns
    
    Dim CGAcondeval As New clsCallFunc
    'Call CGAcondeval.m_SetOrigFuncStr("1")

    Dim RelFilAllInterv As New clsRelationalFilter
    RelFilAllInterv.m_SetParentFormat "PNum & OpDate"
    RelFilAllInterv.m_SetParentDFS TheDFSys
    Call RelFilAllInterv.m_AddGroup("A", "PNum & OpDate & OpType", CGAcondeval)

    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If((A_NbInGrp > 1), 1, 0)")
    Call IGCcondeval.m_AddNewArg("A_NbInGrp")
    Set RelFilAllInterv.m_InterGrpFilter = IGCcondeval
    
    Call RelFilAllInterv.m_ApplyToEntries
    Call RelFilAllInterv.m_PrintFilteredEntries("SPSD_FilterOutput2")
    
End Sub



Sub CompareSPVersions()

'    Dim TheDFSys As New clsDFSystem
'    Call TheDFSys.m_EstablishEssentialData("CompareCharts")
'    Call TheDFSys.m_BuildEntryUIDs
'    Call TheDFSys.m_MatchEntriesAcrossSheets
'    Call TheDFSys.m_OutputValuesOnSheet("CompareOutput")


    Call CountOccurrences(Range("CompareCharts!N63:N83"), Range("CompareOutput!F3:F341"))
    Call CountOccurrences(Range("CompareCharts!R63:R70"), Range("CompareOutput!G3:G341"))
    Call CountOccurrences(Range("CompareCharts!N87:N103"), Range("CompareOutput!H3:H341"))
    
    Call CountOccurrences(Range("CompareCharts!N63:N83"), Range("LobectomyOnly!F3:F266"), Range("CompareCharts!P63:P83"))
    Call CountOccurrences(Range("CompareCharts!R63:R70"), Range("LobectomyOnly!G3:G266"), Range("CompareCharts!T63:T70"))
    Call CountOccurrences(Range("CompareCharts!N87:N103"), Range("LobectomyOnly!H3:H266"), Range("CompareCharts!P87"))

End Sub

Sub CompareEtienne25Jul()
    Dim TheDFSys As New clsDFSystem
    
    'Call TheDFSys.m_AutoExtractDataFields(Range("G6:H6"))
    
    Call TheDFSys.m_EstablishEssentialData("CompareEtienne25Jul")
    
    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    
    Call TheDFSys.m_OutputValuesOnSheet("CE25JulOutput")

End Sub

Sub FusionTransfusion1()

    Dim TheDFSys As New clsDFSystem
    'TheDFSys.m_AutoExtractDataFields (Worksheets("ECMNodeTest").Range("F7:H7"))
    Call TheDFSys.m_EstablishEssentialData("FusionTransfusion")
    
    
    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("TransfusionOutput")
    
End Sub

Sub TestRelFilPourSimon()

    Dim TheDFSys As New clsDFSystem
    Call TheDFSys.m_EstablishEssentialData("SimonProcessing")

    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("TousLesPatients")

    Dim CGAcondeval As New clsCallFunc
    Call CGAcondeval.m_SetOrigFuncStr("If(And(Or(_OpType = ""Segmentectomy"", _OpType = ""Wedge"", And(_OpType = ""Lobectomy"", _SurgSite = ""RUL"")), _OpDate > Date(2000,4,1), _OpDate < Date(2003,1,1)), 1,0)")
    Call CGAcondeval.m_AddNewArg("_OpType")
    Call CGAcondeval.m_AddNewArg("_SurgSite")
    Call CGAcondeval.m_AddNewArg("_OpDate")
    
    Dim CGBcondeval As New clsCallFunc
    Call CGBcondeval.m_SetOrigFuncStr("1")
    'Call CGBcondeval.m_AddNewArg("_OpType")
    
    'Dim CGCcondeval As New clsCallFunc
    
    Dim RelFilSimon As New clsRelationalFilter
    RelFilSimon.m_SetParentFormat "PNum"
    RelFilSimon.m_SetParentDFS TheDFSys
    Call RelFilSimon.m_AddGroup("A", "PNum & OpDate & OpType", CGAcondeval)
    Call RelFilSimon.m_AddGroup("B", "PNum & OpDate & OpType", CGBcondeval)
    
    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If(And(B_OpDate < A_OpDate, B_SurgSide = A_SurgSide), 1, 0)")
    
    Call IGCcondeval.m_AddNewArg("A_OpDate")
    Call IGCcondeval.m_AddNewArg("B_OpDate")
    Call IGCcondeval.m_AddNewArg("A_SurgSide")
    Call IGCcondeval.m_AddNewArg("B_SurgSide")
    
    Set RelFilSimon.m_InterGrpFilter = IGCcondeval
    
    Call RelFilSimon.m_ApplyToEntries
    Call RelFilSimon.m_PrintFilteredEntries("PatientsFiltresPourSimon")

    Call TheDFSys.m_CleanupTempColumns
Exit Sub

ExitCode:
    Call TheDFSys.m_CleanupTempColumns

End Sub

Sub RelFilSimon_FullGroup()

    Dim TheDFSys As New clsDFSystem
    Call TheDFSys.m_EstablishEssentialData("SimonFiltering")

    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("AllAgain")

    Dim CGAcondeval As New clsCallFunc
    Call CGAcondeval.m_SetOrigFuncStr("If(And(Or(_OpType = ""Segmentectomy"", _OpType = ""Wedge"", And(_OpType = ""Lobectomy (1 lobe)"", Or(_SurgSite = ""RUL"", _SurgSite = """"))), _OpDate >= Date(2016,1,1), _OpDate <= Date(2016,12,31)), 1,0)")
    Call CGAcondeval.m_AddNewArg("_OpType")
    Call CGAcondeval.m_AddNewArg("_SurgSite")
    Call CGAcondeval.m_AddNewArg("_OpDate")

    Dim RelFilSimon As New clsRelationalFilter
    RelFilSimon.m_SetParentFormat "PNum"
    RelFilSimon.m_SetParentDFS TheDFSys
    Call RelFilSimon.m_AddGroup("A", "PNum & OpDate & OpType", CGAcondeval)

    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If(A_NbInGrp > 0, 1, 0)")
    Call IGCcondeval.m_AddNewArg("A_NbInGrp")
 
    Set RelFilSimon.m_InterGrpFilter = IGCcondeval
    
    Call RelFilSimon.m_ApplyToEntries
    Call RelFilSimon.m_PrintFilteredEntries("WedgeSegOrLobRUL2016")

    Call TheDFSys.m_CleanupTempColumns
Exit Sub

ExitCode:
    Call TheDFSys.m_CleanupTempColumns

End Sub

Sub RelFilSimon_GroupToReject()
    
    Dim TheDFSys As New clsDFSystem
    Call TheDFSys.m_EstablishEssentialData("SimonFiltering")

    Call TheDFSys.m_BuildEntryUIDs
    Call TheDFSys.m_MatchEntriesAcrossSheets
    Call TheDFSys.m_OutputValuesOnSheet("AllAgain")

    Dim CGAcondeval As New clsCallFunc
    Call CGAcondeval.m_SetOrigFuncStr("If(And(Or(_OpType = ""Segmentectomy"", _OpType = ""Wedge"", And(_OpType = ""Lobectomy (1 lobe)"", Or(_SurgSite = ""RUL"", _SurgSite = """"))), _OpDate >= Date(2016,1,1), _OpDate <= Date(2016,12,31)), 1,0)")
    Call CGAcondeval.m_AddNewArg("_OpType")
    Call CGAcondeval.m_AddNewArg("_SurgSite")
    Call CGAcondeval.m_AddNewArg("_OpDate")
    
    Dim CGBcondeval As New clsCallFunc
    Call CGBcondeval.m_SetOrigFuncStr("1")
    
    Dim RelFilSimon As New clsRelationalFilter
    RelFilSimon.m_SetParentFormat "PNum"
    RelFilSimon.m_SetParentDFS TheDFSys
    Call RelFilSimon.m_AddGroup("A", "PNum & OpDate & OpType", CGAcondeval)
    Call RelFilSimon.m_AddGroup("B", "PNum & OpDate & OpType", CGBcondeval)
    
    Dim IGCcondeval As New clsCallFunc
    Call IGCcondeval.m_SetOrigFuncStr("If(And(B_OpDate < A_OpDate, B_SurgSide = A_SurgSide, Not(B_SurgSide = """") ), 1, 0)")
    
    Call IGCcondeval.m_AddNewArg("A_OpDate")
    Call IGCcondeval.m_AddNewArg("B_OpDate")
    Call IGCcondeval.m_AddNewArg("A_SurgSide")
    Call IGCcondeval.m_AddNewArg("B_SurgSide")
    
    Set RelFilSimon.m_InterGrpFilter = IGCcondeval
    
    Call RelFilSimon.m_ApplyToEntries
    Call RelFilSimon.m_PrintFilteredEntries("WithSameSideOpBefore")

    Call TheDFSys.m_CleanupTempColumns
Exit Sub

ExitCode:
    Call TheDFSys.m_CleanupTempColumns
    


End Sub

