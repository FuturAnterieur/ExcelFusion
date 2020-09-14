VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelationalFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_Groups As Object 'dico of clsChildEntryGrp objects
Public m_PairsOfGrpMembers As Object
Public m_InterGrpFilter As clsCallFunc 'dico of clsCallFunc objects
Public m_NumOfMembersCriteria As Object
Public m_FilteredEntries As Object 'of clsEntry objects -- both parent and child entries
Public m_ParentFormat As String
Private m_ParentDFSys As clsDFSystem
Private m_CurGrpsWithMembers As Integer

Public Sub Class_Initialize()
    Set m_Groups = CreateObject("Scripting.Dictionary")
    Set m_InterGrpCriteria = CreateObject("Scripting.Dictionary")
    Set m_NumOfMembersCriteria = CreateObject("Scripting.Dictionary")
    Set m_PairsOfGrpMembers = CreateObject("Scripting.Dictionary")
    Set m_FilteredEntries = CreateObject("Scripting.Dictionary")
    Set m_InterGrpFilter = Nothing
End Sub

Public Sub m_SetParentDFS(pdfs As clsDFSystem)
    Set m_ParentDFSys = pdfs
End Sub

Public Sub m_SetParentFormat(pfrmt As String)
    m_ParentFormat = pfrmt
End Sub

Public Sub m_AddGroup(Name As String, frmt As String, condeval As clsCallFunc)
    If Not m_Groups.Exists(Name) Then
        m_Groups.Add Key:=Name, Item:=New clsChildEntryGrp
    End If
    
    m_Groups(Name).m_SetFormat (frmt)
    m_Groups(Name).m_SetName (Name)
    Set m_Groups(Name).m_ConditionEvaluator = condeval
    Set m_Groups(Name).m_ParentDFSys = m_ParentDFSys

End Sub

Public Sub m_BuildPairs()
    m_PairsOfGrpMembers.RemoveAll
    m_CurGrpsWithMembers = 0

    Dim NumPairs As Integer
    If m_Groups.count > 0 Then
        NumPairs = 1
    Else
        NumPairs = 0
    End If
    
    For Each CEGrp In m_Groups.Items
'       Debug.Print CEGrp.m_GetName()
'       Debug.Print CEGrp.m_Members.count
        NumPairs = NumPairs * CEGrp.m_Members.count
        m_CurGrpsWithMembers = m_CurGrpsWithMembers + 1
    Next CEGrp
        
    For i = 1 To NumPairs
        m_PairsOfGrpMembers.Add Key:=i, Item:=""
    Next i
    
    Dim RepetitionRatio As Integer: RepetitionRatio = NumPairs
    For Each CEGrp In m_Groups.Items
        If CEGrp.m_Members.count > 0 Then
            RepetitionRatio = RepetitionRatio / CEGrp.m_Members.count
            For IndexInDict = 1 To NumPairs
                ElemValue = (((IndexInDict - 1) \ RepetitionRatio) Mod CEGrp.m_Members.count) + 1
                m_PairsOfGrpMembers(IndexInDict) = m_PairsOfGrpMembers(IndexInDict) & "/" & CEGrp.m_GetName & "_" & ElemValue
            Next IndexInDict
        End If
    Next CEGrp
    
'    For Each Pair In m_PairsOfGrpMembers.Items
'        Debug.Print (Pair)
'    Next Pair

End Sub

Public Sub m_ApplyToEntries()
    
    With m_ParentDFSys
    Sheets(.m_OutputSheetName).Select
    For Each MainUIDKey In .m_EntriesChart.Keys
        For Each CEGrp In m_Groups.Items
            CEGrp.m_Members.RemoveAll
            Set CEGrp.m_IndexedTable = Nothing
            Set CEGrp.m_IndexedTable = New Collection
        Next CEGrp
        

        
        Dim EntryObj As clsEntry: Set EntryObj = .m_EntriesChart(MainUIDKey)
        Dim EntryMatchesCriteria As Boolean: EntryMatchesCriteria = False
        If EntryObj.m_Format = m_ParentFormat Then
            'Debug.Print ("New Parent : " & MainUIDKey)
            For Each ChildName In EntryObj.m_Children
                Dim ChildObj As clsEntry
                Set ChildObj = EntryObj.m_Children(ChildName)
                For Each CEGrp In m_Groups.Items
                    Dim Res As Boolean
                    
                    Res = CEGrp.m_PassTestForApplicant(ChildObj)
'                    If (Res) Then
'                        Debug.Print "success"
'                    End If
                Next CEGrp
            Next ChildName
            'une fois que les groupes sont construits, il s'agit de filtrer selon deux méthodes principales:
            '1- nombre de chaque groupe
            '2- critères inter-groupes (relationnels)
            'du genre "Au moins 1 du Groupe A, et au moins 1 du groupe B dont la date précède celle du membre groupe A"
            'relational criteria should be evaluated before "number in each group" :
            'le "number in each group" réfère au nombre de membres qui respectent tous les critères voulus.
            If Not m_InterGrpFilter Is Nothing Then
                Dim Result As Integer
                Dim Parts() As String
                Dim GrpName As String
                Dim IDFName As String
                Dim Index As Integer
                Dim NoIDFArgs As Boolean: NoIDFArgs = True
                
                For Each ArgName In m_InterGrpFilter.m_ArgNamesList
                    Parts = Split(ArgName, "_")
                    
                    GrpName = Parts(0): IDFName = Parts(1)
                    If Not IDFName = "NbInGrp" And Not IDFName = "NumInGrp" Then
                        NoIDFArgs = False
                    Else
                        Call m_InterGrpFilter.m_SetArgValue(CStr(ArgName), m_Groups(GrpName).m_Members.count)
                    End If
                Next ArgName
                If NoIDFArgs = True Then
                    Call m_InterGrpFilter.m_ReplaceIDFNamesWithVals
                    Result = CInt(m_InterGrpFilter.m_EvalExcelExpression)
                    If Result > 0 Then
                        'Debug.Print ("success for pair " & Pair)
                        EntryMatchesCriteria = True 'i.e. the parent entry is admitted into the filtered entries
                    End If
                Else
                    m_BuildPairs
                    For Each Pair In m_PairsOfGrpMembers.Items
                        Result = 0
                        Dim PairMemberNames() As String
                        Dim CurPairMembers As Object
                        Set CurPairMembers = CreateObject("Scripting.Dictionary")
                        CurPairMembers.RemoveAll
                        
                        PairMemberNames = Split(Right(Pair, Len(Pair) - 1), "/")
                        For Each Name In PairMemberNames
                            
                            Parts = Split(Name, "_")
                            
                            GrpName = Parts(0)
                            Index = Val(Parts(1))
                            If m_Groups.Exists(GrpName) Then
                                CurPairMembers.Add Key:=GrpName, Item:=m_Groups(GrpName).m_IndexedTable(Index)
                                'as such, we'll get GrpName - Entry UID pairs
                            End If
                        Next Name
                        
                        
                        Dim TargetRowOnOS As Integer, TargetColOnOS As Integer, TargetCell As Range, ArgVal As Variant
                        Dim AllRequestedArgNamesExistInThisPair As Boolean: AllRequestedArgNamesExistInThisPair = True
                        For Each ArgName In m_InterGrpFilter.m_ArgNamesList
                            Parts = Split(ArgName, "_")
                            
                            GrpName = Parts(0): IDFName = Parts(1)
                            Dim MemberUID As String
                            If .m_DFManager.m_DataFieldsChart.Exists(IDFName) Then
                                If CurPairMembers.Exists(GrpName) Then
                                    MemberUID = CurPairMembers(GrpName)
                                    'two types of Arg : "GrpName_IDFName, or GrpName_NumberOfMembersInGroup i.e. A_NumInGrp"
                                    
                                    TargetRowOnOS = .m_EntriesChart(MemberUID).m_MainValidMatchGrp.m_RowOnOutputSheet
                                    TargetColOnOS = .m_DFManager.m_DataFieldsChart(IDFName).m_ColOnOutputSheet
                                    Set TargetCell = Worksheets(.m_OutputSheetName).Cells(TargetRowOnOS, TargetColOnOS)
                                    'Debug.Print ("For " & GrpName & " (" & MemberUID & "), the target cell is " & TargetCell.Address)
                                    ArgVal = TargetCell.Address
                                    'Excel function can take range adresses as arguments (and this is what usually happens in normal Excel)
                                Else
                                    AllRequestedArgNamesExistInThisPair = False
                                End If
                            ElseIf IDFName = "NumInGrp" Or IDFName = "NbInGrp" Then
                                ArgVal = m_Groups(GrpName).m_Members.count
                                'note that NumInGrp args don't depend on pair-specific values, and going thru this
                                'everytime will be redundant...
                                'UNLESS the user uses something crazy like "A_LengthOfStay (or any numbered IDF param) > B_NumInGrp"
                            End If
                    
                            Call m_InterGrpFilter.m_SetArgValue(CStr(ArgName), ArgVal)
                            
                        Next ArgName
                        
                        Call m_InterGrpFilter.m_ReplaceIDFNamesWithVals
                        'Debug.Print m_InterGrpFilter.m_CurFuncStr
                        
                        If AllRequestedArgNamesExistInThisPair Then
                            Result = CInt(m_InterGrpFilter.m_EvalExcelExpression)
                        End If
                        
                        If Result > 0 Then
                            'Debug.Print ("success for pair " & Pair)
                            EntryMatchesCriteria = True 'i.e. the parent entry is admitted into the filtered entries
                            Exit For
                        End If
                    Next Pair
                End If
            Else
                 For Each CEGrp In m_Groups.Items
                    For Each MemberUIDKey In CEGrp.m_Members.Keys
                        EntryMatchesCriteria = True
                        Exit For
                        'CEGrp.m_IGCCompliantMembers.Add Key:=MemberUIDKey, Item:=0
                    Next MemberUIDKey
                    If EntryMatchesCriteria Then
                        Exit For
                    End If
                Next CEGrp
            End If
        End If
        If EntryMatchesCriteria Then
            m_FilteredEntries.Add Key:=MainUIDKey, Item:=0
        End If
    Next MainUIDKey
    End With
End Sub

Public Sub m_PrintFilteredEntries(StrShName As String)
    If Not SheetExists(StrShName) Then
        Sheets.Add(after:=ActiveSheet).Name = StrShName
        'if there's a way to get an equivalent statement with prettier syntax, I'm all ears
    Else
        Worksheets(StrShName).Cells.ClearContents
    End If
    
    Sheets(m_ParentDFSys.m_OutputSheetName).Select
    Dim LOC As Integer: LOC = m_ParentDFSys.m_LastOutputCol
    Dim DFNamesRng As Range
    Set DFNamesRng = Worksheets(m_ParentDFSys.m_OutputSheetName).Range(Cells(1, 1), Cells(2, LOC))
    DFNamesRng.Copy Worksheets(StrShName).Range("A1")
    'DFNamesRng.Copy Cells(1, 1)
    
    
    Dim i As Integer, j As Integer
    i = 2: j = 2
    
    With Worksheets(StrShName)
    For Each ParentUID In m_FilteredEntries.Keys
        i = i + 1
        Dim ParentEntryObj As clsEntry: Set ParentEntryObj = m_ParentDFSys.m_EntriesChart(ParentUID)
        
        .Cells(i, 1).Value = ParentUID
        
        If Not ParentEntryObj.m_MainValidMatchGrp Is Nothing Then
            .Cells(i, 2).Value = ParentEntryObj.m_GetMainLocationString
            Dim RoOS As Integer: RoOS = ParentEntryObj.m_MainValidMatchGrp.m_RowOnOutputSheet
            Dim DataRng As Range: Set DataRng = Worksheets(m_ParentDFSys.m_OutputSheetName).Range(Cells(RoOS, 3), Cells(RoOS, LOC))
            DataRng.Copy Worksheets(StrShName).Range(.Cells(i, 3), .Cells(i, LOC))
        End If
        For Each ChildObj In ParentEntryObj.m_Children.Items
            If Not ChildObj.m_MainValidMatchGrp Is Nothing Then
                i = i + 1
                RoOS = ChildObj.m_MainValidMatchGrp.m_RowOnOutputSheet
                Set DataRng = Worksheets(m_ParentDFSys.m_OutputSheetName).Range(Cells(RoOS, 1), Cells(RoOS, LOC))
                DataRng.Copy Worksheets(StrShName).Range(.Cells(i, 1), .Cells(i, LOC))
            End If

'            For Each LocIndex In ChildObj.m_InvalidInstances
            'Maybe TODO : add display of non-valid (duplicate) instances for entries matching the filter criteria
            ',even though duplicate instances may never be admitted to a child entry group during the filtering process --
            ' so this would be only useful in cases where some entries have filter-qualifying children as well as
            ' children with invalid instances, AND we want to see these invalid instances too.
'            Next LocIndex

        Next ChildObj
    Next ParentUID
    End With
End Sub
