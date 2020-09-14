Attribute VB_Name = "OldCodeVolume2"
Sub ShowOldCodeAndNeverBeCalled()


'Public Function ReplaceSheetNumsBySheetNames(InputStr As String) As String
'    Dim Res As String: Res = InputStr
'    For i = 1 To m_DFManager.m_SheetsChart.count
'        Res = Replace(Res, i & " : ", m_DFManager.m_SheetsChart.Items(i - 1).m_SheetName & " : ")
'    Next i
'    ReplaceSheetNumsBySheetNames = Res
'End Function

'Old version of the matching code
'Next version creates child/parents links even in cases of doublons,
'but only includes non-doublons in the SoloSheetsMatchGrp
'Public Sub m_MatchEntriesAcrossSheets()
'    'This function sets up all the matchgrps necessary for the OutputValues function to work properly.
'    'For each entry, it first matches every valid instance (i.e. it excludes instances from sheets
'    'where more than one instance of the same entry appears), and then it matches these valid instances
'    'with every valid instance of the parents of the entry in question.
'
'    Debug.Print ("In MatchEntries")
'    For Each UIDKey In m_EntriesChart.Keys
'        Debug.Print (UIDKey)
'        Dim MGCount As Integer
'        Dim EntryObj As clsEntry
'        Set EntryObj = m_EntriesChart(UIDKey)
'
'        Dim Location As clsEntryInstance
'        Dim SoloSheetsMatchGrp As clsMatchGroup
'
'        If EntryObj.m_ValidInstances.count > 0 Then
'            m_AllMatchGrps.Add New clsMatchGroup
'            MGCount = m_AllMatchGrps.count
'            Set SoloSheetsMatchGrp = m_AllMatchGrps(MGCount)
'            SoloSheetsMatchGrp.m_OwnerUID = UIDKey
'            Set EntryObj.m_MainValidMatchGrp = m_AllMatchGrps(MGCount)
'
'            For Each LocIndex In EntryObj.m_ValidInstances
'                 SoloSheetsMatchGrp.m_Participants.Add EntryObj.m_WhereCanIBeFound(LocIndex)
'            Next LocIndex
'
'            Dim ParentNodesToCheck As New Collection
'            Set ParentNodesToCheck = Nothing
'
'            Dim ECMNodeObj As clsECMNode
'            Dim PNObj As clsECMNode
'            Dim ChildrenMetInPassing As Object 'how poetic
'            Set ChildrenMetInPassing = CreateObject("Scripting.Dictionary")
'            ChildrenMetInPassing.RemoveAll
'
'            'Debug.Print UIDKey & " - " & EntryObj.m_Format
'            If m_EntriesConnectMap.Exists(EntryObj.m_Format) Then
'                Set ECMNodeObj = m_EntriesConnectMap(EntryObj.m_Format)
'
'                For Each ParentNodeName In ECMNodeObj.m_ParentNodesToCheck.Keys
'                    Set PNObj = m_EntriesConnectMap(ParentNodeName)
'                    Dim PNID As String: PNID = PNObj.m_IsolateInIDString(CStr(UIDKey))
'                    Dim ParentEntryObj As clsEntry
'
'                    If Not PNID = "" Then
'                        If Not m_EntriesChart.Exists(PNID) Then
'                            Debug.Print (PNID)
'                            m_EntriesChart.Add Key:=PNID, Item:=New clsEntry
'                            m_EntriesChart(PNID).m_Format = ParentNodeName
'
'                        End If
'                        Set ParentEntryObj = m_EntriesChart(PNID)
'                        EntryObj.m_Parents.Add Key:=PNID, Item:=ParentEntryObj
'                        ParentEntryObj.m_Children.Add Key:=UIDKey, Item:=EntryObj
'
'                        For Each PastFormat In ChildrenMetInPassing.Keys
'                            If PNObj.m_ChildNodesToCheck.Exists(PastFormat) Then
'                                Dim PastID As String: PastID = ChildrenMetInPassing(PastFormat)
'                                If Not ParentEntryObj.m_Children.Exists(PastID) Then
'                                    ParentEntryObj.m_Children.Add Key:=PastID, Item:=m_EntriesChart(PastID)
'                                End If
'                            End If
'                        Next PastFormat
'
'                        ChildrenMetInPassing.Add Key:=ParentNodeName, Item:=PNID
'
'
'                        'TODO : add linking for intermediate links produced
'                        '-- this might actually get recursive. -- DONE, and it didn't.
'                        'But it required yet another dictionary. I must have, like, 5 million of 'em by now
'
'                        For Each LocIndex In ParentEntryObj.m_ValidInstances
'                            SoloSheetsMatchGrp.m_Participants.Add ParentEntryObj.m_WhereCanIBeFound(LocIndex)
'                        Next LocIndex
'                    End If
'                Next ParentNodeName
'
'            End If
'
'
'        End If
'
'        For Each LocIndex In EntryObj.m_InvalidInstances
'            m_AllMatchGrps.Add New clsMatchGroup
'            MGCount = m_AllMatchGrps.count
'            m_AllMatchGrps(MGCount).m_Participants.Add EntryObj.m_WhereCanIBeFound(LocIndex)
'            m_AllMatchGrps(MGCount).m_OwnerUID = UIDKey
'        Next LocIndex
'    Next UIDKey
'End Sub
'
'
'
'End Sub
'
'
''    For Each ShObj In TheDFSys.m_DFManager.m_SheetsChart
''        Debug.Print (ShObj.m_SheetName & " : ")
''        For Each ESName In ShObj.m_LocalEntrySpecifiers
''            Debug.Print (ESName)
''        Next ESName
''    Next ShObj
''
'
''    Debug.Print ("Children of the tested entry: ")
''    Set EntryObj = TheDFSys.m_EntriesChart("OpDate = 1998-04-08")
''    For Each ChildEntryUID In EntryObj.m_Children.Keys
''        Debug.Print ChildEntryUID
''    Next ChildEntryUID
'
''    For Each NodeName In TheDFSys.m_EntriesConnectMap.Keys
''        Dim NodeObj As clsECMNode
''        Set NodeObj = TheDFSys.m_EntriesConnectMap(NodeName)
''        Debug.Print (NodeName & " :")
''        For Each PNode In NodeObj.m_ParentNodesToCheck.Keys
''            Debug.Print ("a parent is found in " & PNode)
''        Next PNode
''        For Each CNode In NodeObj.m_ChildNodesToCheck.Keys
''            Debug.Print "a child is found in " & CNode
''        Next CNode
''    Next NodeName
'
''    Dim EntryObj As clsEntry
''    For Each UID In TheDFSys.m_EntriesChart.Keys
''        Debug.Print UID
''        Set EntryObj = TheDFSys.m_EntriesChart(UID)
''        Dim LocCount As Integer: LocCount = 0
''        For Each Location In EntryObj.m_WhereCanIBeFound
''            LocCount = LocCount + 1
''            Debug.Print (LocCount & " : " & Location.m_ShNum & " " & Location.m_RowNum)
''        Next Location
''
''        Debug.Print ("Number of Sheets containing many instances : " & EntryObj.m_NbOfPolyRowSheets)
''
''        Debug.Print ("Instances that appear alone on their sheet : ")
''        For Each LocIndex In EntryObj.m_ValidInstances
''            Debug.Print (LocIndex)
''        Next LocIndex
''
''        Debug.Print ("Instances that appear in group on their sheet : ")
''        For Each LocIndex In EntryObj.m_InvalidInstances
''            Debug.Print (LocIndex)
''        Next LocIndex
''
''    Next UID
'
''Old data field instructions
''       ElseIf DFInstr Like "FullCellReplace(*)" Then
''                        TCName = Mid(DFInstr, FirstParenth + 1, SecondParenth - FirstParenth - 1)
''                        If m_TranslationCharts.Exists(TCName) Then
''                            DFObj.m_SingleValueInstructions.Add New clsDataFieldInstruction
''                            InstrCounter = DFObj.m_SingleValueInstructions.count
''                            DFObj.m_SingleValueInstructions(InstrCounter).m_FuncStr = "m_RetrieveDicoTerm"
''                            Set DFObj.m_SingleValueInstructions(InstrCounter).m_CallingObject = m_TranslationCharts(TCName)
''                            Debug.Print ("New TC (full cell mode) named " & TCName & " for data field " & IDFKey)
''                        End If
'
''                    ElseIf DFInstr = "FixDate" Then
''                        DFObj.m_SingleValueInstructions.Add New clsDataFieldInstruction
''                        InstrCounter = DFObj.m_SingleValueInstructions.count
''                        Set DFInstrObj = DFObj.m_SingleValueInstructions(InstrCounter)
''                        DFInstrObj.m_FuncStr = "FixOneDate"
''                        Set DFInstrObj.m_CallingObject = m_Utilities
'
''Malfunctioning ECM tree checking code that used RelativeTreeLevel
''        ElseIf NodesMetInPassing(CurNodeName) < NodesMetInPassing(PreviousNodeName) Then
''            MsgBox "Error : Node " & CurNodeName & " is being linked reciprocally to the same node."
''            Err.Raise 1003, "clsDFSystem - m_SetParentCollectionForNode", "Reciprocal parent-child link between two nodes"
''        ElseIf NodesMetInPassing(CurNodeName) <> RelativeTreeLevel Then
''            'it means the same node was already met at another (lower) level
''            Debug.Print NodesMetInPassing(CurNodeName) & " " & NodesMetInPassing(PreviousNodeName)
''            If Abs(WorksheetFunction.Min(RelativeTreeLevel, NodesMetInPassing(CurNodeName)) - NodesMetInPassing(PreviousNodeName)) <= 1 Then
''                'the next thing to check for (only warning though, not an error) are direct links between
''                'parents and children that are already linked through a separate chain.
''                MsgBox "Warning : Node " & CurNodeName & " is repetitively linked to the same child."
''            End If
''        End If
'
'
''the old, long version of DPSystem's establish essential data, before
''the creation of the m_DataFieldsManager class
'Public Sub m_EstablishEssentialData(ChartsSheetName As String)
'
'    Dim SrcSheetHeaderCell As Range, SrcSheetDFNames As Range, InternalDFNames As Range, OutputDFNames As Range, DFInstrRng As Range
'
'    Dim ChartsSheet As Excel.Worksheet
'    'First things first :
'    If Not SheetExists(ChartsSheetName) Then
'        MsgBox ("Specified directives sheet name " & ChartsSheetName & " does not exist in this workbook. Aborting.")
'        Err.Raise 1999
'    End If
'
'    Set ChartsSheet = Worksheets(ChartsSheetName)
'    ChartsSheet.Select
'    'first : determine Internal DF Name header : it will tell us where to find
'    'the rest of the cells that make up the all-important data fields chart.
'    Dim IDFHeaderCell As Range, IDFEnd As Range
'    Set IDFHeaderCell = ChartsSheet.Cells.Find(What:="Internal DF Name", _
'                                                        LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
'
'    If IDFHeaderCell Is Nothing Then
'        MsgBox ("Error : No Cell containing the text : InternalDFName could be found, so no data fields chart could be found. Aborting.")
'        Err.Raise 2000
'    End If
'
'    'Determine the edges of the Header field (to the left and right of the IDFHeaderCell)
'    Dim RightEdgeOfDFH As Range
'
'    Set SrcSheetHeaderCell = IDFHeaderCell.Offset(0, -1)
'    If SrcSheetHeaderCell.Text = "" Then
'        MsgBox ("Error : No input sheet specified. Aborting.")
'        Err.Raise 2001
'    End If
'
'    Set RightEdgeOfDFH = IDFHeaderCell.Offset(0, 1)
'    If Not IDFHeaderCell.Offset(0, 1).Text = "" Then
'        Set RightEdgeOfDFH = IDFHeaderCell.End(xlToRight)
'    End If
'
'    Set DataFieldsChartHeaders = Range(SrcSheetHeaderCell, RightEdgeOfDFH)
'
'    'Knowing these edges, determine the extent of the data fields chart by seeking
'    'the first empty row under the determined header row.
'    Dim DFCRowCount As Integer, FoundEmptyRow As Boolean
'    FoundEmptyRow = False
'    DFCRowCount = 0
'    Do While FoundEmptyRow = False
'        DFCRowCount = DFCRowCount + 1
'        Dim CurRow As Range
'        Set CurRow = DataFieldsChartHeaders.Offset(DFCRowCount, 0)
'        If Application.WorksheetFunction.CountA(CurRow) = 0 Then
'            FoundEmptyRow = True
'        End If
'    Loop
'
'    Set IDFEnd = IDFHeaderCell.Offset(DFCRowCount - 1, 0)
'
'    'Then determine every column's range of cells (going from "one below the current column header"
'    'to "at the level of the lowest cell in the chart")
'
'    Set InternalDFNames = Range(IDFHeaderCell.Offset(1, 0), IDFEnd)
'    Set SrcSheetDFNames = Range(SrcSheetHeaderCell.Offset(1, 0), IDFEnd.Offset(0, -1))
'
'    Set ODFHeaderCell = DataFieldsChartHeaders.Find(What:="Output DF Name")
'    If Not ODFHeaderCell Is Nothing Then
'        Set OutputDFNames = Range(ODFHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, ODFHeaderCell.Column))
'    End If
'
'    Set DFInstrHeaderCell = DataFieldsChartHeaders.Find(What:="Instructions")
'    If Not DFInstrHeaderCell Is Nothing Then
'        Set DFInstrRng = Range(DFInstrHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, DFInstrHeaderCell.Column))
'    End If
'
'    Set DFQualifHeaderCell = DataFieldsChartHeaders.Find(What:="Qualifiers")
'    If Not DFQualifHeaderCell Is Nothing Then
'        Set DFQualifRng = Range(DFQualifHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, DFQualifHeaderCell.Column))
'    End If
'
'    Dim DigitFinder As Object: Set DigitFinder = CreateObject("VBScript.RegExp")
'    DigitFinder.Global = False: DigitFinder.MultiLine = False: DigitFinder.IgnoreCase = True
'
'    Dim Descs() As String
'    Descs = Split(SrcSheetHeaderCell.Text, ";", 3) 'output no more than 3 strings.
'    'i.e. Split("A;B;C;D", ";", 3) -> Result: {"A", "B", "C;D"}
'    'The expected format is NameOfInputSheet; "Anything" "Number of header rows for this sheet"; Sheet qualifiers
'
'    Dim QualifStr As String: QualifStr = ""
'    Dim NumHR As Integer: NumHR = 1
'    If UBound(Descs) < 1 Then
'        Debug.Print ("No number of header rows specified for sheet " & Descs(0) & ". We will default it to 1.")
'    Else
'        DigitFinder.Pattern = "\d+"
'        Set Matches = DigitFinder.Execute(Descs(1))
'        NumHR = Matches(0)
'        'it just finds the first digit group and keeps that value, ignoring any text
'        If UBound(Descs) > 1 Then
'            QualifStr = Descs(2)
'            'sheet qualifiers will be processed later on.
'            'Up to now, the main sheet qualifier is "DFP = [value]", indicating sheet-specific data fusion priority.
'
'        End If
'    End If
'    'The Val function didn't work here, as it stops as once as it finds a non-digit character
'    If SheetExists(Descs(0)) Then
'        Call m_SourceSheet.m_DoInit(Descs(0), NumHR, QualifStr)
'    Else
'        MsgBox ("Error : Specified source sheet name " & Descs(0) & " does not belong to an existing sheet on this workbook.")
'        Err.Raise 2002
'    End If
'
'
'    Dim IDFName As String, ODFName As String
'    For Each IDFCell In InternalDFNames
'        ChartIndex = IDFCell.Row - InternalDFNames.Row + 1 '[Range].Row returns the same value as [Range].Cells(1,1).Row
'        If Not IDFCell.Text = "" Then
'            If Not m_DataFieldsChart.Exists(IDFCell.Text) Then
'                IDFName = IDFCell.Text
'
'                If Not ODFHeaderCell Is Nothing Then
'                    ODFName = OutputDFNames.Cells(ChartIndex, 1).Value
'                Else
'                    ODFName = IDFName
'                End If
'
'                m_DataFieldsChart.Add Key:=IDFName, Item:=New clsDataField
'                m_DataFieldsChart(IDFName).m_InternalOfficialName = IDFName
'                m_DataFieldsChart(IDFName).m_IndexOnDFChart = ChartIndex
'                m_DataFieldsChart(IDFName).m_NameOnOutputSheet = ODFName
'
'                If Not DFInstrHeaderCell Is Nothing Then
'                    m_DataFieldsChart(IDFName).m_Instructions = DFInstrRng.Cells(ChartIndex, 1).Text
'                End If
'                If Not DFQualifHeaderCell Is Nothing Then
'                    m_DataFieldsChart(IDFName).m_Qualifiers = DFQualifRng.Cells(ChartIndex, 1).Text
'                End If
'
'                If Not m_DFOutputNameToInternal.Exists(ODFName) Then
'                    m_DFOutputNameToInternal.Add Key:=ODFName, Item:=IDFName
'                End If
'            Else
'                MsgBox ("Warning : " & IDFCell.Text & " was already specified as an internal data field name on row " _
'                        & m_DataFieldsChart(IDFCell.Text).m_IndexOnDFChart & " and will be ignored on row " & ChartIndex)
'            End If
'        End If
'    Next IDFCell
'
'    Call m_CopyPasteIdenticalColumns(SrcSheetDFNames, InternalDFNames)
'
'End Sub

'
'Public Function FixOneDate(InputVal As Variant) As Variant
'
'    Dim Result As Variant
'    Dim DateResult As Date
'
'    Dim CaseNum As Integer
'    CaseNum = 0
'    If InputVal = "" Then
'        Result = ""
'    Else
'        If IsDate(InputVal) Then
'            CaseNum = 2
'            Result = InputVal
'        'TODO : use regexes to capture other less-common patterns
'        ElseIf InputVal Like "##/##/####" Or InputVal Like "##-##-####" Then
'            Dim YearVal As Integer, MonthVal As Integer, DayVal As Integer
'            YearVal = Right(InputVal, 4)
'            MonthVal = Mid(InputVal, 4, 2)
'            DayVal = Left(InputVal, 2)
'            Dim CheckValid As Boolean
'            CheckValid = YearVal > 1900 And MonthVal > 0 And MonthVal < 13 And DayVal > 0 And DayVal < 32
'            'CheckValidHarder should also check for NumDaysPerMonth and February 29th shenanigans
'            If CheckValid Then
'                CaseNum = 3
'                Result = DateSerial(YearVal, MonthVal, DayVal)
'            Else
'                CaseNum = -1
'                Result = "Cannot convert date (invalid day or month, or year inferior to 1900)"
'            End If
'        Else 'something, but that was not recognized as a date
'            CaseNum = 1
'            Result = InputVal
'        End If
'    End If
'
'    If CaseNum > 1 Then
'        DateResult = CDate(Result)
'        DateResult = Format(DateResult, "yyyy-MM-dd;@")
'        FixOneDate = DateResult
'    Else
'        FixOneDate = Result
'    End If
'
'End Function
'
'Public Sub FixAllDatesInRange(InputCells As Range, Optional OutputCells As Range)
'
'    If OutputCells Is Nothing Then
'        Set OutputCells = InputCells
'    End If
'
'    Dim IC As Range
'    Dim OC As Range
'
'    For Col = 1 To InputCells.Columns.count
'        For Row = 1 To InputCells.Rows.count
'            Set IC = InputCells.Cells(Row, Col)
'            Set OC = OutputCells.Cells(Row, Col)
'                OC.Value = FixOneDate(IC.Value)
'                OC.NumberFormat = "yyyy-MM-dd;@" 'this doesn't seem to be necessary, at least as far as my needs are concerned.
'                'well maybe it does change smth with the @. I should check this out further.
'        Next Row
'    Next Col
'
'    OutputCells.TextToColumns Destination:=OutputCells, DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)
'
''    ActualDateCell.Select
''    Selection.TextToColumns Destination:=ActualDateCell, DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)
''    ActualDateCell.NumberFormat = "yyyy-MM-dd;@"
''    this actually seems to be unnecessary
''    so the only scenario in which a whole-range method (and not one value at a time) would be useful would be if
''   the dates we are correcting are ambiguous, i.e. 06-05-1999, or even 05-03-02.
''   in a case like this, we would have to check other dates in the same range until we find one that isn't ambiguous
''   i.e. 12-13-1999 (or 05-20-00) and (hopefully) in the same format as the others around it.
''ALSO : Take heed that bulk operations (like calling TextToColumns on the whole column at once) would be faster
''than looping through each cell, but requires that dates are all uniformly formatted beforehand -- whereas
''the functions I am making are aimed at handling cases where many date formats appear in the same column
'
'
'End Sub
'Public Sub FixDuration(InputDurations As Variant)
'
'
'
'End Sub



'        If Mode = 0 Then
'            For Each TCITerm In m_Dictionary.Keys
'                ReplaceBy = m_GetFirstTCOTerm(m_Dictionary(TCITerm))
'                OutputCells.Replace What:=TCITerm, Replacement:=ReplaceBy, LookAt:=xlPart, MatchCase:=m_CaseSensitive
'            Next TCITerm
'        ElseIf Mode = 1 Then
'            For Each TCITerm In m_Dictionary.Keys
'
'                ReplaceBy = m_GetFirstTCOTerm(m_Dictionary(TCITerm))
'                OutputCells.Replace What:=TCITerm, Replacement:=ReplaceBy, LookAt:=xlWhole, MatchCase:=m_CaseSensitive
'
'            Next TCITerm
'        End If
