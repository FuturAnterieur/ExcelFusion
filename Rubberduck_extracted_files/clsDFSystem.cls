VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDFSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'TODO : add the possibility of having many "Commandes" or "Orders" on the same sheet,
         'so that they can all be executed at once
         'OH! And add several output sheet names, so that data output is automated too.
         'Ah! and a possibility of having sheet qualifiers too, like "Solo" or "Mute" (to select which go into the data fuser)

'MAIN TODOS For the future of now (august 18th) :
'
'   - Relational Filter operating from an already-existing normal or Fusion output sheet, without having to
'       run a full data fusion beforehand. Basically give it its own way to call (a version of) MatchEntryAcrossSheets.
'
'   - A way to call m_TranslateOneString passes in a configurable order, so I don't have to rewrite an algorithm
'   for every conceivable order; or at least, correct the algorithm for m_TranslateOneString_NormalFirst
'   (mainly with the addition of FTTClean, preventing untranslated parts of strings from being CapsUnlocked and
'   AccentStripped.
'
'
'   - "Bug"fix : Make it possible to replace single digits/numbers, "$", "#" and "_" in strings using SlowTextReplace.
'       For now, the TTR system makes this impossible.
'
'
               
Public m_DFManager As clsDataFieldsManager

Public m_EntriesChart As Scripting.Dictionary 'dictionary of UID - where it can be found
Public m_AllEntrySpecifiers As Scripting.Dictionary

Public m_ECMTree As clsECMTree
Public m_DoubleEntriesSeparated As Boolean
Public m_AllMatchGrps As Collection

Public m_DirectivesSheetName As String
Public m_OutputSheetName As String
Public m_DuplicatesOutputSheetName As String
Public m_LastOutputCol As Integer
Public m_LastOutputRow As Integer
Public m_UsesEntrySpecifiers As Boolean 'if this is False, then no datafusion is possible between sheets.

Public m_TranslationCharts As Scripting.Dictionary


Public Sub Class_Initialize()
    Set m_DFManager = New clsDataFieldsManager
    Set m_EntriesChart = New Scripting.Dictionary
    Set m_AllEntrySpecifiers = New Scripting.Dictionary
    Set m_ECMTree = New clsECMTree
    Set m_TranslationCharts = New Scripting.Dictionary
    Set m_AllMatchGrps = New Collection
    m_UsesEntrySpecifiers = False
    m_DuplicatesOutputSheetName = ""
End Sub

Public Sub m_AutoExtractDataFields(SheetNamesRng As Range)

    Call m_DFManager.m_AutoExtractDataFields(SheetNamesRng)
    
End Sub

' m_EstablishEssentialData is the main procedure that reads the directives that will make the clsDFSystem work.
' These directives are to be all written on an Excel sheet, which is why m_EstablishEssentialData only takes
' this sheet's name as an argument.

Public Sub m_EstablishEssentialData(ChartsSheetName As String)

    m_DFManager.m_InitFromChartsSheet (ChartsSheetName)
    m_DirectivesSheetName = ChartsSheetName
    

    Call FindAllTransChartsOnSheet(ChartsSheetName, Me, True)
    
    
    'The construction of the data structures (each IDF's m_WhereCanIBeFound
    ' and each sheet's m_LocalDFChart, mostly) that associate IDFNames with the
    'columns from the source sheets has to be done before IDF qualfiers and
    'instructions are processed.
       
    'Then for each IDF : qualifier and instruction assignation
    
    For Each IDF In m_DFManager.m_DataFieldsChart.Items
        Call IDF.m_ReadDataFusionQualifiers(Me)
        Call IDF.m_ReadDataFusionInstructions(Me)
    Next IDF
    
    
    'And finally : setting parent-child relationships between entries.
    'The idea is to have a family tree of sorts, for entries.
    'This step is entirely optional : it is only required if the user desires to use hierarchical (parent-to-child entry) data fusion
    'and/or relational filtering.
    
   Call m_ECMTree.m_ReadECMGraphOnChartsSheet(ChartsSheetName, m_AllEntrySpecifiers, True)
       
    
End Sub


'   m_BuildEntryUIDs is where the actual work done by clsDFSystem begins, once all orders
'   from the charts sheet have been read and understood.
'
'   The function accomplishes the first step of data fusion :
'   the building of the list of entries and their instances
'   on the source sheets.
'
'   Time for some terminological explanations :
'
'   As I had begun to explain in the notes to the "Entry Specifier" data field qualifier,
'   an Entry (represented in the clsEntry object) is uniquely identified by the name of
'   the entry specifiers it uses (its format), and the value associated to each of these specifiers.

'   The core idea of data fusion is that a single Entry may be found on several different source sheets.
'   The fact of finding an entry we already know of (i.e. "PNum = A555333 & OpDate = 1999-09-09") on a new
'   sheet is enough to assume the data columns from both of these sheets for the same entry belong
'   to the same set of data.
'
'   Since Entry Specifier status is bestowed to IDFs "as a whole", all source sheet data columns belonging
'   to an Entry Specifier IDF will be considered as containing Entry-Specifying data.
'   If an IDF has the ES qualifier (which is the only way for it to become
'   an Entry Specifier), then it will obligatorily be an entry specifier
'   on all sheets where it appears. Also, all entries found on a
'   given source sheet will obligatorily share the same format; after all,
'   if an IDF exists on a given source sheet, then it exists for all the rows
'   on the sheet.

'   Then comes the concept of an entry instance. It's not that hard.
'   While the Entry itself, as a concept, spans across source sheets and somehow exists in and of itself,
'   the rows where an Entry appears on the source sheets are called
'   this Entry's instances (encapsulated in the clsEntryInstance object).
'   An EntryInstance is mostly defined by the "source sheet identifier - row number" couple
'   indicating its location among the source sheets.

'   So for example, an entry such as "PNum = A555333 & OpDate = 1999-09-09" may have
'   two instances; one on row 367 of sheet "PatientDataA", and another on row 144 of
'   sheet "PatientDataB". The system will keep track of those instances and, later on,
'   will fuse the data requested for them in m_OutputValuesOnSheet.

'   But what happens if many instances of the same entry are found on the same source sheet?
'   Since pretty early in the development of this program, I have always considered this
'   case as problematic. Why? Because it can only mean two things :

'   1. The Entry Specifiers currently in use on this source sheet are not precise enough
'   and, as such, some rows that were supposed to represent different entries
'   end up with the same Entry ID. The solution would be to add at least one more
'   Entry Specifier in order to make every row on the sheet unique.

'   2. In the case where no other data field may be given Entry Specifier status
'   to distinguish between similar entries, then it means that a same "data owner" (i.e. the Entry) is given
'   more than one potentially contradictory data sets (the data columns) from the same source. In other words, we have
'   actual duplicates, and we have to choose which one is right; humans will probably be better at handling such cases.

'   The code currently in use will separate an entry's non-duplicate ("valid") instances
'   from its duplicate ("non-valid") instances; data will only be fused among valid instances
'   of a given entry, and non-valid instances will be left alone, since we cannot know which one
'   to compare to valid instances coming from other sheets.

'   One thing I COULD do, though, would be to add an option where the system, for a given entry,
'   attempts to 1. fuse invalid instances with other invalid instances in pairs from the same source sheet,
'   or 2. fuse all invalid and valid instances together and see what happens.
'   These modifications would mostly affect m_MatchEntriesAcrossSheets, though.


'   m_BuildEntryUIDs mostly builds the UIDs necessary to uniquely identify Entries;
'   it takes into account instructions given to the Entry Specifier data fields that
'   it needs to read. Once a UID is built for a given row of a given source sheet,
'   it checks if the Entry already exists and updates the relevant data structures
'   (like the Entry's m_WhereCanIBeFound, which keeps track of all that Entry's instances).

'   It builds up the m_EntriesChart for all entries that do have instances
'   on the source sheets. If parent-child relationships were specified through an
'   ECM Graph, m_MatchEntriesAcrossSheets may add "parent entries witout instances"
'   to the m_EntriesChart afterwards.
'
'   FINAL NOTE : if no Entry Specifiers were given (in the data fields chart),
'   then every row from every sheet will get a "sheet name & row number" entry ID,
'   and no matches nor any duplicate instances may occur.
'
'

Public Sub m_BuildEntryUIDs()
    
    Dim GlobalESCounter As Integer 'number of UID entry specifiers
    GlobalESCounter = 0
    Dim SheetESCounter As Integer
    Dim ESName As Variant
    Dim FoundIDFN As Boolean
    
    Dim Separator As String, EqualSign As String
    Separator = " & "
    EqualSign = " = "
    Dim LenSep As Integer
    LenSep = Len(Separator)
    
    For Each ShObj In m_DFManager.m_SheetsChart.Items
        Dim ShName As String
        ShName = ShObj.m_SheetName
        ShObj.m_ECMNode = ""
        If m_UsesEntrySpecifiers Then
            For Each ESDFName In ShObj.m_LocalEntrySpecifiers
                ShObj.m_ECMNode = ShObj.m_ECMNode & " & " & ESDFName
            Next ESDFName
            ShObj.m_ECMNode = Right(ShObj.m_ECMNode, Len(ShObj.m_ECMNode) - Len(" & "))
        End If

        Dim ShNum As Integer: ShNum = ShObj.m_InternalIndex
        Dim RowNum As Integer
        
        Dim ParentNodesToCheckForSheet As New Collection
        Set ParentNodesToCheckForSheet = Nothing
        
        
        For RowNum = ShObj.m_NumHeaderRows + 1 To ShObj.m_FurthestDataRow
            
            Dim CurRowID As String, IDVal As Variant
            CurRowID = ""
            
            Dim CurIDVal As Variant
            If m_UsesEntrySpecifiers Then
                For Each ESDFKey In ShObj.m_LocalEntrySpecifiers
                    Dim ESDF As clsDataField
                    Dim ESObj As clsEntrySpecifier
                    Set ESObj = m_AllEntrySpecifiers(ESDFKey)
                    Set ESDF = m_DFManager.m_DataFieldsChart(ESDFKey)
                    IDVal = "!empty!"
                    ESDFName = ESDFKey
                    Dim TargetCell As Range
                    
                    'The most extreme case that may be treated here is :
                    '2 (or more) cols on the same source sheet are meant to be fused as one on the output sheet
                    'AND this fused column is an entry specifier.
                    '(i guess the user would have to be kinda idiotic to try this seriously, but the program
                    'has to keep cool in such a case)
                     
                    'i.e.
                    
                    'SrcSheet         --   IDFN      -- Qualifiers
                    '_inCol5, _inCol6 --   EventDate -- ES
                    Dim ESCols() As String
                    ESCols = Split(ShObj.m_LocalDFChart(ESDFName), ";")
                    Dim CurDFPVal As Integer, HighestDFPVal As Integer
                    CurDFPVal = 0
                    HighestDFPVal = 0
                    CurIDVal = ""
                    
                    ContainsSeparator = False
                    For Each ESCol In ESCols
                        ESCol = Val(Replace(ESCol, "_", ""))
                        Set TargetCell = Sheets(ShName).Cells(RowNum, ESCol)
                        If ESDF.m_DFSpecificDFPMode = 2 Then
                            'if this is not the case, then the DFPriority is data sheet or
                            '(data sheet + internal data field) dependent
                            'and that means that every value we'll see here has the same priority.
                            Dim ESLocation As New clsDFLocationInfo
                            ESLocation.m_ShNum = ShObj.m_InternalIndex
                            ESLocation.m_ColNum = ESCol
                            CurDFPVal = ESDF.m_DFSpecificDFPVals(ESLocation.m_ConvertToString())
                        End If
                        
                        If CurDFPVal >= HighestDFPVal Then
                            If TargetCell.Text = "" Then
                                CurIDVal = "!empty!"
                            Else
                                CurIDVal = TargetCell.Value
                                For Each DFIFP In ESDF.m_SingleValueInstructions
                                    CurIDVal = CallByName(DFIFP.m_CallingObject, DFIFP.m_FuncStr, VbMethod, CurIDVal)
                                Next DFIFP
                                'a caveat here : we're going to reconvert this value in m_OutputValuesOnSheet
                            End If
                            
                            If CurDFPVal > HighestDFPVal Then
                                IDVal = CurIDVal
                                HighestDFPVal = CurDFPVal
                            Else
                                If IDVal = "!empty!" Then
                                    IDVal = CurIDVal
                                Else
                                    IDVal = IDVal & " vs " & CurIDVal
                                End If
                            End If
                                
                        End If
                    Next ESCol
                    
                    CurRowID = CurRowID & Separator & ESDFName & EqualSign & IDVal
                Next ESDFKey
                
                CurRowID = Right(CurRowID, Len(CurRowID) - LenSep)
                
            Else
                CurRowID = ShName & " & " & RowNum
            End If
            
            Dim LocCount As Integer
            If Not m_EntriesChart.Exists(CurRowID) Then
                m_EntriesChart.Add Key:=CurRowID, Item:=New clsEntry
                m_EntriesChart(CurRowID).m_UIDString = CurRowID
            End If
            Dim EntryObj As clsEntry
            Set EntryObj = m_EntriesChart(CurRowID)
            
            EntryObj.m_WhereCanIBeFound.Add New clsEntryInstance
            LocCount = EntryObj.m_WhereCanIBeFound.count
            EntryObj.m_WhereCanIBeFound(LocCount).m_ShNum = ShObj.m_InternalIndex
            EntryObj.m_WhereCanIBeFound(LocCount).m_ShName = ShObj.m_SheetName
            EntryObj.m_WhereCanIBeFound(LocCount).m_RowNum = RowNum
            EntryObj.m_WhereCanIBeFound(LocCount).m_OwnerUID = CurRowID
            
            If Not EntryObj.m_NbOfRowsOnEachSheet.Exists(ShNum) Then
                EntryObj.m_NbOfRowsOnEachSheet.Add Key:=ShNum, Item:=1
                EntryObj.m_ValidInstances.Add LocCount
            Else
                If EntryObj.m_NbOfRowsOnEachSheet(ShNum) = 1 Then
                    EntryObj.m_NbOfPolyRowSheets = EntryObj.m_NbOfPolyRowSheets + 1
                    Dim LatestIndexInVI As Integer
                    LatestIndexInVI = EntryObj.m_ValidInstances.count
                    EntryObj.m_InvalidInstances.Add (EntryObj.m_ValidInstances(LatestIndexInVI))
                    EntryObj.m_ValidInstances.Remove (LatestIndexInVI)
                End If
                EntryObj.m_NbOfRowsOnEachSheet(ShNum) = _
                EntryObj.m_NbOfRowsOnEachSheet(ShNum) + 1
                EntryObj.m_InvalidInstances.Add LocCount
            End If
            
            EntryObj.m_Format = ShObj.m_ECMNode
            EntryObj.m_HasInstancesOnSheets = True
        Next RowNum
    Next ShObj
    
End Sub

'   m_MatchEntriesAcrossSheets accomplishes two related tasks :
'   - creating parent-child links between entries, based on the m_ECMTree (if there is one)
'   - Matching valid instances of entries that do appear on the source sheets.

'   It should first be mentioned that, while the m_ECMTree keeps the structure of a hierarchical tree between
'   its nodes (which are entry formats), entries themselves do not. For a given child entry, any entry with
'   a format that appears above it in the m_ECMTree is considered as its parent, and this is reciprocal.

'   An interesting thing that m_MatchEntries does is that if it sees a given child entry with valid instances,
'   but for which no parent entry already exists, it will create an Entry object for this parent entry and add it to the m_EntriesChart.
'   This makes it possible to have parent entries that do not possess instances on the source sheets.
'   These are mostly useful for relational filtering, as they are completely inactive in the data fusion process.
'   It should be noted that any entry, even if it only has duplicate (non-valid) instances, is eligible
'   to become a parent and/or child to other entries. However, a non-valid entry (i.e. with instances,
'   but all of its instances are invalid) will be completely ignored, both in the data fusion process
'   and in relational filtering.

'   m_MatchEntriesAcrossSheets also sets up all the matchgrps necessary for the OutputValues function to work properly.
'   For each entry, it first matches every valid instance (it excludes instances from sheets
'   where more than one instance of the same entry appears i.e. duplicate entries), and then it matches these valid instances
'   with every valid instance of the parents of the entry in question.

'   As I said for m_BuildEntryUIDs, one thing I COULD do, though, would be to add an option where the system, for a given entry,
'   attempts to 1. fuse invalid instances with other invalid instances in pairs from the same source sheet,
'   or 2. fuse all invalid and valid instances together and see what happens.
'   These modifications would mostly affect m_MatchEntriesAcrossSheets, though.

'   Knowing this, behaviour can be summarized here
'   1. For parent-to-child data fusion :
'   Data found on sheets that have the "parent" entry format will be fused on sheets that have a matching "child" format.
'   For example :
'   A sheet "PatientInfo" only has "PNum" as an ES
'   A sheet "Admissions2014" has "PNum" and "OpDate" as ESs
'   A sheet "Operations2014" has "PNum", "OpDate" and "OpType" as ESs
'   Predictably, the user added an ECM Graph that looks like this:
'   ECM Graph
'   PNum -> PNum & OpDate
'   PNum & OpDate -> PNum & OpDate & OpType
'
'   Then, on the output sheet, all selected data from matching entries of "PatientInfo" (i.e. from the columns of PatientInfo that do appear in the
'   data fields chart on the charts sheet) will be fused to matching entries from "Admissions2014" and "Operations2014", and all data
'   from "Admissions2014" will be fused to matching entries from "Operations2014".
'
'   Of course, data that is specific to the PNum & OpDate & OpType entries on sheet "Operations2014" cannot be matched to
'   the less-specific PNum and PNum & OpDate entries from the other sheets.
'
'   As I had begun explaining this earlier, an entry that does not have any valid instance on the source
'   cannot participate to any data fusion process, as it does not "own" any data on the source sheets.
'
'   In most cases, this should not cause problems. Take for example the following case :
'   We only have a sheet "Operations2014" this time, still with "PNum", "OpDate" and "OpType" as ESs.
'   The user uses the same ECMGraph:
'   ECM Graph
'   PNum -> PNum & OpDate
'   PNum & OpDate -> PNum & OpDate & OpType
'
'   This time, the m_MatchEntriesAcrossSheets will see all the PNum & OpDate & OpType entries
'   on the sheet, and will add PNum & OpDate as well as PNum entries to the m_EntriesChart,
'   creating the logical hierarchical links between them on the go (which is the purpose of ChildrenMetInPassing)
'

Public Sub m_MatchEntriesAcrossSheets()
    
    Dim StartingDictCount As Integer: StartingDictCount = m_EntriesChart.count
    For DictIndex = 0 To StartingDictCount - 1
        Dim UIDKey As String: UIDKey = m_EntriesChart.Keys(DictIndex)
        Dim MGCount As Integer
        Dim EntryObj As clsEntry
        Set EntryObj = m_EntriesChart.Items(DictIndex)
        
        Dim ParentNodesToCheck As New Collection
        Set ParentNodesToCheck = Nothing

        Dim ECMNodeObj As clsECMNode
        Dim PNObj As clsECMNode
        Dim ChildrenMetInPassing As Scripting.Dictionary 'how poetic
        Set ChildrenMetInPassing = New Scripting.Dictionary
        ChildrenMetInPassing.RemoveAll
        
        'establish child/parent connections, regardless of instances
        If m_ECMTree.m_ConnectMap.Exists(EntryObj.m_Format) Then
            Set ECMNodeObj = m_ECMTree.m_ConnectMap(EntryObj.m_Format)
            
            For Each ParentNodeName In ECMNodeObj.m_ParentNodesToCheck.Keys
                Set PNObj = m_ECMTree.m_ConnectMap(ParentNodeName)
                Dim PNID As String: PNID = PNObj.m_IsolateInIDString(CStr(UIDKey))
                Dim ParentEntryObj As clsEntry

                If Not PNID = "" Then
                    If Not m_EntriesChart.Exists(PNID) Then
                        m_EntriesChart.Add Key:=PNID, Item:=New clsEntry
                        m_EntriesChart(PNID).m_Format = ParentNodeName
                        m_EntriesChart(PNID).m_UIDString = PNID
                    End If
                   
                    Set ParentEntryObj = m_EntriesChart(PNID)
                    EntryObj.m_Parents.Add Key:=PNID, Item:=ParentEntryObj
                    ParentEntryObj.m_Children.Add Key:=UIDKey, Item:=EntryObj
                    
                    For Each PastFormat In ChildrenMetInPassing.Keys
                        If PNObj.m_ChildNodesToCheck.Exists(PastFormat) Then
                            Dim PastID As String: PastID = ChildrenMetInPassing(PastFormat)
                            If Not ParentEntryObj.m_Children.Exists(PastID) Then
                                ParentEntryObj.m_Children.Add Key:=PastID, Item:=m_EntriesChart(PastID)
                            End If
                        End If
                    Next PastFormat
                    
                    'the way the ECM node tree is navigated could be optimized, maybe.
                    'in fact, the whole ParentNodesToCheck system (where every node above becomes a parent)
                    'could be replaced by an actual tree traversal; it would take more time and maybe more
                    'resources to implement, but it could have more applications/options for data fusion
                    'and relational filtering.
                    
                    ChildrenMetInPassing.Add Key:=ParentNodeName, Item:=PNID
                    
                End If
            Next ParentNodeName
        End If
        
        Dim Location As clsEntryInstance
        Dim SoloSheetsMatchGrp As clsMatchGroup
        
        'Then match with other valid (non-duplicate) entries.
        If EntryObj.m_ValidInstances.count > 0 Then
            
            'Match with other instances of the same entry
            m_AllMatchGrps.Add New clsMatchGroup
            MGCount = m_AllMatchGrps.count
            Set SoloSheetsMatchGrp = m_AllMatchGrps(MGCount)
            SoloSheetsMatchGrp.m_OwnerUID = UIDKey
            SoloSheetsMatchGrp.m_IsDuplicate = False
            Set EntryObj.m_MainValidMatchGrp = m_AllMatchGrps(MGCount)
    
            For Each LocIndex In EntryObj.m_ValidInstances
                 SoloSheetsMatchGrp.m_Participants.Add EntryObj.m_WhereCanIBeFound(LocIndex)
            Next LocIndex
            
            'match with all parent entries that do have valid instances somewhere on the source sheets
            If m_ECMTree.m_ConnectMap.Exists(EntryObj.m_Format) Then
                For Each ParentNodeName In ECMNodeObj.m_ParentNodesToCheck.Keys
                        Set PNObj = m_ECMTree.m_ConnectMap(ParentNodeName)
                        PNID = PNObj.m_IsolateInIDString(CStr(UIDKey))
                        Set ParentEntryObj = m_EntriesChart(PNID)
                        For Each LocIndex In ParentEntryObj.m_ValidInstances
                            SoloSheetsMatchGrp.m_Participants.Add ParentEntryObj.m_WhereCanIBeFound(LocIndex)
                        Next LocIndex
                Next ParentNodeName
            End If
        End If
        
        'Then manage the invalid (i.e. duplicate) instances of the current entry.
        'Place them in separate match groups, each containing one invalid instance.
        For Each LocIndex In EntryObj.m_InvalidInstances
            m_AllMatchGrps.Add New clsMatchGroup
            MGCount = m_AllMatchGrps.count
            m_AllMatchGrps(MGCount).m_Participants.Add EntryObj.m_WhereCanIBeFound(LocIndex)
            m_AllMatchGrps(MGCount).m_OwnerUID = UIDKey
            m_AllMatchGrps(MGCount).m_IsDuplicate = True
        Next LocIndex
        'the alternative to having one-instance match groups dedicated to duplicate instances
        'would be to add code sections to m_PrintEntriesChart and m_OutputValuesOnSheet that
        'just directly loop over each Entry's m_InvalidInstances (since the current loops
        'in these functions loop over m_AllMatchGrps).
        
    Next DictIndex
End Sub

'm_PrintEntriesChart is the first step of the m_OutputValuesOnSheet function, though
'it can also be called as a standalone procedure. It has the relatively simple
'task of printing the contents of the m_EntriesChart as it was built by
'm_BuildEntryUIDs and m_MatchEntriesAcrossSheets.

'It does so, though, by proceeding through the m_AllMatchGrps dictionary,
'which will set up matched Entry instances to appear on the same row
'of the output sheet.

'Duplicate (non-valid) entry instances will appear in brown; if the DuplicatesInSepSheet
'option is set to True, they will also appear on a separate sheet, named "[StrShName] - Duplicates".

Public Sub m_PrintEntriesChart(StrShName As String, Optional DuplicatesInSepSheet As Boolean = False)
    
    Dim OutputSheet As Excel.Worksheet
    If StrShName = "" Then
        StrShName = "FusionOutput"
    End If
    If Not SheetExists(StrShName) Then
        Sheets.Add(after:=ActiveSheet).Name = StrShName
        'if there's a way to get an equivalent statement with prettier syntax, I'm all ears
    End If
    
    m_OutputSheetName = StrShName
    
    
    If DuplicatesInSepSheet Then
        m_DuplicatesOutputSheetName = StrShName & " - Duplicates"
        If Not SheetExists(m_DuplicatesOutputSheetName) Then
            Sheets.Add(after:=Sheets(StrShName)).Name = m_DuplicatesOutputSheetName
        End If
        Sheets(m_DuplicatesOutputSheetName).Cells.ClearContents
        Sheets(m_DuplicatesOutputSheetName).Cells(2, 1).Value = "Entry UID"
        Sheets(m_DuplicatesOutputSheetName).Cells(2, 2).Value = "Where it can be found"
        Sheets(m_DuplicatesOutputSheetName).Cells.Font.Color = vbBlack
    End If

    Sheets(StrShName).Select
    ActiveSheet.Cells.ClearContents
    ActiveSheet.Cells.ClearFormats
    ActiveSheet.Cells.Font.Color = vbBlack

    Cells(2, 1).Value = "Entry UID"
    Cells(2, 2).Value = "Where it can be found"
    
    Dim CurCell As Range
    Dim i As Integer, j As Integer
    i = 2: j = 2
    For Each MatchGrp In m_AllMatchGrps
    
        If Not MatchGrp.m_IsDuplicate Or Not DuplicatesInSepSheet Then
            i = i + 1
            MatchGrp.m_RowOnOutputSheet = i
            Set CurCell = Cells(i, 1)
        Else
            j = j + 1
            Set CurCell = Sheets(m_DuplicatesOutputSheetName).Cells(j, 1)
            MatchGrp.m_RowOnOutputSheet = j
        End If
        
        LocationStr = MatchGrp.m_GetLocationString
    
        If MatchGrp.m_IsDuplicate Then
            Range(CurCell, CurCell.Offset(0, 1)).Font.Color = RGB(127, 60, 0)
        End If
    
        CurCell.Value = MatchGrp.m_OwnerUID
        CurCell.Offset(0, 1).Value = LocationStr
        
    Next MatchGrp
    
    m_LastOutputRow = i
    
End Sub

Public Function m_GetCurDFPValue(DFObj As clsDataField, ShName As String, ShNum As Integer, ColNum As Integer) As Integer
    Dim CurDFPVal As Integer
    
    If DFObj.m_DFSpecificDFPMode = 1 Then 'datafield/sheet-specific data fusion priority
        CurDFPVal = DFObj.m_DFSpecificDFPVals(ShName)
    ElseIf DFObj.m_DFSpecificDFPMode = 2 Then
        Dim SearchDFLoc As String
        SearchDFLoc = ShNum & "&" & ColNum
        'priority specific to a certain sheet on a certain column for this Internal Data Field.
        CurDFPVal = DFObj.m_DFSpecificDFPVals(SearchDFLoc)
    Else
        CurDFPVal = m_DFManager.m_SheetsChart.Items(ShNum - 1).m_DataFusionPriority
    End If
    m_GetCurDFPValue = CurDFPVal
End Function


'The m_OutputValuesOnSheet function is in itself the end goal of a data fusion process, which can be followed
'by regular Excel filtering (which doesn't require the tools I have written),
'relational filtering or adjunction of calculated columns.

'   m_OutputValuesOnSheet follows a laborious albeit straightforward process.
'   It first calls m_PrintEntriesChart on the sheet it was asked to output data on.
'   (see the description I wrote for m_PrintEntriesChart).
'
'   It then proceeds methodically:
'   Loop through every IDFN  -- check if it does have a corresponding output name
'   (IDFs without output names don't get printed on the output sheet)
'
'   For every IDF (in other words, for every column on the output sheet) :
'       Go through every match group in m_AllMatchGrps, check its row on the output sheet
'       (which m_PrintEntriesChart kindly took note of)

'       For each match group:
'           We just have zoomed on one cell of the output sheet.
'           Reminder : A match group contains entry instances; it normally contains
'           every instance of a child entry, plus every instance of the child's parents.
'           It is time to perform data fusion, that is :
'           Go through all entry instances for this match group;
'           For every instance in the match group, check if the current IDF is available for this instance
'           (i.e. if the IDF has at least one column on the sheet the instance is on).

'           So for every "currentIDF-owned column on the sheet the current instance is on",
'           at the row where the current instance is found, we have a source sheet cell; a data fusion source.
'
'
'           Then comes the data fusion process. The idea is to keep track of what we find in the source cells, while taking
'           data fusion priority values into account. Now, two modes of behavior around empty cells and data fusion priority are possible.
'           Mode 1 was coded first, and will thus be described first:

'           In EmptyValBehaviorMode 1, only values with the highest data fusion priority value for this IDF
'           and this Match Group will be kept, and differing values among those kept will be marked as conflicts.
'           Once we know that the current source cell has the highest DFP we have seen yet, we have to check whether or not
'           its value matters to us; in the current system, it will matter to us IF:
'                - it is not empty
'               OR
'                - Empty values matter for the current IDF (i.e. it has the qualifier "Empty Values Matter")
'
'           If its value matters to us, it will be added to the cache of already-seen values;
'           otherwise, it will be ignored. This means that, for an IDF devoid of the Empty Values Matter qualifier,
'           cells that remain empty in the source sheets will always get written over
'           by corresponding non-empty cells (for the same Match Group and IDF, with superior or equal DFP),
'           unless no such cells can be found.
'           For an IDF with the Empty Values Matter qualifier, such empty cells will not be
'           overwritten by equal-priority non-empty cells. Their empty value will be cached along all
'           the others.
'           New cells, as long as they matter and they have the same DFP as the highest recorded DFP yet,
'           will get added to the cache; however, if a cell with a higher DFP is found, then
'           all current cached values will be erased and a new "highest DFP value" will be set.

'
'
'
'
Public Sub m_OutputValuesOnSheet(OutputShName As String, Optional DuplicatesOnSeparateSheet As Boolean = False)
    
    Call m_PrintEntriesChart(OutputShName, DuplicatesOnSeparateSheet)
    
    Sheets(m_OutputSheetName).Select
    
    Dim IDF As Variant
    Dim NumInfoCols As Integer: NumInfoCols = 2
    Dim OutputCol As Integer: OutputCol = 2
    
    For Each IDFName In m_DFManager.m_DataFieldsChart.Keys
        Dim DFObj As clsDataField
        Set DFObj = m_DFManager.m_DataFieldsChart(IDFName)
        If Not DFObj.m_NameOnOutputSheet = "" Then
            OutputCol = OutputCol + 1
            DFObj.m_ColOnOutputSheet = DFObj.m_ColOnOutputSheet + NumInfoCols
            Dim full_locstr As String: full_locstr = ""
            If DFObj.m_WhereCanIBeFound.count > 0 Then
                For Each DFLocationKey In DFObj.m_WhereCanIBeFound.Keys
                    Dim dfsh_dfcol() As String
                    dfsh_dfcol = Split(CStr(DFLocationKey), "&")
                    full_locstr = full_locstr & "; " & m_DFManager.m_SheetsChart.Items(Val(dfsh_dfcol(0)) - 1).m_SheetName & " - " & dfsh_dfcol(1)
                Next DFLocationKey
                full_locstr = Right(full_locstr, Len(full_locstr) - Len("; "))
            End If
            
            Cells(1, OutputCol).Value = full_locstr
            Cells(2, OutputCol).Value = DFObj.m_NameOnOutputSheet
            
            If DuplicatesOnSeparateSheet Then
                Sheets(m_DuplicatesOutputSheetName).Cells(1, OutputCol).Value = full_locstr
                Sheets(m_DuplicatesOutputSheetName).Cells(2, OutputCol).Value = DFObj.m_NameOnOutputSheet
            End If
            
            
           For Each MatchGrp In m_AllMatchGrps 'a MatchGrp is a collection of entry locations
                                                            'that are meant to be fused together.
                Dim CachedValues As Scripting.Dictionary
                Set CachedValues = New Scripting.Dictionary
                Dim CurValue As Variant
                
                Dim HighestDFPVal As Integer
                Dim WinnerNumberFormat As Variant
                Dim WinnerColor As Variant
                
                If Not DuplicatesOnSeparateSheet Or Not MatchGrp.m_IsDuplicate Then
                    Set TargetOutCell = Cells(MatchGrp.m_RowOnOutputSheet, OutputCol)
                Else
                    Set TargetOutCell = Sheets(m_DuplicatesOutputSheetName).Cells(MatchGrp.m_RowOnOutputSheet, OutputCol)
                End If
                
                'i.e. each sheet has a DataFusionPriority score, in case of conflicting values.
                'Unless m_EmptyValsMatter is on, empty values are seen as informationnally insignificant
                'and can be replaced by any non-empty corresponding values.
                '
                'Depending on the m_EmptyValBehaviorMode variable, the management of empty values vs data fusion priority will differ.
                'As it was written in clsDataField:
                'in mode 0, the empty/non-empty criterion takes precedence over data fusion priority; that is, source cells considered as empty are ALWAYS ignored
                '(which means the corresponding output cell will be left empty if no other un-empty source cell is available for it).
                'in mode 1, data fusion priority takes precedence over un-emptiness; that is, cells with higher priority will ALWAYS be written over cells with lower priority,
                '

                TargetOutCell.Font.Color = vbBlack
                TargetOutCell.Value = "Mark" 'for debugging purposes
                Dim LocCount As Integer, Mode0LocCount As Integer
                LocCount = 0
                Mode0LocCount = 0
                For Each InstanceLoc In MatchGrp.m_Participants
                    Dim ShNum As Integer, RowNum As Integer
                    ShNum = InstanceLoc.m_ShNum
                    RowNum = InstanceLoc.m_RowNum
                    Dim ShObj As clsDataSheet, DFCols() As String
                    Set ShObj = m_DFManager.m_SheetsChart.Items(ShNum - 1)
                    If ShObj.m_LocalDFChart.Exists(IDFName) Then
                        DFCols = Split(ShObj.m_LocalDFChart(IDFName), ";") 'so for each entry location, we check which
                        'Internal Data Field columns are available (an entry location can only be on one sheet; and
                        'the dfcols will be the same for every entry on the sheet they are in -  so going thru
                        'available dfcols every time can be redundant in the case where many entry locations happen to be on the
                        'same sheet.
                        For Each dfcol In DFCols
                            LocCount = LocCount + 1
                            Dim TIRow As Integer, TICol As Integer
                            TICol = Val(Replace(dfcol, "_", ""))
                            TIRow = RowNum
                            Set TargetInCell = Sheets(ShObj.m_SheetName).Cells(TIRow, TICol)
                            Dim ConvertibleValue As Boolean: ConvertibleValue = True
                            If Application.WorksheetFunction.IsNA(TargetInCell) Then
                                CurValue = "#N/A on original sheet"
                                ConvertibleValue = False
                            Else
                                CurValue = TargetInCell.Value
                            End If
                            
                            If (ConvertibleValue) Then
                                For Each DFIFP In DFObj.m_SingleValueInstructions
                                    CurValue = CallByName(DFIFP.m_CallingObject, DFIFP.m_FuncStr, VbMethod, CurValue)
                                Next DFIFP
                            End If
                            
                            Dim NotEmptyCriterion As Boolean
                            NotEmptyCriterion = CStr(CurValue) <> "" Or DFObj.m_EmptyValsMatter
                            Dim CurDFPVal As Integer
                            If DFObj.m_EmptyValBehaviorMode = 1 Then 'mode 1 below
                                
                                CurDFPVal = m_GetCurDFPValue(DFObj, ShObj.m_SheetName, ShNum, TICol)
                                If (LocCount = 1) Then
                                    HighestDFPVal = CurDFPVal
                                End If
                                
                                ' Debug.Print LocCount & " " & ShObj.m_SheetName & " " & TargetInCell.Address & " : " _
                                ' & CurValue & ", " & CurDFPVal & " vs " & HighestDFPVal
                            
                                If CurDFPVal > HighestDFPVal Then
                                    CachedValues.RemoveAll
                                    HighestDFPVal = CurDFPVal
                                    CachedValues.Add Key:=CurValue, Item:=CurDFPVal
                                    WinnerNumberFormat = TargetInCell.NumberFormat
                                    WinnerColor = TargetInCell.Interior.Color
                                ElseIf CurDFPVal = HighestDFPVal Then
                                    If NotEmptyCriterion Then
                                        If Not CachedValues.Exists(CurValue) Then
                                            CachedValues.Add Key:=CurValue, Item:=CurDFPVal
                                            WinnerNumberFormat = TargetInCell.NumberFormat
                                            WinnerColor = TargetInCell.Interior.Color
                                        End If
                                    End If
                                End If
                                
                            ElseIf NotEmptyCriterion Then
                            
                                Mode0LocCount = Mode0LocCount + 1
                                CurDFPVal = m_GetCurDFPValue(DFObj, ShObj.m_SheetName, ShNum, TICol)
                                If Mode0LocCount = 1 Then
                                    HighestDFPVal = CurDFPVal
                                End If
                                If CurDFPVal > HighestDFPVal Then
                                    CachedValues.RemoveAll
                                    HighestDFPVal = CurDFPVal
                                    CachedValues.Add Key:=CurValue, Item:=CurDFPVal
                                    WinnerNumberFormat = TargetInCell.NumberFormat
                                    WinnerColor = TargetInCell.Interior.Color
                                ElseIf CurDFPVal = HighestDFPVal Then
                                    If Not CachedValues.Exists(CurValue) Then
                                        CachedValues.Add Key:=CurValue, Item:=CurDFPVal
                                        WinnerNumberFormat = TargetInCell.NumberFormat
                                        WinnerColor = TargetInCell.Interior.Color
                                    End If
                                End If
                            End If
                        Next dfcol
                    End If
                Next InstanceLoc
                
                Dim VarVal As Variant, count As Integer
                count = 0
                VarVal = ""
                For Each cv In CachedValues.Keys
                    count = count + 1
                    If count = 1 Then
                        VarVal = cv
                    Else
                        VarVal = CStr(VarVal) & " vs " & CStr(cv)
                    End If
                Next cv
            
                If count = 1 Then
                    TargetOutCell.NumberFormat = WinnerNumberFormat
                   ' TargetOutCell.Interior.Color = WinnerColor
                End If
                
                If count > 1 Then
                    TargetOutCell.Font.Color = vbRed
                End If

                TargetOutCell.Value = VarVal
                CachedValues.RemoveAll
            Next MatchGrp
        End If
        
    Next IDFName
    m_LastOutputCol = OutputCol
    
End Sub


'm_CleanupTempColumns cleans up columns used by pre-processed instructions (the instructions that were directly processed
'in EstablishEssentialData, before Entry matching and data fusion itself).

Public Sub m_CleanupTempColumns()

    'Dim ShObj As clsDataSheet
    For Each ShObj In m_DFManager.m_SheetsChart.Items
        Dim ShName As String: ShName = ShObj.m_SheetName
        Dim FirstTempCol As Integer: FirstTempCol = ShObj.m_FurthestHRColumn + 1
        Dim LastTempCol As Integer: LastTempCol = FirstTempCol + ShObj.m_NumTempColumns
        Dim LastRow As Integer: LastRow = ShObj.m_FurthestDataRow
        Dim RangeToClear As Range
        Set RangeToClear = Range(Sheets(ShName).Cells(1, FirstTempCol), Sheets(ShName).Cells(LastRow, LastTempCol))
        RangeToClear.Cells.Clear
    Next ShObj

End Sub

Public Sub m_EarlyExit()


End Sub

