VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_InternalOfficialName As String
Public m_NameOnOutputSheet As String
Public m_IndexOnDFChart As Integer
Public m_WhereCanIBeFound As Scripting.Dictionary

Public m_Instructions As String
Public m_Qualifiers As String

Public m_IsEntrySpecifier As Boolean
Public m_EmptyValsMatter As Boolean
Public m_EmptyValBehaviorMode As Integer

Public m_DFSpecificDFPMode As Integer
Public m_DFSpecificDFPVals As Object 'data field-specific data fusion priority values.

Public m_IsMultiChoice As Boolean

Public m_SingleValueInstructions As Collection
Public m_RangeInstructions As Collection
Public m_Units 'if there are, like mL and such
Public m_PrefNumberFormat As Variant 'the preferred number format for this data field

Public m_ColOnOutputSheet As Integer

Public Sub Class_Initialize()

Set m_WhereCanIBeFound = New Scripting.Dictionary 'quite useful, it turned out. it will now be a dictionary of clsDFLocationInfo objects

Set m_SingleValueInstructions = New Collection
Set m_RangeInstructions = New Collection

Set m_DFSpecificDFPVals = New Scripting.Dictionary

m_IsEntrySpecifier = False
m_EmptyValsMatter = False
m_UsesDFSpecificDFP = False
m_DFSpecificDFPMode = 0
' 0 for not used (i.e. will default to the current sheet's data fusion priority), 1 for BySheet, 2 for ByLocation(i.e. specific column on specific sheet)
m_EmptyValBehaviorMode = 0
'in mode 0, the empty/non-empty criterion takes precedence over data fusion priority; that is, source cells considered as empty are ALWAYS ignored
'(which means the corresponding output cell will be left empty if no other un-empty source cell is available for it).
'in mode 1, data fusion priority takes precedence over un-emptiness; that is, cells with higher priority will ALWAYS be written over cells with lower priority,
m_IsMultiChoice = False

m_ColOnOutputSheet = -1

End Sub

'Data field qualifiers are to be placed in the aptly-named "Qualifiers" column of the data field chart;
'many qualifiers can be assigned to the same IDF. Qualifiers can be put in any order, and even without
'separators between them, as the system just scans for qualifier strings it can recognize through the whole
'string found in the qualifiers cell. So far, adding separators between qualifiers didn't seem necessary.

'
'LIST OF POSSIBLE QUALIFIERS
'
'   ENTRY SPECIFIER
'   Writing "Entry Specifier" or "ES" in an IDF's qualifier cell will bestow Entry Specifier status
'   upon this IDF. This means that data found in the source sheet columns pertaining to this IDF
'   will be used to determine each entry's identity and to match entries across source sheets.
'   (an entry being the set of data that appears on a given row).
'   The user is expected to provide entry specifiers that allow to uniquely identify every entry
'   on each sheet, such that no entry has more than one instance on the same sheet.
'   Having more than one instance of a given entry on the same sheet is considered a problem; the
'   instances are considered as duplicates.
'
'   At a higher level, entry specifiers can be used in conjunction with an Entry Connect Map
'   to link data across sheets that don't share the same set of Entry Specifiers, but for which
'   heritage relationships can be described. ECM ultimately leads to the possibility of
'   Relational Filtering. I'll talk about that later on.
'
'   But for what matters most here, the system simply updates the relevant
'   data structures when it finds an Entry Specifier IDF : the clsDFSystem's m_AllEntrySpecifiers dictionary,
'   its m_UsesEntrySpecifiers flag (which clsDFSystem's m_BuildEntryUIDs needs to know)
'   and each source data sheet object's own list of local entry specifiers.

'   EMPTY VALUES MATTER
'   Writing "Empty Values Matter" or "Empty Cells Matter" will indicate to the system
'   that, when completing data fusion for the current IDF,
'   data cells left empty have a meaning and are not to be automatically filled in by equal-priority
'   non-empty data cells. The role of this parameter can be better understood in the
'   m_OutputValuesOnSheet function, and the corresponding write-up (I will eventually make) for it.
'
'   MULTICHOICE
'   Writing "MultiChoice" indicates that the current IDF belongs to a series
'   of multi-choice checkbox fields. This is mostly useful for transferring
'   data to other applications, like RedCap. I actually plan to transfer
'   the MultiChoice functionality from the clsDFSystem to a new PostProcessing program,
'   designed to read an Outputted sheet and add columns to it for Excel functions, calculations
'   and other fields that solely depend on already existing fields.
'
'   DATA FUSION PRIORITY, MODE 1
'
'   As I explained in the notes to clsDFSystem's m_BuildDataFieldDict, there are three ways to specify
'   data fusion priority. One of these, dubbed "Mode 1", is by writing data fusion priority
'   values in the Qualifiers cell of a given IDF. This lets the user specify DFP values specific to
'   the current IDF, each source sheet being assigned its own DFP value. On IDFs for which
'   no field-specific DFP was given (not in mode 1 nor 2), the system will default to the (Mode 0) DFP values attributed to
'   each sheet as a whole; these values default to 0.
'
'   To specify Mode 1 DFP values for a given IDF, the user writes something in this format :
'   either "FP", "DFP", or "Priority", followed by a ":" (with or without surrounding whitespaces), followed by a series
'   of "sheetname = value" groups. These groups can be separated by anything.
'   EXAMPLE OF VALID DFP (mode 1) QUALIFIERS
'   Priority : PatientDataA = 2, All2012 = 1
'   FP:PatientDataA= 2; All2012 =1
'
'   If a source sheet name is omitted from the "sheetname = value" enumeration, its DFP value
'   for the current IDF will be set to 0.
'
'
'

Public Sub m_ReadDataFusionQualifiers(ParentDFS As clsDFSystem)

    If InStr(m_Qualifiers, "Entry Specifier") + InStr(m_Qualifiers, "ES") > 0 Then
        m_IsEntrySpecifier = True
        ParentDFS.m_AllEntrySpecifiers.Add Key:=m_InternalOfficialName, Item:=New clsEntrySpecifier
        ParentDFS.m_AllEntrySpecifiers(m_InternalOfficialName).m_IDFName = m_InternalOfficialName
        ParentDFS.m_UsesEntrySpecifiers = True
        For Each ShObj In ParentDFS.m_DFManager.m_SheetsChart.Items
            If ShObj.m_LocalDFChart.Exists(m_InternalOfficialName) Then
                ShObj.m_LocalEntrySpecifiers.Add m_InternalOfficialName
            End If
        Next ShObj

    End If
    If InStr(LCase(m_Qualifiers), "empty values matter") + InStr(LCase(m_Qualifiers), "empty cells matter") > 0 Then
        m_EmptyValsMatter = True
    End If
    
    If InStr(LCase(m_Qualifiers), "dfp before non-empty") > 0 Then
        m_EmptyValBehaviorMode = 1
    End If
    
    'TODO: test for more cases of multiple qualifiers including a DFP one.
    
    Dim DFPFinder As Object
    Set DFPFinder = CreateObject("VBScript.RegExp")
    DFPFinder.Global = True
    DFPFinder.MultiLine = False
    DFPFinder.IgnoreCase = True
    DFPFinder.Pattern = "(D?FP|Priority)\s*\:\s*(.+)"
    If DFPFinder.Test(m_Qualifiers) And Not m_DFSpecificDFPMode = 2 Then
        'Data fusion priority mode 2 is taken into consideration over mode 1
        '(and mode 0, which is the default one for every IDF)
        'init the values
        
        For Each ShObj In ParentDFS.m_DFManager.m_SheetsChart.Items
            m_DFSpecificDFPVals.Add Key:=ShObj.m_SheetName, Item:=0
        Next ShObj
        
        Set Matches = DFPFinder.Execute(m_Qualifiers)
        Dim StrWithPriors As String
        StrWithPriors = Matches(0).submatches(1)
        DFPFinder.Pattern = "(\w(\w|\s)+\w)\s*=\s*(\d+)"
        Set Allumettes = DFPFinder.Execute(StrWithPriors)
        For Each Allumette In Allumettes
            Dim ShName As String: ShName = Allumette.submatches(0)
            Dim DFPValue As Integer: DFPValue = Val(Allumette.submatches(2))
            If m_DFSpecificDFPVals.Exists(ShName) Then
                m_DFSpecificDFPMode = 1
                m_DFSpecificDFPVals(ShName) = DFPValue
                'Debug.Print ("DFPriority value for field " & m_InternalOfficialName & " on sheet " & ShName & " is " & DFPValue & ".")
            End If
        Next Allumette
    End If
    

End Sub

'And then for each IDF : compilation of the data field instructions
           'At first, I thought of managing all sorts of functions/modificators here,
           'but I think I will only keep those useful for data fusion (i.e. mostly
           'uniformizing data so that it can be fused more logically).
           'I'll add the other ones (i.e. calculating fields from other fields, data transfers
           'where the number of input columns differs from the number of output columns)
           'for another class/program : clsDPPSystem (Data Post Processor)
           '
           
           'Instructions can be divided in three main groups, according to type of behaviour:
           '
           '1- Instructions which can be processed during the data fusion process. These are
           '    usually processed on only one value at a time, which makes it possible to process them
           '    in m_BuildEntryUIDs (for entry specifier data fields) and m_OutputValuesOnSheet (for all data fields)
           
           '    For now, the Group 1 instructions are "TextReplace(TC_Name_)" and "FullCellReplace(_TC_Name_)"
           '    Both are placed in their IDF's m_SingleValueInstructions collection.
           
           '2- Instructions that are processed in m_EstablishEssentialData, before
           '   entries are built (m_BuildEntryUIDs), matched (m_MatchAcrossSheets) and their data fused (m_OutputValuesOnSheet).
           '    These instructions could be calculated one at a time, but the fact that they are pre-processed makes it
           '    possible to process them in bulk, on whole ranges, which is usually faster in Excel.
           '   Since they are processed before entry matching and fusion, these instructions have to cache their results
           '   somewhere -- and I didn't want to create huge dictionaries just for that, so I cached them in helper columns
           '   at the end of the source sheets. These columns can be cleared in clsDFSystem's m_CleanTempColumns
           '   (meant to be called after m_OutputValuesOnSheet).
           '
           '    For now, the Group 2 instructions are "QuickTextReplace(TC_Name_)", "QuickCellReplace(TC_Name_)" and "FixDate()".
           '    They are all processed directly upon reading the instruction, one source data column at a time.
           '    As such, they do not have to be stored in a collection for further use.
           '    QuickCellReplace is exactly similar to FullCellReplace, but is much faster (actually, I just removed FullCellReplace);
           '    QuickTextReplace is almost similar to TextReplace, but is also much faster; it only lacks support
           '    for regular expressions and output terms containing input terms.
           '
           '
           '3- Instructions that can only be processed after data fusion is complete (or on a single
           '    sheet which will not be fused with any other). This is mostly the case of instructions
           '    that depend on values taken in other, already-existing and clearly determined fields.
           '    These can be processed in whole columns at once. These instructions will eventually
           '    be transferred to clsDPPSystem and its own interface system.


Public Sub m_ReadDataFusionInstructions(ParentDFS As clsDFSystem)
    Dim DFInstructions() As String
    Dim InstrCounter As Integer
    DFInstructions = Split(m_Instructions, "&")
    For Each DFInstr In DFInstructions
        DFInstr = Trim(DFInstr)
        Dim DFInstrObj As clsDataFieldInstruction
        Dim FuncStr As String
        Dim FirstParenth As Integer: FirstParenth = InStr(DFInstr, "(")
        Dim SecondParenth As Integer: SecondParenth = InStr(DFInstr, ")")
        Dim TCName As String, Args As String
        If DFInstr Like "SlowTextReplace_SingleCell(*)" Then
            TCName = Mid(DFInstr, FirstParenth + 1, SecondParenth - FirstParenth - 1)
            If ParentDFS.m_TranslationCharts.Exists(TCName) Then
                m_SingleValueInstructions.Add New clsDataFieldInstruction
                InstrCounter = m_SingleValueInstructions.count
                m_SingleValueInstructions(InstrCounter).m_FuncStr = "m_TranslateOneString_RegExpComboThenNormal"
                Set m_SingleValueInstructions(InstrCounter).m_CallingObject = ParentDFS.m_TranslationCharts(TCName)
                
            Else
                MsgBox ("Warning : given TransChart name " & TCName & " does not belong to a recognized TransChart.")
            End If
        ElseIf DFInstr Like "TextReplace(*)" Or DFInstr Like "CellReplace(*)" Or DFInstr = "FixDate()" Or DFInstr Like "SlowTextReplace(*)" Then
            Args = Mid(DFInstr, FirstParenth + 1, SecondParenth - FirstParenth - 1)
            Dim TCNameAndREMode() As String
            TCName = Args
            
            Dim InstrName As String: InstrName = Left(DFInstr, FirstParenth - 1)
            If ParentDFS.m_TranslationCharts.Exists(TCName) Or InstrName = "FixDate" Then
                
                'Actually, it can all be done here. Let's go!
                For Each DFLoc In m_WhereCanIBeFound.Items
                    Dim DFLocSheetName As String: DFLocSheetName = ParentDFS.m_DFManager.m_SheetsChart.Keys(DFLoc.m_ShNum - 1)
                    
                    Dim DFLocShObj As clsDataSheet: Set DFLocShObj = ParentDFS.m_DFManager.m_SheetsChart(DFLocSheetName)
                    DFLocShObj.m_NumTempColumns = DFLocShObj.m_NumTempColumns + 1
                    
                    Dim DFLocNumHR As Integer: DFLocNumHR = DFLocShObj.m_NumHeaderRows
                    Dim DFLocLastRow As Integer: DFLocLastRow = DFLocShObj.m_FurthestDataRow
                    Dim DFLocLastCol As Integer: DFLocLastCol = DFLocShObj.m_FurthestHRColumn + _
                                                                DFLocShObj.m_NumTempColumns
                    
                    
                    Dim InputRange As Range, OutputRange As Range
                    Set InputRange = Range(Sheets(DFLocSheetName).Cells(DFLocNumHR + 1, DFLoc.m_ColNum), _
                                                Sheets(DFLocSheetName).Cells(DFLocLastRow, DFLoc.m_ColNum))
                    
                    Set OutputRange = Range(Sheets(DFLocSheetName).Cells(DFLocNumHR + 1, DFLocLastCol), _
                                                Sheets(DFLocSheetName).Cells(DFLocLastRow, DFLocLastCol))
                    
                    If InstrName = "TextReplace" Then
                        Call ParentDFS.m_TranslationCharts(TCName).m_QuickTranslateCells(InputRange, OutputRange, 0)
                    ElseIf InstrName = "CellReplace" Then
                        Call ParentDFS.m_TranslationCharts(TCName).m_QuickTranslateCells(InputRange, OutputRange, 1)
                    ElseIf InstrName = "SlowTextReplace" Then
                        
                        Call ParentDFS.m_TranslationCharts(TCName).m_TranslateCells(InputRange, OutputRange, 0)
                    ElseIf InstrName = "FixDate" Then
                        Call FixAllDatesInRange(InputRange, OutputRange)
                    End If
                    
                    'add the translation/conversion as a temp column at the end of the sheet
                    'and then replace the column name on the appropriate charts
                    DFLocShObj.m_LocalDFChart.Item(m_InternalOfficialName) = Replace(DFLocShObj.m_LocalDFChart(m_InternalOfficialName), _
                                                             "_" & DFLoc.m_ColNum & "_", "_" & DFLocLastCol & "_")
                    'this is what prompted me to place every colnum between underscores in the m_LocalDFChart
                    Dim NewLoc As New clsDFLocationInfo
                    NewLoc.m_ShNum = DFLoc.m_ShNum
                    NewLoc.m_ColNum = DFLocLastCol
                    If m_DFSpecificDFPMode = 2 Then
                        Dim SaveValue As Integer
                        SaveValue = m_DFSpecificDFPVals(DFLoc.m_ConvertToString())
                        m_DFSpecificDFPVals.Remove (DFLoc.m_ConvertToString())
                        m_DFSpecificDFPVals.Add Key:=NewLoc.m_ConvertToString(), Item:=SaveValue
                    End If
                    
                    DFLoc.m_ColNum = DFLocLastCol
                    
                Next DFLoc
            Else
                MsgBox ("Warning : given TransChart name " & TCName & " does not belong to a recognized TransChart.")
            End If
                
        ElseIf Not DFInstr = "" Then
            MsgBox ("Warning : The data field instruction " & DFInstr & " for IDF " & m_InternalOfficialName & ", on row " & m_IndexOnDFChart & " was not recognized.")
        End If
    Next DFInstr
End Sub

Public Sub m_ReadDataProcessingQualifiers(ParentDPS As clsDPSystem)


    If InStr(m_Qualifiers, "MultiChoice") Then
        m_IsMultiChoice = True
    End If

End Sub


