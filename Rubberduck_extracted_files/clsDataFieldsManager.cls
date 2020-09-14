VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataFieldsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_DataFieldsChart As Scripting.Dictionary
Public m_SheetsChart As Scripting.Dictionary
Public m_DFOutputNameToInternal As Scripting.Dictionary
Private m_SrcDataSheetColumnCounter As Integer

Public Sub Class_Initialize()

    Set m_DataFieldsChart = New Scripting.Dictionary
    Set m_SheetsChart = New Scripting.Dictionary
    Set m_DFOutputNameToInternal = New Scripting.Dictionary
    m_SrcDataSheetColumnCounter = 0

End Sub

'm_AutoExtractDataFields takes a range of cells as its sole argument; this range of cells is expected to contain
'source sheet descriptors, of the kind found in the data field chart's header, on the designated Charts sheet.
'For each given cell, m_AutoExtractDataFields will attempt to identify source sheets from the descriptors, read this sheet's
'header rows and then output its data field names in the column just below the cell.
'
'm_AutoExtractDataFields will automatically place data field names found in the same column in the same cell,
'with the designated separator (by default, "_or_") between them. It will also place appropriate _inCol tags
'next to data field names appearing on more than a single column.

Public Sub m_AutoExtractDataFields(SheetNamesRng As Range)

    Dim ShColl As New Collection
    Dim ShCount As Integer: ShCount = 0
    For Each ShCell In SheetNamesRng
        Dim Descs() As String
        Dim Qualifiers As String: Qualifiers = ""
        Dim ShName As String
        
        Descs = Split(ShCell.Text, ";", 2)
        ShName = Descs(0)
        If UBound(Descs) = 1 Then
            Qualifiers = Descs(1)
        End If
        
        If SheetExists(ShName) Then
            ShColl.Add New clsDataSheet
            ShCount = ShCount + 1
            Call ShColl(ShCount).m_DoInit(ShCount, ShCount, ShName, Qualifiers)
            Call ShColl(ShCount).m_ExtractDataFields(ShCell.Offset(1, 0))
        End If
    Next ShCell
End Sub
Public Sub m_IncludeDataSheet(ShName As String, Optional Qualifiers As String = "")
    
    If SheetExists(ShName) Then
        If Not m_SheetsChart.Exists(ShName) Then
            m_SheetsChart.Add Key:=ShName, Item:=New clsDataSheet
            Call m_SheetsChart(ShName).m_DoInit(m_SheetsChart.count, m_SrcDataSheetColumnCounter, ShName, Qualifiers)
        Else
            MsgBox ("Warning : A sheet named " & ShName & " was already registered in this directives chart. Subsequent instances of it will be ignored.")
        End If
    Else
        MsgBox ("Warning : The specified sheet name " & ShName & "doesn't belong to an existing sheet on this workbook. It will be ignored")
    End If
        
End Sub

Public Sub m_InitFromChartsSheet(ChartsSheetName As String)
    'I initially tried to pass a worksheet as argument here,
    'but it didn't work.
    
    If Not SheetExists(ChartsSheetName) Then
        MsgBox ("Specified directives sheet name " & ChartsSheetName & " does not exist in this workbook. Aborting.")
        Err.Raise 1999
    End If
    
    Set ChartsSheet = Worksheets(ChartsSheetName)
    ChartsSheet.Select
    'first : determine Internal DF Name header : it will tell us where to find
    'the rest of the cells that make up the all-important data fields chart.
    Dim IDFHeaderCell As Range, IDFEnd As Range
    Set IDFHeaderCell = ChartsSheet.Cells.Find(What:="Internal DF Name", _
                                               LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
     
    If IDFHeaderCell Is Nothing Then
        MsgBox ("Error : No Cell containing the text : Internal DF Name could be found, so no data fields chart could be found. Aborting.")
        Err.Raise 2000
    End If
    
    'Determine the edges of the Header field (to the left and right of the IDFHeaderCell)
    Dim LeftEdgeOfDFH As Range, RightEdgeOfDFH As Range
    
    Set LeftEdgeOfDFH = IDFHeaderCell.Offset(0, -1)
    If Not IDFHeaderCell.Offset(0, -1).Text = "" Then
        Set LeftEdgeOfDFH = IDFHeaderCell.End(xlToLeft)
    End If
    
    Set RightEdgeOfDFH = IDFHeaderCell.Offset(0, 1)
    If Not IDFHeaderCell.Offset(0, 1).Text = "" Then
        Set RightEdgeOfDFH = IDFHeaderCell.End(xlToRight)
    End If
    
    Dim DataFieldsChartHeaders As Range
    Set DataFieldsChartHeaders = Range(LeftEdgeOfDFH, RightEdgeOfDFH)
    
    'Seek the first empty row below
    Dim DFCRowCount As Integer, FoundEmptyRow As Boolean
    FoundEmptyRow = False
    DFCRowCount = 0
    Do While FoundEmptyRow = False
        DFCRowCount = DFCRowCount + 1
        Dim CurRow As Range
        Set CurRow = DataFieldsChartHeaders.Offset(DFCRowCount, 0)
        If Application.WorksheetFunction.CountA(CurRow) = 0 Then
            FoundEmptyRow = True
        End If
        'just be aware that CountA counts cells containing empty text ("") -- i.e. a formula returning empty text.
        'let's just hope that an erased cell counts as an empty cell. It sure seems to!
    Loop
    
    Call m_InitFromDFCHeaders(IDFHeaderCell, DataFieldsChartHeaders, DFCRowCount)
    
End Sub

Public Sub m_InitFromDFCHeaders(IDFHeaderCell As Range, DataFieldsChartHeaders As Range, DFCRowCount As Integer)


    Dim SrcSheetDFNames As Range, InternalDFNames As Range, OutputDFNames As Range, DFInstrRng As Range, DFQualifRng As Range
    Dim IDFEnd As Range
    
    Set IDFEnd = IDFHeaderCell.Offset(DFCRowCount - 1, 0)
    
    'Then determine every column's range of cells (going from "one below the current column header"
    'to "at the level of the lowest cell in the chart")
    
    Set InternalDFNames = Range(IDFHeaderCell.Offset(1, 0), IDFEnd)
    
    Dim SrcSheetDataHeaders As Range, ODFHeaderCell As Range, DFInstrHeaderCell As Range, DFQualifHeaderCell As Range
    'SrcSheetDFNames has to be to the left of InternalDFNames.
    Set SrcSheetDataHeaders = Range(DataFieldsChartHeaders.Cells(1, 1), IDFHeaderCell.Offset(0, -1))
    Set SrcSheetDFNames = Range(DataFieldsChartHeaders.Cells(1, 1).Offset(1, 0), IDFEnd.Offset(0, -1))
    
    Set ODFHeaderCell = DataFieldsChartHeaders.Find(What:="Output DF Name")
    If Not ODFHeaderCell Is Nothing Then
        Set OutputDFNames = Range(ODFHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, ODFHeaderCell.Column))
    End If
    
    Set DFInstrHeaderCell = DataFieldsChartHeaders.Find(What:="Instructions")
    If Not DFInstrHeaderCell Is Nothing Then
        Set DFInstrRng = Range(DFInstrHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, DFInstrHeaderCell.Column))
    End If
    
    Set DFQualifHeaderCell = DataFieldsChartHeaders.Find(What:="Qualifiers")
    If Not DFQualifHeaderCell Is Nothing Then
        Set DFQualifRng = Range(DFQualifHeaderCell.Offset(1, 0), Cells(IDFEnd.Row, DFQualifHeaderCell.Column))
    End If
    
     'So now the Data Fields Chart has been found and delimited.
    
    'Determine basic data sheet parameters for each input data sheet.
    'Each input sheet has its own cell in the data fields chart header.
    For Each SrcSheetDesc In SrcSheetDataHeaders
        m_SrcDataSheetColumnCounter = m_SrcDataSheetColumnCounter + 1
        
        Dim Descs() As String
        Descs = Split(SrcSheetDesc.Text, ";", 2) 'output no more than 2 strings.
        'i.e. Split("A;B;C;D", ";", 3) -> Result: {"A", "B", "C;D"}
        'The expected format is NameOfInputSheet; Sheet qualifiers
        Dim QualifStr As String: QualifStr = ""
        If UBound(Descs) = 1 Then
            QualifStr = Descs(1)
        End If
        
        'The Val function didn't work here, as it stops as once as it finds a non-digit character
        Call m_IncludeDataSheet(Descs(0), QualifStr)
    Next SrcSheetDesc

    'Construction of the m_DataFieldsChart, from the main chart written on the charts sheet
    '(i.e. the one with Source Data Sheet names, "Internal DF Name", "Output DF Name", "Qualifiers" and "Instructions")
    
    'The first step is to go through every Internal DF Name and scan the surrounding rows for data pertaining to it.
    'An important thing to note is that every IDF name has to be unique; if the system finds an IDF name it has already
    'seen while going through the Internal DF Name column, it will simply skip it.
    
    'On the other hand, if a row on Internal DF Name column is left empty, no IDF Name will be registered for it and
    'any Source sheet DF names on the same row will be ignored (well, unless they appear on another row which does have an IDFName on it).
    'An output DF Name left on a row without an IDF Name will also not appear on the output sheet, since the Output function loops
    'through the IDFNames in the m_DataFieldsChart.
    
    Dim IDFName As String, ODFName As String
    Dim ColOnOutputSheetCounter As Integer: ColOnOutputSheetCounter = 0
    For Each IDFCell In InternalDFNames
        ChartIndex = IDFCell.Row - InternalDFNames.Row + 1 '[Range].Row returns the same value as [Range].Cells(1,1).Row
        If Not IDFCell.Text = "" Then
            If Not m_DataFieldsChart.Exists(IDFCell.Text) Then
                IDFName = IDFCell.Text
                
                If Not ODFHeaderCell Is Nothing Then
                    ODFName = OutputDFNames.Cells(ChartIndex, 1).Value
                Else
                    ODFName = IDFName
                End If
                
                m_DataFieldsChart.Add Key:=IDFName, Item:=New clsDataField
                m_DataFieldsChart(IDFName).m_InternalOfficialName = IDFName
                m_DataFieldsChart(IDFName).m_IndexOnDFChart = ChartIndex
                If Not ODFName = "" Then
                    m_DataFieldsChart(IDFName).m_NameOnOutputSheet = ODFName
                    ColOnOutputSheetCounter = ColOnOutputSheetCounter + 1
                    m_DataFieldsChart(IDFName).m_ColOnOutputSheet = ColOnOutputSheetCounter
                End If
                 
                If Not DFInstrHeaderCell Is Nothing Then
                    m_DataFieldsChart(IDFName).m_Instructions = DFInstrRng.Cells(ChartIndex, 1).Text
                End If
                If Not DFQualifHeaderCell Is Nothing Then
                    m_DataFieldsChart(IDFName).m_Qualifiers = DFQualifRng.Cells(ChartIndex, 1).Text
                End If
                
                If Not m_DFOutputNameToInternal.Exists(ODFName) Then
                    m_DFOutputNameToInternal.Add Key:=ODFName, Item:=IDFName
                End If
            Else
                MsgBox ("Warning : " & IDFCell.Text & " was already specified as an internal data field name on row " _
                        & m_DataFieldsChart(IDFCell.Text).m_IndexOnDFChart & " and will be ignored on row " & ChartIndex)
            End If
        End If
    Next IDFCell
    
    
    'The next step is to build the "on-src-sheet data field name" to "Internal data field name" dictionary, and
    'related data structures. Originally, this was through a TransChart, but the peculiarities of data fields
    'finally made me opt for a tailor-made procedure, which is contained in m_BuildDataFieldDict.
    'See the notes on this sub (above) for details on data field routing peculiarities.
     
    Call m_BuildDataFieldDict(SrcSheetDFNames, InternalDFNames)
          

End Sub


'So here's the write-up on assigning source data fields to internal data fields (which then get transferred to the output sheet).

'Now, we know that, in the data field chart, empty IDF rows are ignored, as well as IDF rows where
'the IDFName does not appear for the first time. This is taken care of below too, at the beginning the row loop.

'But what is possible to specify for IDFNames that are valid?

'The basic way to use the data field chart is to place every source sheet name on its column to the left of the Internal DF Name column;
'and then, for each sheet, place data field names from this sheet in that sheet's column.
'Data field names that happen to be on the same row in the will get "routed" to the Internal DF Name from the corresponding row.
'The system will search for the source data field name on the header row(s) of the sheet it is attributed to and, if it does find it,
'it will associate the matching IDFName to the source data sheet column it found the source DFName on.


'Now, after I had elaborated this, I realized I had to add features in order to support other types of issues users might run into.

'MANY COLUMNS FROM THE SAME SOURCE SHEET SHARING AN IDENTICAL FIELD NAME

'The first is the possibility of having data fields from the same sheet that share the same name.
'If the user writes a data field name in a sheet's column and that dfname does appear on said sheet more than once,
'the system will only take the first occurence of the dfname on the sheet.
'In order to which occurence of that dfname to take, the user will have to use the _inCol[Column Number] tag.
'This will make the system directly associate the desired column number to the IDFName that appears on the same row.

'The "normal" way to use _inCol[ColNum] is to append it to the source data field name. For example, say we take for granted that two "Diagnosis" fields
'exist on sheet "PatientListA" -one on column 5, and one on column 10. Then (with some imagination) the data field chart might look like this :

'   PatientListA        |   Internal DF Name
'   Diagnosis_inCol5    |   Diagnosis1
'   Diagnosis_inCol 10  |   Diagnosis2
'
'Note that any number of whitespaces can be placed between _inCol and the actual digits indicating the column number.

'The funny thing is that it is not necessary for the given source df name (here, "Diagnosis") to actually match what appears
'in the source sheet's ("PatientList1") header rows at the column number given by _inCol. For example, column 10 on PatientListA
'could actually be called "PathologicalDx", or "Capibara population", or what have you, and the system would still assign
'column 10 from PatientListA to the IDFName "Diagnosis2". It would simply send a warning MsgBox and keep plowing through.

'What's even funnier, and what you might have guessed, is that since the _inCol indication takes precedence over
'the name attached to it, this name can actually be entirely omitted, leaving only an "_inCol[Column number]" as
'an indication for which source data field to take.

'MORE THAN ONE HEADER ROW ON A GIVEN SOURCE SHEET

' In my short but eventful career here, I have seen my share of sheets having more than one header row,
' where many different field names may appear on the same column.
' As I said, when the system encounters a field name (without an _inCol tag) in a source sheet's column, it searches for the
' dfname on all the sheet's header rows; as such, it is not necessary to give all the names associated with a given source column
' for this column to be correctly processed by the system. As long as the correct number of header rows was specified
' in the Source data sheet's cell (i.e. [Source data sheet name]; [number of header rows]; [Qualifiers]), everything should be fine.


'But this leads us to the next point :

'ROUTING DIFFERENT DFNAMES/COLUMNS FROM THE SAME SOURCE SHEET TO THE SAME IDFNAME
'
'By sepratating them with a designated separator (by default, it is "_or_", which has a fair chance of never appearing in a field name),
'it is possible to write many source field names --with or without _inCol tags, or just _inCol tags, in the same cell of the
'data fields chart, so that they are all routed to the same IDFName. As you might have guessed, these source data field identifiers
'can belong to the same column on their source sheet, or to different columns on their source sheet, which gives rise to interesting possibilities.
'
'
'EXTRACTION OF DATA FIELDS FROM A SOURCE SHEET
'
'   You now might be telling yourself that it would be tedious to transcribe all a source sheet's data field names on a neatly
'   written column, ready to be routed to adjacent columns for other source sheets and internal DF Names.
'   To ease this process, I have two options to offer you:
'
'   1- Copy all the header rows of the source sheet, and paste them on the charts sheet with the transpose option.
'       A quick method, but especially suited for "simple" source sheets, where there is ideally only one header row,
'       and few (or no) columns sharing the same name.
'
'   2- Use the tailor-made m_AutoExtractDataFields sub from the clsDFSystem class.
'       It is especially suited for complicated source sheets, with more than one header row and/or many columns sharing the same name.
'       See its description/usage above.
'       m_AutoExtractDataFields will automatically place data field names found in the same column in the same cell,
'       with the designated separator (by default, "_or_") between them. It will also place appropriate _inCol tags
'       next to data field names appearing on more than a single column.
'       After that, if you want to only have one data field name for each row of the source sheet column,
'       just do Find "_or_*" -> Replace with "" (but that is common Excel knowledge, I didn't invent that).

'   In the (rare, I hope) cases where the user wants to have many different columns,
'   from the same source sheet routed to the same IDFName, he or she will have to do so manually, though.

'SPECIFYING DATA FUSION PRIORITY IN THE MOST EXTREME AND SPECIFIC WAY (MODE 2)
'
'   Data fusion is kind of a subject in itself, and will be mostly discussed where it actually kicks in action,
'   in the m_OutputValuesOnSheet function. It offers the user the option to choose whether whole sheets,
'   or certain columns from certain sheets, have a priority over others when it comes to comparing their data
'   together -- data sources having equal priorities will be compared together, unless a data source with higher
'   priority comes along and just dumps them down the cliff.
'
'   Let it be known that data fusion priority values can be specified in three ways:
'   Mode 0 - For each source sheet, for all IDFs on it (specify it in the source sheet descriptor, in the header of the data fields chart)
'   Mode 1 - Specific to each source sheet, for a given IDFName (specify it in the internal data field qualifiers, in the data field chart)
'   Mode 2 - For specific columns from each source sheet, for a given IDFName
'       (specify it with tags attached to data field indicators in the source sheet range of the data field chart)
'
'   For now, I have not yet programmed data fusion priority that would depend on
'   the values assigned to entries themselves -- i.e. the actual data on the sheets.
'   The highest level of precision goes to "which columns/which IDFName" we are processing.
'
'   Every method listed above (0 through 2) takes precedence over the one preceding it,
'   as it gets more precise. For most cases, Mode 1 should be sufficient;
'   Mode 2 can actually only make a difference in the case of many columns from the same source
'   sheet being routed to the same IDF.
'   On IDFs for which no field-specific DFP was given (not in mode 1 nor 2), the system will default to the (Mode 0) DFP values attributed to
'   each sheet as a whole; these values default to 0.
'
'   Values specified through mode 2 are processed in m_BuildDataFieldDict, and it would be topical to present it here.
'   HOW IT WORKS
'   Use the data fusion priority tag. Its format is _DFP=[value], with or without spaces around the equal sign.
'

Public Sub m_BuildDataFieldDict(SrcSheetDFNames As Range, InternalDFNames As Range, Optional Separator As String = "_or_")
    'taking into account that both the m_DataFieldsChart and the m_SheetsChart are established by this point

Dim SplitCompMethod As VbCompareMethod
SplitCompMethod = vbBinaryCompare
If Len(Separator) > 1 Then
       SplitCompMethod = vbTextCompare
       'vbTextcompare is not case-sensitive, but works with multi-char separators
End If

For dfsh = 1 To m_SheetsChart.count
'Scan each sheet separately

    Dim SrcShName As String: SrcShName = m_SheetsChart.Items(dfsh - 1).m_SheetName
    Dim NumHR As Integer: NumHR = m_SheetsChart.Items(dfsh - 1).m_NumHeaderRows
    Dim ColOnDFChart As Integer: ColOnDFChart = m_SheetsChart.Items(dfsh - 1).m_ColOnDFChart
    
    'Debug.Print ("Now in sheet " & SrcShName & " which has " & NumHR & " header row(s)")
    For RowNum = 1 To SrcSheetDFNames.Rows.count
    'For every row on this sheet's column
        Dim IDFName As String: IDFName = InternalDFNames.Cells(RowNum, 1).Text
        If Not IDFName = "" Then
            If m_DataFieldsChart(IDFName).m_IndexOnDFChart = RowNum Then
                'ignore rows without internal df names, and ignore
                'rows where the IDFName does not appear for the first time
                Dim DFObj As clsDataField: Set DFObj = m_DataFieldsChart(IDFName)
                Set SrcDFCell = SrcSheetDFNames.Cells(RowNum, ColOnDFChart)
                Dim dfnames() As String
                dfnames = Split(SrcDFCell.Text, Separator, , SplitCompMethod)
                'separate dfnames found in the same cell, based on the given separator (by default, "_or_")
                For Each dfname In dfnames
                    
                    dfname = Trim(dfname)
                    Dim SpecificFPForThisCol As Integer: SpecificFPForThisCol = 0
                    Dim dfcol As Integer: dfcol = -1
                    
                    Dim LocalRegex As Object: Set LocalRegex = CreateObject("VBScript.RegExp")
                    LocalRegex.Global = True: LocalRegex.MultiLine = False: LocalRegex.IgnoreCase = False
                    'LocalRegex will be used to check for tags attributed to the current dfname
                    
                    LocalRegex.Pattern = "_DFP\s*=\s*(\d+)"
                    'Begin by testing if specific FP was given.
                    If LocalRegex.Test(dfname) Then
                        Set captures = LocalRegex.Execute(dfname)
                        
                        SpecificFPForThisCol = captures(0).submatches(0)
                        'take only the first match's value (if there are many, which there shouldn't be)
                        
                        DFObj.m_DFSpecificDFPMode = 2
                        'The presence of any dfname with an _FP tag will automatically set the data fusion priority mode
                        'for this IDF to 2; and nothing can set it back to 1, nor 0.
                        
                        dfname = LocalRegex.Replace(dfname, "")
                        'remove all occurrences of the match in dfname.
                        
                    End If
                    
                    'Next : our objective for the given dfname is to find which column number it is on.
                    
                    LocalRegex.Pattern = "_inCol\s*(\d+)"
                    'Test if a column number was explicitly specified.
                    If LocalRegex.Test(dfname) Then
                        Set Matches = LocalRegex.Execute(dfname)
                        dfcol = Val(Matches(0).submatches(0))
                        'Take the given column number and keep it for later.
    
                        Dim JustTheName As String: JustTheName = LocalRegex.Replace(dfname, "")
                        Dim CheckIfReallyThere As Range
                        Dim RangeToCheck As Range: Set RangeToCheck = Range(Sheets(SrcShName).Cells(1, dfcol), Sheets(SrcShName).Cells(NumHR, dfcol))
                        Set CheckIfReallyThere = RangeToCheck.Find(What:=JustTheName, LookAt:=xlWhole, LookIn:=xlValues)
                        
                        'Search for the given dfname in the portion of the source sheet where we expect it to be
                        '(i.e. on the given column number, in the header rows)
                        
                        If CheckIfReallyThere Is Nothing Then
                            Debug.Print ("Warning : no data field " & JustTheName & " found in column " & dfcol & " of sheet " & SrcSheetName & _
                                                        ". Keeping this col number anyway.")
                            'As I said before, if the given data field name isn't found in the expected column, a warning message will be issued,
                            'but the system will move on with the given column number.
                        End If
                        
                    Else
                        'if no recognizable "_inCol" tag was found, just search for the given dfname all across the header rows
                        'and keep the first column number you find.
                    
                        Dim SingleLocOnSrcSh As Range
                        Set SingleLocOnSrcSh = Sheets(SrcShName).Rows("1:" & NumHR).Find(What:=dfname, after:=Sheets(SrcShName).Cells(1, 1), _
                                                                                            LookAt:=xlWhole, LookIn:=xlValues)
                        
                        If Not SingleLocOnSrcSh Is Nothing Then
                            dfcol = SingleLocOnSrcSh.Column
                        Else
                            Debug.Print ("No data field " & dfname & " could be found on sheet " & SrcShName & ".")
                            'if no column was found with this dfname, then we have a problem.
                        End If
                    End If
                    
                    If dfcol > 0 Then 'i.e. if a valid column was found,
                        'the time has come to check if we already stumbled on the same column in the same sheet.
                        'To this end, the m_WhereCanIBeFound dictionary contains keys of the format [sheet number]&[column number].
                        'A same-key case for this dictionary will happen when
                        'a same column has many of its names listed for the same IDFName in the DFChart.
                        
                        If Not DFObj.m_WhereCanIBeFound.Exists(dfsh & "&" & dfcol) Then
                            
                            DFObj.m_WhereCanIBeFound.Add Key:=dfsh & "&" & dfcol, Item:=New clsDFLocationInfo
                            Dim curindex As Integer: curindex = DFObj.m_WhereCanIBeFound.count - 1
                            DFObj.m_WhereCanIBeFound.Items(curindex).m_ShNum = dfsh
                            DFObj.m_WhereCanIBeFound.Items(curindex).m_ColNum = dfcol
                            
                            
                            'Once the all-important m_WhereCanIBeFound dictionary is managed,
                            'we also have to update the m_SheetsChart accordingly.
                            'Both are used internally at different points of the data processing.
                            
                            'Every clsDataSheet object has a m_LocalDFChart, which indicates
                            'where every IDF can be found on the sheet.
                            'So the m_LocalDFChart keys are IDFNames, and the values are column numbers,
                            'with a special syntax mostly for cases where multiple columns from the same
                            'sheet are routed to the same IDF.
                            'The syntax is thus : _12_;_1_;_26_ etc.
                            
                            If Not m_SheetsChart.Items(dfsh - 1).m_LocalDFChart.Exists(IDFName) Then
                                m_SheetsChart.Items(dfsh - 1).m_LocalDFChart.Add Key:=IDFName, Item:="_" & dfcol & "_"
                            Else
                                Dim AlreadyExistingCols As String: AlreadyExistingCols = m_SheetsChart.Items(dfsh - 1).m_LocalDFChart(IDFName)
                                m_SheetsChart.Items(dfsh - 1).m_LocalDFChart(IDFName) = AlreadyExistingCols & ";" & "_" & dfcol & "_"
                            End If
                            
                            'SpecificFPForThisCol is 0 by default. If the system already knows that the current DFObj's m_DFSpecificDFPMode is 2,
                            'then it is a simple matter to assign the correct data fusion priority value to each source column/dfname it encounters.
                            'However, the only way the DFSpecificDFPMode can be set to 2 is by the system encountering a column identifier with
                            'the _(D)FP tag; as such, dfnames that preceded the first one bearing a _(D)FP tag will not be caught in the
                            'following If clause, and have to be taken care of in the closing loop at the end of the function.
                                
                            If DFObj.m_DFSpecificDFPMode = 2 Then
                                DFObj.m_DFSpecificDFPVals.Add Key:=DFObj.m_WhereCanIBeFound.Items(curindex).m_ConvertToString(), Item:=SpecificFPForThisCol
                                
                            End If
                        
                        Else
                            'If a same column is indeed written more than once for the same IDF,
                            'we have to handle the case of a different Data Fusion Priority value being given on each of the times it is written.
                            
                            'The system will take the last value expressly specified - remember that a dfname without an _FP tag, for an IDF using mode 2 DFP,
                            'will be assigned an FP value of 0 (see the last loop below).
                            
                            'FOR EXAMPLE, say PatientListA has two header rows. On column 3, we have "OperationDate" on HeaderRow 1, and "cSurgeryDate" on HeaderRow 2.
                            'We have auto-extracted the data fields, which explains why they appeared like this :
                            '
                            '           PatientListA                   |   Internal DF Name
                            '   OperationDate _or_ cSurgeryDate        |      OpDate
                            '
                            'So say we are a bit silly and write this :
                            '
                            '           PatientListA                             |   Internal DF Name
                            '   OperationDate_FP=1 _or_ cSurgeryDate_FP=2        |    OpDate
                            '
                            '   The fusion priority value for column 3 of the PatientListA sheet, for the OpDate IDF  will be 2.
                            '
                            'But say we are right-minded people and we just didn't bother removing the "_or_ cSurgeryDate" :
                            '
                            '           PatientListA                        |   Internal DF Name
                            '   OperationDate_FP=1 _or_ cSurgeryDate        |    OpDate
                            '
                            '   The fusion priority value for column 3 of the PatientListA sheet, for the OpDate IDF will be 1.
                            '
                            
                            If SpecificFPForThisCol <> 0 And DFObj.m_DFSpecificDFPMode = 2 Then
                                
                                Dim KeyVal As String
                                KeyVal = DFObj.m_WhereCanIBeFound.Items(curindex).m_ConvertToString()
                                If DFObj.m_DFSpecificDFPVals.Exists(KeyVal) Then
                                    m_DataFieldsChart(IDFName).m_DFSpecificDFPVals(KeyVal) = SpecificFPForThisCol
                                    'yup, replacing m_DataFieldsChart(IDFName) by DFObj yielded a compilation error. The more you know
                                End If
                                    
                                Debug.Print ("New DFP Value for " & dfname & ", so for Location " & DFObj.m_WhereCanIBeFound.Items(curindex).m_ShNum & " " & _
                                DFObj.m_WhereCanIBeFound.Items(curindex).m_ColNum & " : " & SpecificFPForThisCol)
                            End If
                        End If
                    End If
                    
                Next dfname
                
                'If we have encountered a dfname with an _FP tag, we have to go back through all the others for
                'which we did not specify a fusion priority value through an FP tag.
                
                If DFObj.m_DFSpecificDFPMode = 2 Then
                    For Each Location In DFObj.m_WhereCanIBeFound.Items
                        If Not DFObj.m_DFSpecificDFPVals.Exists(Location.m_ConvertToString()) Then
                            DFObj.m_DFSpecificDFPVals.Add Key:=Location.m_ConvertToString(), Item:=0
                        End If
                    Next Location
                End If
            End If
        End If
    Next RowNum
Next dfsh



End Sub


