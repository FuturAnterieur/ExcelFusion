VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_Name As String
Public m_TransChartIn As Range
Public m_TransChartOut As Range

Private m_AllInTermsArray As Collection
Private m_TCIIndexArray As Collection
Public m_Dictionary As Scripting.Dictionary
Public m_RegexpTermsDict As Scripting.Dictionary
'a simplified, direct InputTerm To OutputIndex Dictionary, without regex escape-backslashes
Private m_regex As Object

Public m_TermSeparator As String
Public m_SplitCompMethod As VbCompareMethod

Public m_CaseSensitive As Boolean
Public m_UseSpecialChars As Boolean

Public Sub Class_Initialize()

    m_CaseSensitive = False
    m_UseSpecialChars = False
    
    Set m_regex = CreateObject("VBScript.RegExp")
    m_regex.Global = True
    m_regex.MultiLine = False
    m_regex.IgnoreCase = True
    
    Set m_Dictionary = New Scripting.Dictionary
    
    Set m_RegexpTermsDict = New Scripting.Dictionary

End Sub
Public Sub m_ApplyToField(FieldToTranslate As Range)

    Call AutoTranslateOrConvertMk5(FieldToTranslate, m_TransChartIn, m_TransChartOut, m_CaseSensitive, m_UseSpecialChars)
    'The AToCMk5 sub will have to be corrected with the revisions brought in the portable versions below.

End Sub


'CREATING A TRANSCHART ON THE CHARTS SHEET (i.e. the sheet sent to clsDFSystem's m_EstablishEssentialData) :
            'How to specify a TransChart for the system :
            '1- Select the TransChart's name, which has to be of the form TC_[anything]_
            '2- Place the TransChart's in a cell the charts sheet (the sheet this DFSystem is currently parsing through).
            '3- Write your TransChart just below the cell where you placed the TransChart's name.
            '   See the "Writing a TransChart" manual below, (still) in the clsTransChart module.
            '   Here, in EstablishEssentialData for the clsDFSystem, it is not (currently) possible to specify
            '   the TCI and TCO's ranges individually. The TCO is always assumed to be the last column of the TransChart,
            '   and the TCI's range depends on the "NotIncludeTCOInTCI" option detailed below.
            
            '4- The cell directly to the right of the TransChart's name is where you can specify options for your TransChart.
            
            'Below is a list of the possible options :
            
            'Option to specify TransChart size through format [ColNum]x[RowNum] (example : 2x40) -  DONE
            'take note that if the size is kept unspecified, the system will try to find the TransChart's extent
            'by finding the last non-empty cell down and then the last non-empty cell to the right (see the algorithm below).
            'As such, if no gaps are present between the lower and right ends of the TransChart and another block of content,
            'the system will include non-TransChart material into the TransChart and that will lead to bugs.
            
            'Option to specify TransChart CaseSensitive and SpecialChar Sensitive
            '- DONE : "Match Case" and "Match Special Characters"
            
            'Option to specify preferred translation mode for TransChart (text-search or full cell)
            '- DONE in the data field instructions, it's actually better that way, as the same TransChart can be used in the two manners
            
            'Option to specify is the TCO is not to be included inside the TCI (it is included in the TCI by default) :
            '"NotIncludeTCOInTCI"

Public Sub m_BuildDictFromDescCell(CurTCCell As Range)
    
    Dim SizeIndicFinder As Object
    Set SizeIndicFinder = CreateObject("VBScript.RegExp")
    SizeIndicFinder.Global = True
    SizeIndicFinder.MultiLine = False
    SizeIndicFinder.Pattern = "(\d+)x(\d+)"
    'it should also be possible to only specify the number of columns
    
    Dim CurTCOptionsCell As Range
    Set CurTCOptionsCell = CurTCCell.Offset(0, 1)
    
    Dim TCI As Range, TCO As Range
    Dim MatchCase As Boolean: MatchCase = False
    Dim MatchSC As Boolean: MatchSC = False
    
    Dim ChartFirstCell As Range
    Set ChartFirstCell = CurTCCell.Offset(1, 0)

    Dim LastRow As Integer, LastCol As Integer, FirstRow As Integer, FirstCol As Integer
    
    If Not SizeIndicFinder.Test(CurTCOptionsCell.Text) Then
        
        If ChartFirstCell.Text = "" Then
            LastRow = ChartFirstCell.Row
        Else
            LastRow = CurTCCell.End(xlDown).Row
        End If
        Dim CurCol As Integer
        LastCol = 0
        For Each FCCell In Range(ChartFirstCell, Cells(LastRow, ChartFirstCell.Column))
            If FCCell.Offset(0, 1).Text = "" Then
                CurCol = ChartFirstCell.Column + 1
            Else
                CurCol = FCCell.End(xlToRight).Column
            End If
            
            If CurCol > LastCol Then
                LastCol = CurCol
            End If
        Next FCCell
    Else
        Set Matches = SizeIndicFinder.Execute(CurTCOptionsCell.Text)
        Dim NbCols As Integer: NbCols = Val(Matches(0).submatches(0))
        Dim NbRows As Integer: NbRows = Val(Matches(0).submatches(1))
        LastRow = ChartFirstCell.Row - 1 + NbRows
        LastCol = ChartFirstCell.Column - 1 + NbCols
    End If
    
    If InStr(CurTCOptionsCell.Text, "NotIncludeTCOInTCI") Then
        Set TCI = Range(ChartFirstCell, Cells(LastRow, LastCol - 1)) ' in this case, it is outside it
    Else
        Set TCI = Range(ChartFirstCell, Cells(LastRow, LastCol))  'in this case, the TCO is inside the TCI
    End If
    
    Set TCO = Range(Cells(ChartFirstCell.Row, LastCol), Cells(LastRow, LastCol))
    
    If InStr(CurTCOptionsCell.Text, "Match Case") > 0 Then
        MatchCase = True
    End If
    
    If InStr(CurTCOptionsCell.Text, "Match Special Characters") > 0 Then
        MatchSC = True
    End If
        
    If CurTCCell.Text = "TC_NormalizeOps_" Then
        Debug.Print "debug time. TCI:"
        For Each TCICell In TCI
            Debug.Print TCICell.Address & " : " & TCICell.Text
        Next TCICell
        Debug.Print "TCO:"
        For Each TCOCell In TCO
            Debug.Print TCOCell.Address & " : " & TCOCell.Text
        Next TCOCell
    End If
        
    Call m_BuildDictionary(TCI, TCO, CS:=MatchCase, USC:=MatchSC)

End Sub

'How to write the TransChart itself :
'BASIC PRESENTATION
'A TransChart works according to a set of quite simple rules; but as time passed, I added more and more features, so it got kind of complicated.
'I even removed some features when I realized that they weren't really useful and made the code more complicated for no good reason (see
'the TooComplicatedTransChartCode module).

'But basically, the idea is to have a simple Excel table establishing links of equivalence and/or conversion between matching terms.

'To speak in general terms, each column is meant to be dedicated to a similar category of terms,
'and each row is meant to be dedicated to a same "meaning".

'Examples are always much easier to understand at first. Consider the following Excel range of cells :

'    (french terms)              (english terms)         (STS codes [made up for this example])
' Lobectomie                | Lobectomy            |        32012
' Biopsie ganglionnaire     | Lymph Node Biopsy    |        30955
' Médiastinoscopie          | Mediastinoscopy      |        31382

'The TransChartIn (or TCI for short) can be any set of adjacent columns;
'a multi-column range can be given as a TransChartOut, but only the first column of this range will matter.

'So let's say we take Column 1 as TCI and Column 2 as TCO :
'when parsing through the chart, the system will consider "Lobectomie", "Biopsie ganglionnaire" and "Médiastinoscopie"
'as input terms and respectively assign them row numbers 1, 2 and 3 in what is called the
'm_Dictionary, or translation dictionary (a dictionary data structure in VBA).
'In a nutshell, this is what m_BuildDictionary does.

'When asked to perform a conversion on a range of cells or on a single string (in one of the many conversion functions found below m_BuildDictionary),
'the system will search for these three terms (either in text-mode or whole-cell mode) and, when it does find a match,
'it will replace them with the term found at the corresponding row number inside the TCO (1- "Lobectomy", 2-"Lymph Node Biopsy", 3-"Mediastinoscopy").

'If we specify Columns 1 and 2 as TCI, and Column 3 as TCO, then the input terms will be all the french terms and their english equivalents;
'it will go "Lobectomie" - 1, "Biopsie ganglionnaire" - 2, "Médiastinoscopie" - 3, "Lobectomy" - 1, "Lymph Node Biopsy" -2, "Mediastinoscopy" - 3.
'When asked to translate a text, the system will look for these 6 possibilities, and if there is a match, it will replace the match with
'the term found at the corresponding row number in the TCO, which now happens to be the STS code list.

'It is totally possible (and even sometimes advisable) to place different input terms sharing the same translated output
'in different rows of the TCI. On the other hand, if a single input term appears on many rows of the TCI, only its first
'appearance will be included in the translation dictionary and taken into account for the conversion process.

'Whole-cell replacement (known as Mode 1 as you will see if you read the code below) is generally simpler; it only
'matches and replaces a cell's text if the whole cell's contents corresponds to a term in the TCI.
'Text-wise replacement (known as Mode 0) is more complex; it goes through the cell's text and matches any terms it finds in the TCI.
'Let it be known that the case of some TCI terms containing other shorter TCI terms is taken care of by sorting the TCI terms
'in descending order of length in the m_BuildDictionary function.
'Another mode, Text To Multi Choice Range, is detailed later on, further below.


'FEATURES
'So for the TransChart construction, what are the features supported by m_BuildDictionary?

'CASE SENSITIVENESS AND SPECIAL CHARACTER SENSITIVENESS
'It is simple to specify if the conversion engine is to be sensitive to case (uppercase vs lowercase letters) and special characters
'(i.e. any accented characters vs their unaccented counterpart, i.e. é vs e). By default, both of these sensitivenesses are turned off,
'making conversion operations generally simpler. For example, if a range of cells to translate contains many variations of "epidermoide", like
'"Épidermoïde", "Epidermoide", "épidermoide", etc., and all these instances are to be translated to "Squamous cell carcinoma",
'a single translation row (with only one TCI term) is necessary to convert it :
' epidermoide | Squamous cell carcinoma

'Both of the sensitivenesses can be individually turned on. Case sensitiveness will make the conversion engine only look for words with the
'uppercase/lowercase pattern, and the same principle applies to special character sensitiveness. The way to switch these on and off is through
'optional arguments CS (for case sensitive) and USC (for uses special characters) when calling the m_BuildDictionary function.
'As such, the sensitivenesses apply equally to all the terms in the TCI and are specified at the time of building the internal data structures from the written
'TransChart on the Excel sheet.
'
'

'MANY TERMS IN THE SAME CELL
'It is possible to place several distinct terms in a single TCI (or TCO) cell.
'These terms have to be separated by a separating character/expression that, by definition,
'cannot be found inside any one of the terms. That is why I have settled on the "code-expression" (or tag)
'"_or_" as the default separator. For quite a while, I used the pipe ("|") (which also often means "or" in programming)
'but I realised that doctors sometimes used it in their Excel sheets as actual data, so I had to resort
'to something that I thought would never be used outside of TransCharts. Take note that it is possible
'to specify a different separator when calling the m_BuildDictionary function.

'Placing many terms in the same TCI cell and splitting them with "_or_" will simply assign them the same row number,
'and they will yield the same output term (taken at the corresponding TCO row) when converted.
'It will thus have the same effect as if they were placed in equivalent rows in different columns in the same TCI.
'Placing many terms in the same TCI cell can be particularly useful when many TCI terms have the same meaning but
'separating them in different columns is not justified and would be too cumbersome.

'And as a side note, if a TCO cell contains more than one term, only the first term will be used to replace the corresponding TCI term(s).
'This job is handled by m_GetFirstTCOTerm, a utility function used by pretty much all conversion functions.

'TERMS USING A REGULAR EXPRESSION
'It is also possible to use regular expressions as TCI terms (and appropriate regular expression substitutions as TCO terms)
'Regular expressions really are a subject of their own, and many websites are devoted to explaining how to use them.
'Many flavors of regular expressions exist, with small differences from one to the other, but the general principles remain the same.
'Naturally, this program uses the implementation of regular expressions that comes with VBA.
' A reference sheet for matching patterns with this implementation can be found here :
' https://msdn.microsoft.com/en-us/library/ae5bf541(v=vs.100).aspx
' (of course, many sites like http://www.regular-expressions.info/ and http://www.rexegg.com/ explain regular expressions in detail
' and their reference sheets will do the job without any problem for most cases, even if they were not made for the Microsoft implementation).

'As it was mentioned above, TransChartOut regular expression terms can used regex substitution codes, mostly beginning
'with the "$" character. You can check them out on any regular expression reference sheet,
'https://docs.microsoft.com/en-us/dotnet/standard/base-types/substitutions-in-regular-expressions
'(using substitutions - the "output term" part)

'In the TCI proper, a term has to be designated as a regular expression term for the regular expression interpreter to kick in.
'The way to designate it as such is to write the "code-expression" (or tag) "_regexp_" right before it.
'The system will then remove "_regexp_" from it and store it in a dictionary that is distinct
'from the dictionary reserved for "normal", non-regex terms.

'Terms using regular expressions will only be used in text-wise (Mode 0) translation, and not full-cell (Mode 1) translation.
'In Mode 0, all terms from the regular expression dictionnary are fully processed before normal terms,
'so that regular expression results can then be translated by the normal-terms dictionnary.

'A _regexp_ tag takes only effect for its current term;
'it does not spread to other terms in the same cell, which are split with the _or_ tag.

'LEAVING TERMS UNTRIMMED
'By default, all separate TCI terms are trimmed (to trim a string of characters means to remove leading and trailing whitespaces from it, i.e.
' "  a word   " becomes "a word" when trimmed). In most cases, this makes perfect sense : if, by mistake, we leave a space after
'writing "Lobectomie " in the example above, we still want any match of "Lobectomie" to be translated to "Lobectomy".
'The actual cells to be translated, though, do not have to be trimmed when using text-wise (Mode 0) translation, since
'we are searching through their text; but cell's contents have to be trimmed in Mode 1, since only the whole cell's text is checked at once.
'
'When creating the TransChart, the user has the option to specify to leave a term untrimmed; as I have just explained, though, this untrimmed
'term will only properly work when using the TransChart for Mode 0 conversion. All the user has to do is to write the "code-expression" (tag)
' "_notrim_" before the term to be left untrimmed.
'
'The feature applies only to the terms it was appended to; it makes it possible to translate terms with leading and/or trailing whitespaces,
'or terms entirely made out of whitespaces.


Public Sub m_BuildDictionary(TCI As Range, TCO As Range, Optional CS As Boolean = False, Optional USC As Boolean = False, Optional Sep As String = "_or_")

    Set m_TransChartIn = TCI
    Set m_TransChartOut = TCO
    m_TransChartIn.Parent.Select
    
    Dim UnsortedDict As Scripting.Dictionary
    Set UnsortedDict = New Scripting.Dictionary
    
    m_regex.IgnoreCase = Not CS
    m_UseSpecialChars = USC
    m_CaseSensitive = CS
    
    m_TermSeparator = Sep
    
    m_SplitCompMethod = vbBinaryCompare
    If Len(m_TermSeparator) > 1 Then
            m_SplitCompMethod = vbTextCompare
            'vbTextcompare is not case-sensitive, but works with multi-char separators
    End If
    
    Dim CurTerms As String
    Dim CurTermsArray() As String
    
    Dim TCIRow As Integer, TCICol As Integer
    Dim TCICell As Range
    Dim FirstRowToAdd As Integer
    
    For TCICol = 1 To m_TransChartIn.Columns.count
        
        For TCIRow = 1 To m_TransChartIn.Rows.count
        
            Set TCICell = m_TransChartIn.Cells(TCIRow, TCICol)
            'Debug.Print (TCICell.Text)
            
            CurTerms = TCICell.Text
            
            If m_CaseSensitive = False Then
                CurTerms = LCase(CurTerms)
            End If
        
            If m_UseSpecialChars = False Then
                CurTerms = StripAccent(CurTerms)
            End If
            
            CurTermsArray = Split(CurTerms, m_TermSeparator, , SplitCompMethod)
    
            Dim i As Integer
            For i = 0 To UBound(CurTermsArray)
                If InStr(CurTermsArray(i), "_notrim_") = 0 Then
                    CurTermsArray(i) = Trim(CurTermsArray(i))
                Else
                    'writing "_notrim_" right before a TransChartIn term will tell
                    'the system to leave it untrimmed, thus giving the user
                    'the possibility to convert terms containing trailing whitespaces (i.e. " a word  "),
                    'or even terms only made out of whitespaces.
                    'The "_notrim_" itself and any spaces BEFORE it will be removed.
                    'All other spaces will be left untrimmed.
                
                    m_regex.Pattern = "\s*_notrim_"
                    
                    CurTermsArray(i) = m_regex.Replace(CurTermsArray(i), "")
                End If
                
                If InStr(CurTermsArray(i), "_regexp_") = 0 Then
                    
                     If Not UnsortedDict.Exists(CurTermsArray(i)) Then
                        UnsortedDict.Add Key:=(CurTermsArray(i)), Item:=TCIRow
                    End If
                    
                Else
                    'that means it is a term using regular expressions.
                    CurTermsArray(i) = Replace(CurTermsArray(i), "_regexp_", "")
                    
                    'below is one of the crazier features.
                    'basically, it lets the user plug a series of terms (to be considered as a regex OR-group, i.e. red|blue|white)
                    'INSIDE the current regex-using term.
                    'Of course, all these alternatives will be assigned to the same output term, although regex substitutions
                    'like $"number of group in input pattern" can yield interesting solutions.
                    
                    'the way to use it is thus:
                    'inside the regexp, you write _range_(the excel address of the desired range)_
                    
                    'so for example you want the resulting regex to become :
                    '(.*)(Segmentectomie|Lingulectomie)
                    'but Segmentectomie and Lingulectomie (basically, all the terms to be or'ed by the pipes) are in range A2:A3. So you can do
                    '(.*)(_range_A2:A3_)
                    'and that should do the trick.
                    
                    Dim AddrPos As Integer, EndAddr As Integer
                    AddrPos = InStr(CurTermsArray(i), "_range_")
                    
                    If AddrPos > 0 Then
                        AddrPos = AddrPos + Len("_range_")
                        EndAddr = InStr(AddrPos, CurTermsArray(i), "_")
                        Dim AddrString As String
                        AddrString = Mid(CurTermsArray(i), AddrPos, EndAddr - AddrPos)
                        Dim DicoInputRng As Range
                        Set DicoInputRng = Range(AddrString)
                        Dim DicoInputTerms As String, CurDicoTerms As String
                        DicoInputTerms = ""
                        For Each DICell In DicoInputRng
                            DicoInputTerms = DicoInputTerms & DICell.Text & "|"
                        Next DICell
                        DicoInputTerms = Replace(DicoInputTerms, "| ", "|")
                        DicoInputTerms = Left(DicoInputTerms, Len(DicoInputTerms) - 1)
                        'sorting the terms in order of length would be useless,
                        'since using \b (word boundary) markers in the regex automatically takes
                        'care of the issue.
                        CurTermsArray(i) = Replace(CurTermsArray(i), "_range_" & AddrString & "_", DicoInputTerms)
                    End If
                    
                    If Not m_RegexpTermsDict.Exists(CurTermsArray(i)) Then
                        m_RegexpTermsDict.Add Key:=(CurTermsArray(i)), Item:=TCIRow
                    End If
                    'In previous versions of this program, TCITerms using regular expressions
                    'were put in the same dicionary as normal, non-regex terms.
                    'But I realised that putting them in their own dictionary gave
                    'me more flexibility.
                    
                End If
            Next i
        Next TCIRow
    Next TCICol
    
    'Next up : Sorting this list in descending order of length.
    'Could have begun sorting while I was going through the m_TransChartIn range.
    'But for now, we'll make do with a simple sorting algorithm.
    
    'The idea is that by sorting them in this order,
    'we can just go through the generated list
    'without having to worry about input terms being contained
    'inside other input terms.
    
    Dim x As Integer, y As Integer
    Dim TempTxt1 As String, TempTxt2 As String
    Dim KeySorterArr() As String
    
    ReDim KeySorterArr(UnsortedDict.count)
    
    Dim it As Integer: it = 0
    For Each curKey In UnsortedDict.Keys
        KeySorterArr(it) = curKey
        it = it + 1
    Next curKey
    
    For x = 0 To UnsortedDict.count - 2
        For y = x + 1 To UnsortedDict.count - 1
            If Len(KeySorterArr(y)) > Len(KeySorterArr(x)) Then
                TempTxt1 = KeySorterArr(x)
                TempTxt2 = KeySorterArr(y)
                KeySorterArr(x) = TempTxt2
                KeySorterArr(y) = TempTxt1
            End If
        Next y
    Next x
    
    'And finally : copying the sorted array into the m_Dictionary.
    'The m_Dictionary will thus contain pairs of "TCITerm - row in TransChart" elements
    For j = 0 To UnsortedDict.count - 1
        m_Dictionary.Add Key:=KeySorterArr(j), Item:=UnsortedDict(KeySorterArr(j))
        'Debug.Print ("New Entry  : " & KeySorterArr(j) & "-" & UnsortedDict(KeySorterArr(j)))
    Next j

End Sub
Public Function m_GetFirstTCOTerm(RowIndex As Integer) As String
    'this function is used to get the TransChartOut term contained at row number RowIndex.
    'There may be more than one term in this TransChartOut cell; in that case,
    'only the first one is taken, and that is why the function's name isn't simply m_GetTCOTerm.

    Dim ReplaceBy As String
    Dim AllTCOText As String
    Dim TCOTerms() As String
    AllTCOText = m_TransChartOut.Cells(RowIndex, 1).Text
    
    If Not AllTCOText = "" Then
        TCOTerms = Split(AllTCOText, m_TermSeparator, , m_SplitCompMethod)
        ReplaceBy = TCOTerms(0)
    Else
        ReplaceBy = ""
    End If
    m_GetFirstTCOTerm = ReplaceBy

End Function

'Now, we can see that with the same m_Dictionary, several conversion options are possible :
' - Text replacement (look at the text within each cell and replace any matching string separately) (called Mode 0 here)
' - Full cell replacement (only match all the cell's text at once) (called Mode 1 here)
' - Search cell text for terms and translate them to checks in a multiple choice grid

'And for each of these conversion options, there is the possibility of doing it
'for a single cell, or for a whole range of cells.

' The functions below basically systematize these different options.

Public Sub m_TranslateCells(InputCells As Range, OutputCells As Range, Optional Mode As Integer = 0)

    'm_TranslateCells is used to call a single-cell translation function (selected through the Mode options) on a whole range of cells.
    'Mode 0 is for Text Replacement; Mode 1 is for full cell replacement.
    'The "Search cell text for terms and translate them to checks in a multiple choice grid" is accessed
    'through another set of functions; namely, m_ConvertTextToMultiChoiceRange and m_ConvertTextToMRCForColumn.
    'm_TranslateCells it is generally slower than the aptly-named m_QuickTranslateCells below, although
    'some mode 0 features are only available when translating each cell one by one : that is, using
    'regular expressions, and properly managing situations where output terms contain input terms that
    'are still to be translated.
    
    'Up to now, and as far as I know, there are no difference in conversion behaviour
    'between m_TranslateCells and m_QuickTranslateCells when using Mode 1.
    'Knowing that m_QuickTranslateCells is pretty much always faster, I do have to
    'admit that Mode 1 could be removed from m_TranslateCells. I might do this sometime.
    
    'In addition to this, m_TranslateCells was also designed to support situations
    'where InputCells and OutputCells do not have the same size.

    Dim OutputStr As String, InputStr As String
    Dim InRow As Integer, InCol As Integer, OutRow As Integer, OutCol As Integer

    If InputCells.Rows.count = OutputCells.Rows.count And InputCells.Columns.count = OutputCells.Columns.count Then
        For InRow = 1 To InputCells.Rows.count
            For InCol = 1 To InputCells.Columns.count
                InputStr = InputCells.Cells(InRow, InCol).Text
                If Mode = 0 Then 'text-wise translation
                    OutputStr = m_TranslateOneString(InputStr)
                Else 'Full-Cell translation, by direct m_Dictionary/m_HeaderlessDictionary search
                    OutputStr = m_RetrieveDicoTerm(InputStr)
                End If
                OutputCells.Cells(InRow, InCol).Value = OutputStr
            Next InCol
        Next InRow
    Else
        'in cases where the size of the Input and Output fields differ,
        'the principle is to stop at the "end" of each field, and fill with
        'identical addresses from there.
        Dim ChangedCell As Boolean
        For OutCol = 1 To OutputCells.Columns.count
            
            InCol = OutCol
            If InCol > InputCells.Columns.count Then
                InCol = InputCells.Columns.count
            End If
            
            For OutRow = 1 To OutputCells.Rows.count
                ChangedCell = True
                InRow = OutRow
                If OutRow > InputCells.Rows.count Then
                    InRow = InputCells.Rows.count
                    ChangedCell = False
                End If
            
                If ChangedCell = True Then ' refresh the output string
                    InputStr = InputCells.Cells(InRow, InCol).Text
                    If Mode = 0 Then 'text-wise translation
                        OutputStr = m_TranslateOneString(InputStr)
                    Else 'Full-Cell translation
                        OutputStr = m_RetrieveDicoTerm(InputStr)
                    End If
                End If
                
                OutputCells.Cells(OutRow, OutCol).Value = OutputStr
            Next OutRow
        Next OutCol
        
    End If


End Sub
Public Sub m_QuickTranslateCells(InputCells As Range, OutputCells As Range, Optional Mode As Integer = 0)
    'The quick version that uses Range.Replace on the whole range at once.
    'The regex terms found in the TCI will be ignored (they would have been used in m_TranslateCells with Mode 0).
    'also, the mode 0 algorithm is simpler than the one used in One-by-one cell translation,
    'as it doesn't handle cases where some of the input terms can be found inside some of the output terms
    '(which is fortunately pretty rare)
     
     Application.DisplayAlerts = False
     Dim ReplaceBy As String
     If InputCells.Rows.count = OutputCells.Rows.count And InputCells.Columns.count = OutputCells.Columns.count Then
        
        InputCells.Copy Destination:=OutputCells
        
        If m_UseSpecialChars = False Then
           Call StripAccentOnRange(OutputCells)
        End If

        Dim LookAtMode As XlLookAt
        If Mode = 0 Then
            LookAtMode = xlPart
        Else
            LookAtMode = xlWhole
        End If

        For Each TCITerm In m_Dictionary.Keys
            ReplaceBy = m_GetFirstTCOTerm(m_Dictionary(TCITerm))
            OutputCells.Replace What:=TCITerm, Replacement:=ReplaceBy, LookAt:=LookAtMode, MatchCase:=m_CaseSensitive
        Next TCITerm
    End If
    
    Application.DisplayAlerts = True
End Sub

Public Function m_RetrieveDicoTerm(InputVal As Variant) As Variant
    'This function is a basic m_Dictionary searcher, here taking
    'the effort to get the first TransChartOut term pointed at
    'in the m_Dictionary for the InputVal given as an argument.
    
    Dim ResRowNum As Integer: ResRowNum = m_RetrieveTCIRow(InputVal)
    
    If ResRowNum > -1 Then
        m_RetrieveDicoTerm = m_GetFirstTCOTerm(ResRowNum)
    Else
        m_RetrieveDicoTerm = InputVal
    End If
        
End Function


Public Function m_RetrieveTCIRow(InputVal As Variant) As Integer
    'The dictionary-searching part of m_RetrieveDicoTerm.
    'Can be used in a standalone fashion too.
    
    Dim InputStr As String
    InputStr = CStr(InputVal)
    
    Dim ResRowNum As Integer
    
    If m_UseSpecialChars = False Then
        InputStr = StripAccent(InputStr)
    End If
    
    If m_CaseSensitive = False Then
        InputStr = LCase(InputStr)
    End If
    
    InputStr = Trim(InputStr)
    'Trim is necessary here, as the whole cell/text is searched at once.
    'minor todo : add a feature to take trimmed/untrimmed terms into account.
    'kinda harder, because trim/untrim is not uniform for the whole TransChart, but specified term-wise.
    'also, the trim/untrimmed term feature is quite unlikely to be used in dictionary (cell by cell) mode.
    
    If m_Dictionary.Exists(InputStr) Then
        ResRowNum = m_Dictionary(InputStr)
    Else
        ResRowNum = -1
    End If
    
    m_RetrieveTCIRow = ResRowNum

End Function

    'Here is a great place to explain the Two Pass system used to translate strings.
    
    'First off, translation would always happen in at least two steps :
    'a run-through of the terms using regular expressions, and a run-through (in descending order
    'of length) of the normal terms.
    
    'But a special case made me opt for a two-pass system, at least in this "slow", comprehensive
    'mode of operation (as opposed to m_QuickTranslateCells).
    
    'This special case is when a normal output term (from the TransChartOut) contains a normal input term
    'that is to be translated afterwards.
    'For (a simple) example, let's say we have the following TransChart :
    
    ' Habiletés -> Expertises
    ' Draperie -> Veil
    ' Expert -> Master
    'here, once a young, bustling and innocent translating system has converted "Habiletés" to "Expertises",
    'it will later see the "Expert" in "Expertises", and translate it, yielding "Masterises".
    
    'I thus had to devise a way to "protect" already-translated words from being translated again.
    'After experimenting with a couple of ideas, I settled (maybe erroneously, who knows) on a two-pass system where :
    'on the first pass-thru of the m_Dictionary, terms are searched in the adjusted string; if they are matched,
    'they are marked as "dead-space", i.e. replaced by non-printable characters that (hopefully) won't ever be matched
    'by other text. Then, in the Clean String (the one that kept the case and special chars of the Input String, and
    'that will eventually be outputted as the Result String), match positions are taken from the adjusted string and,
    'at the these positons, the original matched text is replaced with unique ID markers, following
    'the format "_#[match index]_" (basically, a Term To Replace number), and the rows of the matched terms themselves
    'are placed in the TermsToReplace collection.
    
    'The "_#_" markers are the solution I have found to keep track of which terms were already found, as well
    'as the position in the string where they were found, at the same time.
    
    'On the second pass, the "_#_" markers are replaced by the actual corresponding terms from the TransChartOut.
    
    'For terms using regular expressions, the process is largely similar, although it is important to note
    'that regex terms appear in the m_RegexpTermsDict in the order in which they were specified in the
    'inputted TransChartIn, and will thus be prioritized in that order.


    'Mode 0 - now the only mode
    ' regexp mode separated : Regex are treated after Normal dictionary terms
    ' The sequence is Normal Pass 1, Regex Pass 1, Normal Pass 2, Regex Pass 2.
    
    'Mode 1 - removed, now in TooComplicatedForTransChartCode
    'regexp mode combined : Regex results can be translated by the Normal dictionary.
    ' The sequence is Regex Pass 1, Regex Pass 2, Normal Pass 1 and Normal Pass 2.
    
    'I removed it because it was easier to reproduce its effects by just chaining several TransCharts together,
    'while it was not really possible to reproduce Mode 0's effects with the chaining method.
    
Public Function m_TranslateOneString(InputStr As String) As String

    Dim TermsToReplace As New Collection
    Set TermsToReplace = Nothing
    
    Dim REReplacementText As New Collection
    Set REReplacementText = Nothing
    
    Dim FTTWithAdj As String, FTTClean As String
    FTTWithAdj = InputStr
    FTTClean = InputStr
    Dim FTTResult As String: FTTResult = FTTWithAdj
    
    If InputStr <> "" Then
        If m_UseSpecialChars = False Then
            FTTWithAdj = StripAccent(FTTWithAdj)
        End If
        
        If m_CaseSensitive = False Then
            FTTWithAdj = LCase(FTTWithAdj)
        End If
        
        'trim is not really necessary here, as
        'the text is searched through.
        Dim TTRCounter As Integer
        Dim TCITerm As Variant
        Dim TCITermIndex As Integer
    
        Dim NumCharRight As Integer
        
        TTRCounter = 1
        For Each TCITerm In m_Dictionary.Keys
            Dim CurTTRString As String: CurTTRString = "_#" & TTRCounter & "_"
            Dim StartPos As Integer, NextPos As Integer
            StartPos = 1
            NextPos = InStr(StartPos, FTTWithAdj, TCITerm)
            Dim FoundMatch As Boolean: FoundMatch = False
            
            While NextPos <> 0
                FoundMatch = True
                StartPos = NextPos
                
                NumCharRight = Len(FTTClean) - Len(TCITerm) - NextPos + 1
                FTTClean = Left(FTTClean, NextPos - 1) & CurTTRString & Right(FTTClean, NumCharRight)
                FTTWithAdj = Left(FTTWithAdj, NextPos - 1) & String(Len(CurTTRString), 25) & Right(FTTWithAdj, NumCharRight)
                'Debug.Print FTTClean & ", with adj = " & FTTWithAdj
                
                NextPos = InStr(StartPos, FTTWithAdj, TCITerm)
            Wend
            'FTTWithAdj2 = Replace(FTTWithAdj, TCITerm, "_#" & TTRCounter & "_")
            
            If FoundMatch Then
                TCITermIndex = m_Dictionary(TCITerm)
                TermsToReplace.Add TCITermIndex
                TTRCounter = TTRCounter + 1
                'Debug.Print (FTTWithAdj)
            End If
        Next TCITerm
        'At this point, all exact-terms translations are done and protected by _#"MatchID"_ markers
        'so it would be a good time to run the regexes and keep their results inside the FTTWithAdj string as they are.
        
        Dim RETTRCounter As Integer
        RETTRCounter = 1
        For Each ExpToReplace In m_RegexpTermsDict
            m_regex.Pattern = ExpToReplace
            Set Matches = m_regex.Execute(FTTWithAdj)
            Dim RETTRString As String: RETTRString = "_$" & RETTRCounter & "_"
            Dim LenDiff As Integer
            LenDiff = 0
            For Each Match In Matches
                REReplacementText.Add m_regex.Replace(Match, m_TransChartOut.Cells(m_RegexpTermsDict(ExpToReplace), 1).Text)
                'en prenant pour acquis qu'une case de TCO de _REGEXP_ ne contiendra pas plusieurs termes distincts.
                'FTTWithAdj = m_regex.Replace(FTTWithAdj, "_RETTR" & RETTRCounter & "_")
                Debug.Print Match.FirstIndex
                
                NumCharRight = Len(FTTWithAdj) - Len(Match) - (Match.FirstIndex + LenDiff)
                NumCharLeft = Match.FirstIndex + LenDiff
                FTTWithAdj = Left(FTTWithAdj, NumCharLeft) & String(Len(RETTRString), 25) & Right(FTTWithAdj, NumCharRight)
                FTTClean = Left(FTTClean, NumCharLeft) & RETTRString & Right(FTTClean, NumCharRight)
                RETTRCounter = RETTRCounter + 1
                
                LenDiff = LenDiff + Len(RETTRString) - Len(Match)
                
            Next Match
        Next ExpToReplace
        
        FTTResult = FTTClean
        TTRCounter = 1
        For Each ReplaceByIndex In TermsToReplace
            
            ReplaceBy = m_GetFirstTCOTerm(CInt(ReplaceByIndex))
            FTTResult = Replace(FTTResult, "_#" & TTRCounter & "_", ReplaceBy)
            TTRCounter = TTRCounter + 1
            
        Next ReplaceByIndex
        
        RETTRCounter = 1
        For Each RepText In REReplacementText
            'm_regex.Pattern = "_$" & RETTRCounter & "_"
            FTTResult = Replace(FTTResult, "_$" & RETTRCounter & "_", RepText)
            RETTRCounter = RETTRCounter + 1
            'Debug.Print (FTTClean)
        Next RepText
        
    End If
    
    m_TranslateOneString = FTTResult

End Function

Public Sub m_ConvertTextToMCRForVariantArray(AddrArgs() As Variant)
    'This basically call m_ConvertTextToMultiChoiceRange on a whole column.
    'Instead of taking in Range objects as arguments, this function takes a single array argument because
    'of the CallByName function used in clsDFSystem in conjunction with it.
    'The arrray is expected to contain addresses in the format "SheetName!A1:A100" :
    'AddrArgs(0) : the address of the column containing the text data to be converted into a multichoice range
    'AddrArgs(1) : the address of the first column in the multichoice range.
    'InputColAddr As String, FirstMCColAddr As String
    
    Dim InputCol As Range, FirstMCCol As Range
    Set InputCol = Range(AddrArgs(0)): Set FirstMCCol = Range(AddrArgs(1))
    
    Call m_ConvertTextToMCRForColumn(InputCol, FirstMCCol)
End Sub

Public Sub m_ConvertTextToMCRForColumn(InputCol As Range, FirstMCCol As Range)

    For i = 1 To InputCol.Rows.count
       Call m_ConvertTextToMultiChoiceRange(InputCol.Cells(i, 1).Text, FirstMCCol.Cells(i, 1))
    Next i

End Sub

    'Here, the InputStr is searched for any term in the m_Dictionary (in descending order of length).
    'In case of a match, the term is removed from the InputStr
    '(to avoid matching any terms with fewer letters that would come later in m_Dictionary)
    'and the term's row number is obtained; we then use m_GetFirstTCOTerm to
    'get the first TransChartOut term present at this row number, and assume it is
    'an integer; this integer is the index of the multichoice cell where we
    'inscribe the value "1".

Public Sub m_ConvertTextToMultiChoiceRange(InputStr As String, FirstCellInMultiChoices As Range)

    If m_UseSpecialChars = False Then
        InputStr = StripAccent(InputStr)
    End If
    
    If m_CaseSensitive = False Then
        InputStr = LCase(InputStr)
    End If

    Dim PrevLen As Integer, CurLen As Integer
    PrevLen = Len(InputStr)
    For Each TCITerm In m_Dictionary.Keys 'sorted in descending order of length, so we're safe here
        InputStr = Replace(InputStr, TCITerm, ""): CurLen = Len(InputStr)
        If CurLen < PrevLen Then
            Dim TargetCol As Integer: TargetCol = Val(m_GetFirstTCOTerm(m_Dictionary(TCITerm)))
            If TargetCol > 0 Then
                FirstCellInMultiChoices.Cells(1, TargetCol).Value = 1
            End If
        PrevLen = CurLen
        End If
    Next TCITerm
    
    
End Sub

