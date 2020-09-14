Attribute VB_Name = "TooComplicatedTransChartCode"

'this module contains old versions of TransChart that could manage versions of the Terms dictionary that had appended column numbers
'(which could be useful in cases where similar input terms had to be differentiated by the column they were on)-
'and could only be used effectively if the code that called the TransChart also managed these column numbers.
'
'Public m_WithColNum As Boolean
'Private m_ColNumFinder As Object
'Public m_ColNumlessDict As Scripting.Dictionary
'
'Public Sub Class_Initialize()
'
'    m_WithColNum = False
'
'    Set m_ColNumFinder = CreateObject("VBScript.Regexp")
'    m_ColNumFinder.Global = True
'    m_ColNumFinder.MultiLine = False
'    m_ColNumFinder.IgnoreCase = True
'    m_ColNumFinder.Pattern = "_from_[0-9]+"
'
'    Set m_ColNumlessDict = New Scripting.Dictionary
'
'End Sub
'
'
'Public Sub m_BuildDictionary(TCI As Range, TCO As Range, Optional CS As Boolean = False, Optional USC As Boolean = False, _
'Optional UsesWildCards As Boolean = False, Optional Sep As String = "_or_", Optional AppendColNum As Boolean = False)
'
'
'Debug.Print ("In BuildDictionary")
'
'Set m_TransChartIn = TCI
'Set m_TransChartOut = TCO
'm_TransChartIn.Parent.Select
'
'Dim UnsortedDict As Scripting.Dictionary
'Set UnsortedDict = New Scripting.Dictionary
'
'm_regex.IgnoreCase = Not CS
'm_UseSpecialChars = USC
'm_CaseSensitive = CS
'm_UsesWildCards = UsesWildCards
'm_WithColNum = AppendColNum
'
'm_TermSeparator = Sep
'
'm_SplitCompMethod = vbBinaryCompare
'If Len(m_TermSeparator) > 1 Then
'        m_SplitCompMethod = vbTextCompare 'vbTextcompare is not case-sensitive, but works with multi-char separators
'End If
'
'Dim regexTS As String
'regexTS = EscapeRegExp(m_TermSeparator, False)
'Debug.Print ("regexTS is " & regexTS)
'
'Dim CurTerms As String
'Dim CurTermsArray() As String
'
'Dim TCIRow As Integer, TCICol As Integer
'Dim TCICell As Range
'Dim ColNumText As String
'Dim FirstRowToAdd As Integer
'
'For TCICol = 1 To m_TransChartIn.Columns.count
'    If AppendColNum Then
'        ColNumText = "_from_" & TCICol
'    Else
'        ColNumText = ""
'    End If
'
'    For TCIRow = 1 To m_TransChartIn.Rows.count
'
'        Set TCICell = m_TransChartIn.Cells(TCIRow, TCICol)
'        'Debug.Print (TCICell.Text)
'
'        CurTerms = TCICell.Text
'        If m_CaseSensitive = False Then
'            CurTerms = LCase(CurTerms)
'        End If
'
'        If m_UseSpecialChars = False Then
'            CurTerms = StripAccent(CurTerms)
'        End If
'
'        CurTermsArray = Split(CurTerms, regexTS, , SplitCompMethod)
'
'        Dim i As Integer
'        For i = 0 To UBound(CurTermsArray)
'            If InStr(CurTermsArray(i), "_notrim_") = 0 Then
'                CurTermsArray(i) = Trim(CurTermsArray(i))
'            Else
'                m_regex.Pattern = "\s*_notrim_" '_notrim_ and any spaces BEFORE it will be removed. All other spaces will be left untrimmed.
'                CurTermsArray(i) = m_regex.Replace(CurTermsArray(i), "")
'            End If
'
'            If InStr(CurTermsArray(i), "_regexp_") = 0 Then
'                CurTermsArray(i) = EscapeRegExp(CurTermsArray(i), m_UsesWildCards)
'                 If Not UnsortedDict.Exists(CurTermsArray(i) & ColNumText) Then
'                    UnsortedDict.Add Key:=(CurTermsArray(i) & ColNumText), Item:=TCIRow
'                End If
'
'            Else
'                CurTermsArray(i) = Replace(CurTermsArray(i), "_regexp_", "")
'                'partial implementation of dictionary goes here
'                Dim AddrPos As Integer, EndAddr As Integer
'                AddrPos = InStr(CurTermsArray(i), "_range_")
'                If AddrPos > 0 Then
'                    AddrPos = AddrPos + Len("_range_")
'                    EndAddr = InStr(AddrPos, CurTermsArray(i), "_")
'                    Dim AddrString As String
'                    AddrString = Mid(CurTermsArray(i), AddrPos, EndAddr - AddrPos)
'                    Dim DicoInputRng As Range
'                    Set DicoInputRng = Range(AddrString)
'                    Dim DicoInputTerms As String, CurDicoTerms As String
'                    DicoInputTerms = ""
'                    For Each DICell In DicoInputRng
'                        DicoInputTerms = DicoInputTerms & DICell.Text & "|"
'                    Next DICell
'                    DicoInputTerms = Replace(DicoInputTerms, "| ", "|")
'                    DicoInputTerms = Left(DicoInputTerms, Len(DicoInputTerms) - 1)
'                    'sorting the terms in order of length would be useless,
'                    'since using \b (word boundary) markers in the regex automatically takes
'                    'care of the issue.
'                    CurTermsArray(i) = Replace(CurTermsArray(i), "_range_" & AddrString & "_", DicoInputTerms)
'                End If
'
'                If Not m_RegexpTermsDict.Exists(CurTermsArray(i) & ColNumText) Then
'                    m_RegexpTermsDict.Add Key:=(CurTermsArray(i) & ColNumText), Item:=TCIRow
'                End If
'            End If
'        Next i
'    Next TCIRow
'Next TCICol
'
''Next up : Sorting this list.
''Could have begun sorting while I was going through the m_TransChartIn range.
''But I'll settle for cheapo sorting for now, and see what it yields.
'
'Dim x As Integer, y As Integer
'Dim TempTxt1 As String, TempTxt2 As String
'Dim KeySorterArr() As String
'
'ReDim KeySorterArr(UnsortedDict.count)
'
'Dim it As Integer: it = 0
'For Each curKey In UnsortedDict.Keys
'    KeySorterArr(it) = curKey
'    it = it + 1
'Next curKey
'
'For x = 0 To UnsortedDict.count - 2
'    For y = x + 1 To UnsortedDict.count - 1
'        If Len(KeySorterArr(y)) > Len(KeySorterArr(x)) Then
'            TempTxt1 = KeySorterArr(x)
'            TempTxt2 = KeySorterArr(y)
'            KeySorterArr(x) = TempTxt2
'            KeySorterArr(y) = TempTxt1
'        End If
'    Next y
'Next x
'
''to do : sort both the ColNumfull and ColNumless dictionaries by length of the ColNumless version of each term
''for now, they are sorted by ColNumfull lenght, which makes less sense...
''well anyway, if the user is using ColNums, then his Terms to be Translated also
''have to be ColNumed, so the point is kinda moot too.
'
'For j = 0 To UnsortedDict.count - 1
'    m_Dictionary.Add Key:=KeySorterArr(j), Item:=UnsortedDict(KeySorterArr(j))
'    Debug.Print ("New Entry  : " & KeySorterArr(j) & "-" & UnsortedDict(KeySorterArr(j)))
'    If AppendColNum Then 'that means there is a point in compiling a ColNumless dictionary
'        Dim ColNumlessTerm As String
'        ColNumlessTerm = m_ColNumFinder.Replace(KeySorterArr(j), "")
'        If Not m_ColNumlessDict.Exists(ColNumlessTerm) Then
'            m_ColNumlessDict.Add Key:=ColNumlessTerm, Item:=UnsortedDict(KeySorterArr(j))
'        '    Debug.Print ("New ColNumless entry : " & ColNumlessTerm)
'        End If
'    End If
'Next j
'
'
'End Sub
'
'Public Sub m_TranslateCells(InputCells As Range, OutputCells As Range, Optional Mode As Integer = 0, Optional RegexpMode As Integer = 1, _
'Optional RemoveColNumFromInput As Boolean = False)
'
''RegexpMode only effective with Mode = 0;
''RemoveColNumFromInput only effective with Mode = 1.
'
'Dim OutputStr As String, InputStr As String
'Dim InRow As Integer, InCol As Integer, OutRow As Integer, OutCol As Integer
'
'    If InputCells.Rows.count = OutputCells.Rows.count And InputCells.Columns.count = OutputCells.Columns.count Then
'
'        For InRow = 1 To InputCells.Rows.count
'            For InCol = 1 To InputCells.Columns.count
'                InputStr = InputCells.Cells(InRow, InCol)
'                If Mode = 0 Then 'text-wise translation
'                    OutputStr = m_TranslateOneString(InputStr, RegexpMode)
'                Else 'Full-Cell translation, by direct m_Dictionary/m_HeaderlessDictionary search
'                    OutputStr = m_RetrieveDicoTerm(InputStr, RemoveColNumFromInput)
'                End If
'                OutputCells.Cells(InRow, InCol).Value = OutputStr
'            Next InCol
'        Next InRow
'    Else
'        Dim ChangedCell As Boolean
'        For OutCol = 1 To OutputCells.Columns.count
'
'            InCol = OutCol
'            If InCol > InputCells.Columns.count Then
'                InCol = InputCells.Columns.count
'            End If
'
'            For OutRow = 1 To OutputCells.Rows.count
'                ChangedCell = True
'                InRow = OutRow
'                If OutRow > InputCells.Rows.count Then
'                    InRow = InputCells.Rows.count
'                    ChangedCell = False
'                End If
'
'                If ChangedCell = True Then ' refresh the output string
'                    InputStr = InputCells.Cells(InRow, InCol).Text
'                    If Mode = 0 Then 'text-wise translation
'                        OutputStr = m_TranslateOneString(InputStr, RegexpMode)
'                    Else 'Full-Cell translation
'                        OutputStr = m_RetrieveDicoTerm(InputStr, RemoveColNumFromInput)
'                    End If
'                End If
'
'                OutputCells.Cells(OutRow, OutCol).Value = OutputStr
'            Next OutRow
'        Next OutCol
'
'    End If
''TODO : Add Function that does a Range-Replace for a whole range at once, both for Mode0-without-regexes and Mode1
''(where DicoSearch would be replaced by "Match whole cell contents")
'
'End Sub
'
'
'Public Function m_RetrieveDicoTerm(InputVal As Variant, Optional RemoveColNumFromInput As Boolean = False) As Variant
'
'    Dim ResRowNum As Integer: ResRowNum = m_RetrieveTCIRow(InputVal, RemoveColNumFromInput)
'
'    m_RetrieveDicoTerm = m_TransChartOut.Cells(ResRowNum, 1).Value
'End Function
'
'Public Function m_RetrieveTCIRow(InputVal As Variant, Optional RemoveColNumFromInput As Boolean = False) As Integer
'    'minor todo : add a feature to take trimmed/untrimmed terms into account
'    'kinda harder, because trim/untrim is not uniform for the whole TransChart, but specified term-wise.
'    'also, the trim/untrimmed term feature is quite unlikely to be used in dictionary (cell by cell) mode.
'
'    Dim InputStr As String
'    InputStr = CStr(InputVal)
'
'    Dim ResRowNum As Integer
'    'InputStr = Trim(InputStr)
'    If m_UseSpecialChars = False Then
'        InputStr = StripAccent(InputStr)
'    End If
'
'    If m_CaseSensitive = False Then
'        InputStr = LCase(InputStr)
'    End If
'    InputStr = EscapeRegExp(InputStr, m_UsesWildCards)
'
'    If Not RemoveColNumFromInput Or m_ColNumlessDict.count = 0 Then
'        If m_Dictionary.Exists(InputStr) Then
'            ResRowNum = m_Dictionary(InputStr)
'        Else
'            ResRowNum = -1
'        End If
'    Else
'        Set ColNumTexts = m_ColNumFinder.Execute(InputStr)
'        Dim ColNumlessInput As String
'        ColNumlessInput = m_ColNumFinder.Replace(InputStr, "")
'
'        If (m_ColNumlessDict.Exists(ColNumlessInput)) Then
'            ResRowNum = m_ColNumlessDict(ColNumlessInput)
'        Else
'            ResRowNum = -1
'        End If
'    End If
'
'    m_RetrieveTCIRow = ResRowNum
'
'End Function
'
'
'Public Function m_TranslateOneString_NormalTermsFirst(InputVal As Variant) As String
'
'    Dim InputStr As String
'    InputStr = CStr(InputVal)
'
'    Dim TermsToReplace As New Collection
'    Set TermsToReplace = Nothing
'
'    Dim REReplacementText As New Collection
'    Set REReplacementText = Nothing
'
'    Dim FTTInterm As String
'    FTTInterm = InputStr
'
'    If InputStr = "" Then
'        GoTo QuickEnd
'    End If
'
'    If m_UseSpecialChars = False Then
'        FTTInterm = StripAccent(FTTInterm)
'    End If
'
'    Dim TTRCounter As Integer
'    Dim TCITerm As Variant
'    Dim TCITermIndex As Integer
'
'
'    Dim FTTInterm2 As String
'    TTRCounter = 1
'    For Each TCITerm In m_Dictionary.Keys 'i'll have to handle regexes afterwards
'
'        m_regex.Pattern = TCITerm
'
'        FTTInterm2 = m_regex.Replace(FTTInterm, "_TTR" & TTRCounter & "_")
'        If FTTInterm2 <> FTTInterm Then
'            TCITermIndex = m_Dictionary(TCITerm)
'
'            TermsToReplace.Add TCITermIndex
'            TTRCounter = TTRCounter + 1
'            FTTInterm = FTTInterm2
'            'Debug.Print (FTTInterm)
'        End If
'    Next TCITerm
'    'At this point, all exact-terms translations are done and protected by _TTR"MatchID"_ markers
'    'so it would be a good time to run the regexes and keep their results inside the FTTinterm string as they are.
'
'    Dim RETTRCounter As Integer
'
'    RETTRCounter = 1
'    For Each ExpToReplace In m_RegexpTermsDict
'        m_regex.Pattern = ExpToReplace
'        Set Matches = m_regex.Execute(FTTInterm)
'
'        For Each Match In Matches
'            REReplacementText.Add m_regex.Replace(Match, m_TransChartOut.Cells(m_RegexpTermsDict(ExpToReplace), 1).Text)
'            'en prenant pour acquis qu'une case de TCO de _REGEXP_ ne contiendra pas plusieurs termes distincts.
'            'c'est vrai que je suis paresseux des fois
'            FTTInterm = m_regex.Replace(FTTInterm, "_RETTR" & RETTRCounter & "_")
'            RETTRCounter = RETTRCounter + 1
'            'Debug.Print (FTTInterm)
'        Next Match
'    Next ExpToReplace
'
'    Dim FTTResult As String: FTTResult = FTTInterm
'
'    TTRCounter = 1
'    For Each ReplaceByIndex In TermsToReplace
'        Dim ReplaceBy As String
'        Dim AllTCOText As String
'        Dim TCOTerms() As String
'
'        AllTCOText = m_TransChartOut.Cells(ReplaceByIndex, 1).Text
'
'        If Not AllTCOText = "" Then
'            TCOTerms = Split(AllTCOText, m_TermSeparator, , m_SplitCompMethod)
'            ReplaceBy = TCOTerms(0)
'        Else
'            ReplaceBy = ""
'        End If
'        'Debug.Print (ToReplace & " - " & ReplaceBy)
'
'        m_regex.Pattern = "_TTR" & TTRCounter & "_"
'        FTTResult = m_regex.Replace(FTTResult, ReplaceBy)
'
'        TTRCounter = TTRCounter + 1
'    Next ReplaceByIndex
'
'
'    RETTRCounter = 1
'    For Each RepText In REReplacementText
'        m_regex.Pattern = "_RETTR" & RETTRCounter & "_"
'        FTTResult = m_regex.Replace(FTTResult, RepText)
'
'        RETTRCounter = RETTRCounter + 1
'    Next RepText
'
'
'QuickEnd:
'    m_TranslateOneString_NormalTermsFirst = FTTResult
'
'
'End Function
Public Function m_TranslateOneString(InputVal As Variant)
    If InputVal <> "" Then

        If RegexpMode = 1 Then
            m_TranslateOneString = m_TranslateOneString_RegExpComboThenNormal(InputVal)
        Else
            m_TranslateOneString = m_TranslateOneString_NormalTermsFirst(InputVal)
        End If
    Else
        m_TranslateOneString = ""
    End If

End Function
    


Public Function m_TranslateOneString_NormalTermsFirst(InputVal As Variant) As String

    Dim InputStr As String
    InputStr = CStr(InputVal)

    Dim TermsToReplace As New Collection
    Set TermsToReplace = Nothing

    Dim REReplacementText As New Collection
    Set REReplacementText = Nothing

    Dim FTTInterm As String
    FTTInterm = InputStr
    
    If m_CaseSensitive = False Then
        FTTInterm = LCase(FTTInterm)
    End If

    If m_UseSpecialChars = False Then
        FTTInterm = StripAccent(FTTInterm)
    End If
    
    Dim FTTResult As String: FTTResult = FTTInput
    
    If InputStr <> "" Then
        
        Dim TTRCounter As Integer
        Dim TCITerm As Variant
        Dim TCITermIndex As Integer
        
        Dim FTTInterm2 As String
        TTRCounter = 1
        For Each TCITerm In m_Dictionary.Keys 'i'll have to handle regexes afterwards
    
            FTTInterm2 = Replace(FTTInterm, TCITerm, "_TTR" & TTRCounter & "_")
            If FTTInterm2 <> FTTInterm Then
                TCITermIndex = m_Dictionary(TCITerm)
    
                TermsToReplace.Add TCITermIndex
                TTRCounter = TTRCounter + 1
                FTTInterm = FTTInterm2
                Debug.Print FTTInterm2
            End If
        Next TCITerm
        'At this point, all exact-terms translations are done and protected by _TTR"MatchID"_ markers
        'so it would be a good time to run the regexes and keep their results inside the FTTinterm string as they are.
    
        Dim RETTRCounter As Integer
    
        RETTRCounter = 1
        For Each ExpToReplace In m_RegexpTermsDict
            m_regex.Pattern = ExpToReplace
            Set Matches = m_regex.Execute(FTTInterm)
    
            For Each Match In Matches
                REReplacementText.Add m_regex.Replace(Match, m_TransChartOut.Cells(m_RegexpTermsDict(ExpToReplace), 1).Text)
                'en prenant pour acquis qu'une case de TCO de _REGEXP_ ne contiendra pas plusieurs termes distincts.
                'c'est vrai que je suis paresseux des fois
                FTTInterm = m_regex.Replace(FTTInterm, "_RETTR" & RETTRCounter & "_")
                RETTRCounter = RETTRCounter + 1
                'Debug.Print (FTTInterm)
            Next Match
        Next ExpToReplace
    
        FTTResult = FTTInterm
    
        TTRCounter = 1
        For Each ReplaceByIndex In TermsToReplace
            Dim ReplaceBy As String
            Dim AllTCOText As String
            Dim TCOTerms() As String
            ReplaceBy = m_GetFirstTCOTerm(CInt(ReplaceByIndex))
            FTTResult = Replace(FTTResult, "_TTR" & TTRCounter & "_", ReplaceBy)
            TTRCounter = TTRCounter + 1
        Next ReplaceByIndex
    
    
        RETTRCounter = 1
        For Each RepText In REReplacementText
            m_regex.Pattern = "_RETTR" & RETTRCounter & "_"
            FTTResult = m_regex.Replace(FTTResult, RepText)
            RETTRCounter = RETTRCounter + 1
        Next RepText

    End If

    m_TranslateOneString_NormalTermsFirst = FTTResult

End Function


Public Function m_TranslateOneString_RegExpComboThenNormal(InputVal As Variant) As String

    Dim InputStr As String
    InputStr = CStr(InputVal)
    'Debug.Print InputStr
    
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
        
        Dim RETTRCounter As Integer
        RETTRCounter = 1
        For Each ExpToReplace In m_RegexpTermsDict
            m_regex.Pattern = ExpToReplace
            Set Matches = m_regex.Execute(FTTWithAdj)
            Dim RETTRString As String: RETTRString = "_$" & RETTRCounter & "_"
            
            For Each Match In Matches
                REReplacementText.Add m_regex.Replace(Match, m_TransChartOut.Cells(m_RegexpTermsDict(ExpToReplace), 1).Text)
                'en prenant pour acquis qu'une case de TCO de _REGEXP_ ne contiendra pas plusieurs termes distincts.
                'FTTWithAdj = m_regex.Replace(FTTWithAdj, "_RETTR" & RETTRCounter & "_")
                
                NumCharRight = Len(FTTWithAdj) - Len(Match) - Match.FirstIndex
                FTTWithAdj = Left(FTTWithAdj, Match.FirstIndex) & String(Len(RETTRString), 25) & Right(FTTWithAdj, NumCharRight)
                FTTClean = Left(FTTClean, Match.FirstIndex) & RETTRString & Right(FTTClean, NumCharRight)
                RETTRCounter = RETTRCounter + 1
                'Debug.Print (FTTClean)
            Next Match
        Next ExpToReplace
        
        RETTRCounter = 1
        For Each RepText In REReplacementText
            'm_regex.Pattern = "_$" & RETTRCounter & "_"
            FTTClean = Replace(FTTClean, "_$" & RETTRCounter & "_", RepText)
            RETTRCounter = RETTRCounter + 1
            'Debug.Print (FTTClean)
        Next RepText
        
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
                Debug.Print FTTClean & ", with adj = " & FTTWithAdj
                
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
        
        
        FTTResult = FTTClean
        TTRCounter = 1
        For Each ReplaceByIndex In TermsToReplace
            
            ReplaceBy = m_GetFirstTCOTerm(CInt(ReplaceByIndex))
            FTTResult = Replace(FTTResult, "_#" & TTRCounter & "_", ReplaceBy)
            TTRCounter = TTRCounter + 1
            
        Next ReplaceByIndex
    End If
    
    m_TranslateOneString_RegExpComboThenNormal = FTTResult

End Function
