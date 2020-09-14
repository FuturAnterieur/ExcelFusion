Attribute VB_Name = "UtilityFunctions"


Function SheetExists(SheetName As String, Optional wb As Excel.Workbook) As Boolean
   Dim s As Excel.Worksheet
   If wb Is Nothing Then Set wb = ThisWorkbook
   On Error Resume Next
   Set s = wb.Worksheets(SheetName)
   On Error GoTo 0
   SheetExists = Not s Is Nothing
End Function

Public Function IsInArray(stringToBeFound As String, Arr As Variant) As Boolean
  IsInArray = (UBound(Filter(Arr, stringToBeFound)) > -1)
End Function

Public Function dhAge(dtmBD As Date, Optional dtmDate As Date = 0) _
 As Integer
    ' This procedure is stored as dhAgeUnused in the sample
    ' module.
    Dim intAge As Integer
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    intAge = DateDiff("yyyy", dtmBD, dtmDate)
    If dtmDate < DateSerial(Year(dtmDate), Month(dtmBD), Day(dtmBD)) Then
        intAge = intAge - 1
    End If
    dhAge = intAge
End Function


'pris sur https://www.extendoffice.com/documents/excel/707-excel-replace-accented-characters.html

Function StripAccent(thestring As String)
Dim A As String * 1
Dim B As String * 1
Dim i As Integer
Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
For i = 1 To Len(AccChars)
A = Mid(AccChars, i, 1)
B = Mid(RegChars, i, 1)
thestring = Replace(thestring, A, B)
Next
StripAccent = thestring
End Function


Sub StripAccentOnRange(TheRange As Range)

    Dim A As String * 1
    Dim B As String * 1
    Dim i As Integer
    Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
    For i = 1 To Len(AccChars)
        
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        
        TheRange.Replace What:=A, Replacement:=B, LookAt:=xlPart, MatchCase:=True
    Next i
End Sub

Public Function EscapeRegExp(toescstr As String, AdaptWildCards As Boolean)

Dim CharsToEscape() As Variant

If AdaptWildCards = True Then
    CharsToEscape = Array("\", "(", ")", "+", "[", "]", "?", ".", "$", "?", "|")
Else
    CharsToEscape = Array("\", "(", ")", "+", "[", "]", "?", ".", "$", "?", "*", "|")
End If

'should tell the doctors not to use backslashes. handling them would take more effort.

For Each cte In CharsToEscape
    toescstr = Replace(toescstr, cte, "\" & cte)
Next cte

If AdaptWildCards = True Then
    toescstr = Replace(toescstr, "*", ".+")
End If

EscapeRegExp = toescstr

End Function

Public Function FixOneDate(InputVal As Variant, OGCellFormat As Variant) As Variant
    
    Dim Result As Variant
    Dim DateResult As Date
    
    Dim CaseNum As Integer
    CaseNum = 0
    If InputVal = "" Then
        Result = ""
    Else
        If InputVal Like "##/##/####" Or InputVal Like "##-##-####" Then
            Dim YearVal As Integer, MonthVal As Integer, DayVal As Integer
            YearVal = Right(InputVal, 4)
            MonthVal = Mid(InputVal, 4, 2)
            DayVal = Left(InputVal, 2)
            Dim CheckValid As Boolean
            CheckValid = YearVal > 1900 And MonthVal > 0 And MonthVal < 13 And DayVal > 0 And DayVal < 32
            'CheckValidHarder should also check for NumDaysPerMonth and February 29th shenanigans
            If CheckValid Then
                CaseNum = 3
                Result = DateSerial(YearVal, MonthVal, DayVal)
            Else
                CaseNum = -1
                Result = "Cannot convert date (invalid day or month, or year inferior to 1900)"
            End If
        ElseIf IsDate(InputVal) Then
            CaseNum = 2
            Result = InputVal
        'TODO : use regexes to capture other less-common patterns
        
        Else 'something, but that was not recognized as a date
            CaseNum = 1
            Result = InputVal
        End If
    End If

    If CaseNum > 1 Then
        DateResult = CDate(Result)
        DateResult = Format(DateResult, "yyyy-MM-dd;@")
        FixOneDate = DateResult
    Else
        FixOneDate = Result
    End If
    
End Function

Public Sub FixAllDatesInRange(InputCells As Range, Optional OutputCells As Range)
    
    If OutputCells Is Nothing Then
        Set OutputCells = InputCells
    End If
    
    Dim IC As Range
    Dim OC As Range

    For Col = 1 To InputCells.Columns.count
        For Row = 1 To InputCells.Rows.count
            Set IC = InputCells.Cells(Row, Col)
            Set OC = OutputCells.Cells(Row, Col)
                OC.Value = FixOneDate(IC.Value, IC.NumberFormat)
                OC.NumberFormat = "yyyy-MM-dd;@" 'this doesn't seem to be necessary, at least as far as my needs are concerned.
                'well maybe it does change smth with the @. I should check this out further.
        Next Row
    Next Col
    OutputCells.TextToColumns Destination:=OutputCells, DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)


'    ActualDateCell.NumberFormat = "yyyy-MM-dd;@"
'    ActualDateCell.Select
'    Selection.TextToColumns Destination:=ActualDateCell, DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)
'    ActualDateCell.NumberFormat = "yyyy-MM-dd;@"
'    this actually seems to be unnecessary
'    so the only scenario in which a whole-range method (and not one value at a time) would be useful would be if
'   the dates we are correcting are ambiguous, i.e. 06-05-1999, or even 05-03-02.
'   in a case like this, we would have to check other dates in the same range until we find one that isn't ambiguous
'   i.e. 12-13-1999 (or 05-20-00) and (hopefully) in the same format as the others around it.
'ALSO : Take heed that bulk operations (like calling TextToColumns on the whole column at once) would be faster
'than looping through each cell, but requires that dates are all uniformly formatted beforehand -- whereas
'the functions I am making are aimed at handling cases where many date formats appear in the same column


End Sub
Public Sub FixDuration(InputDurations As Variant)

    

End Sub

Public Sub FindAllTransChartsOnSheet(ChartsSheetName As String, ParentObject As Object, Optional SheetExistenceChecked As Boolean = False)

    If Not SheetExistenceChecked Then
        If Not SheetExists(ChartsSheetName) Then
            MsgBox ("Specified trans chart sheet name " & ChartsSheetName & " does not exist in this workbook. Aborting.")
            Err.Raise 1998
        End If
    End If
    
    Set ChartsSheet = Worksheets(ChartsSheetName)
    
    Dim CurTCCell As Range, CurTCOptionsCell As Range
    Dim NextTCAddr As String, FirstTCFoundAddr As String, FirstTCFound As Range
    Set FirstTCFound = ChartsSheet.Cells.Find(What:="TC_*_", after:=Cells(1, 1), LookAt:=xlWhole, MatchCase:=False, SearchDirection:=xlNext)
    If Not FirstTCFound Is Nothing Then
        FirstTCFoundAddr = FirstTCFound.Address
        
        NextTCAddr = FirstTCFoundAddr
        Do
            Set CurTCCell = Range(NextTCAddr)
            
            ParentObject.m_TranslationCharts.Add Key:=CurTCCell.Text, Item:=New clsTransChart
            Call ParentObject.m_TranslationCharts(CurTCCell.Text).m_BuildDictFromDescCell(CurTCCell)
           
            NextTCAddr = ChartsSheet.Cells.FindNext(after:=CurTCCell).Address
            
        Loop While NextTCAddr <> FirstTCFoundAddr
    End If
End Sub



Sub ExtractTermsMk3(IFaddr As String, OCaddr As String, Separators() As String, KeepUnsepVersions As Boolean, CaseSensitive As Boolean, UseSpecialChars As Boolean)

Dim InputField As Range
Set InputField = Range(IFaddr)
Dim OutputCell As Range
Set OutputCell = Range(OCaddr)
'Dim Separators() As Variant
'Dim KeepUnsepVersions As Boolean, CaseSensitive As Boolean, UseSpecialChars As Boolean

'All Separators are Replaced by "|" before splitting by "|" - seems like a legit trick


Dim AllEntriesArray() As Variant
AllEntriesArray = Array("|_\\")

Dim EquivDict As Scripting.Dictionary
Set EquivDict = New Scripting.Dictionary

EquivDict.CompareMode = 0

Dim IFCell As Range
Dim OldBound As Integer
Dim IFCellText As String
Dim IFCellTerms() As String
Dim CurTermAdj As String

Dim MasterSep As String: MasterSep = "_or_"
Dim LenMS As Integer: LenMS = Len(MasterSep)
Dim MSCompMethod As VbCompareMethod
MSCompMethod = vbBinaryCompare
If LenMS > 1 Then
    MSCompMethod = vbTextCompare
End If

Dim CheckIfWord As Object
Set CheckIfWord = CreateObject("VBScript.RegExp")
CheckIfWord.Global = True
CheckIfWord.MultiLine = False
CheckIfWord.IgnoreCase = Not CaseSensitive
CheckIfWord.Pattern = "^\w+$"

Dim ReplaceWordSep As Object
Set ReplaceWordSep = CreateObject("VBScript.RegExp")
ReplaceWordSep.Global = True
ReplaceWordSep.MultiLine = False
ReplaceWordSep.IgnoreCase = Not CaseSensitive

Dim RemoveSpaces As Object
Set RemoveSpaces = CreateObject("VBScript.RegExp")
RemoveSpaces.Global = True
RemoveSpaces.MultiLine = False
RemoveSpaces.IgnoreCase = Not CaseSensitive
RemoveSpaces.Pattern = "\s*" & EscapeRegExp(MasterSep, False) & "\s*"

Dim AdjOn As Boolean
AdjOn = Not UseSpecialChars Or Not CaseSensitive


For Each IFCell In InputField
    
    IFCellText = IFCell.Text

    For Each Separator In Separators
        'what if separator words like "et" are contained in content words, i.e. "Etagère et Lampadaire"?)
        If CheckIfWord.Test(Separator) = True Then
            'if it is a word, only consider it as a separator when it exists as a separate word (i.e. with word boundaries around it)
            ReplaceWordSep.Pattern = "\b" & Separator & "\b"
            IFCellText = ReplaceWordSep.Replace(IFCellText, MasterSep)
        Else
            'if the separator is any combination containing non-word characters
            IFCellText = Replace(IFCellText, Separator, MasterSep)
        End If
    Next Separator
    
    IFCellText = RemoveSpaces.Replace(IFCellText, MasterSep)
    IFCellText = Trim(IFCellText)
    
   
    If Not IFCellText = "" Then
        
        IFCellTerms = Split(IFCellText, MasterSep, , MSCompMethod)
    Else
        If Not IsInArray("", AllEntriesArray) Then
            ReDim Preserve AllEntriesArray(OldBound + 1)
            AllEntriesArray(OldBound + 1) = ""
        End If
            GoTo NextIteration
        
    End If
    
    
    
    If Len(IFCellTerms(0)) < Len(IFCellText) And KeepUnsepVersions = True Then
        ReDim Preserve IFCellTerms(UBound(IFCellTerms) + 1)
        IFCellTerms(UBound(IFCellTerms)) = IFCell.Text
        'could send a half-processed version of IFCell.Text, but hey, no! We're keeping unsep
        'versions specifically for translating them directly, as they are.
    End If
        
    'here, modifications were made for very special cases.
    '"pâtes" and "pates" will get converted to...
    'If UseSpecialChars : two entries, just as expected
    'If Not UseSpecialChars : One entry : "pâtes", and not "pates".
    
    Dim i As Integer, OldBoundA As Integer
    Dim AdjAndOrigSame As Boolean, DictAdjAndOrigSame As Boolean, AdjInDict As Boolean, OrigInArray As Boolean
    Dim AddToAEA As Boolean
    
    
    If Not AdjOn Then
        'terms will never have to be modified before being added to our array
        For i = 0 To UBound(IFCellTerms)
            OldBound = UBound(AllEntriesArray)
            If Not IsInArray(IFCellTerms(i), AllEntriesArray) Then
                ReDim Preserve AllEntriesArray(OldBound + 1)
                AllEntriesArray(OldBound + 1) = IFCellTerms(i)
            End If
        Next i
    Else
    
        For i = 0 To UBound(IFCellTerms)
        
            CurTermAdj = IFCellTerms(i)
            If Not CaseSensitive Then
                CurTermAdj = LCase(CurTermAdj)
            End If
    
            If Not UseSpecialChars Then
                CurTermAdj = StripAccent(CurTermAdj)
            End If
            
            AdjAndOrigSame = (CurTermAdj = IFCellTerms(i))

            
            AdjInDict = EquivDict.Exists(CurTermAdj)
            If Not AdjInDict Then
                EquivDict.Add Key:=CurTermAdj, Item:=IFCellTerms(i)
            Else
                DictAdjAndOrigSame = (CurTermAdj = EquivDict(CurTermAdj))
                If DictAdjAndOrigSame And Not AdjAndOrigSame Then
                    EquivDict(CurTermAdj) = IFCellTerms(i)
                End If
            End If
        Next i
        
    End If

NextIteration:
Next IFCell


If AdjOn Then
Dim AdjTermKey As Variant
For Each AdjTermKey In EquivDict.Keys
    OldBoundA = UBound(AllEntriesArray)
    ReDim Preserve AllEntriesArray(OldBoundA + 1)
    AllEntriesArray(OldBoundA + 1) = EquivDict.Item(AdjTermKey)
Next AdjTermKey
End If


Dim EntryCell As Range
For i = 1 To UBound(AllEntriesArray)
    Set EntryCell = OutputCell.Cells(i, 1)
    EntryCell.Value = AllEntriesArray(i)
Next i

OutputCell.Worksheet.Select

End Sub

Sub CountOccurrences(TermsRng As Range, FieldToCount As Range, Optional OutputRng As Range)
    
    If OutputRng Is Nothing Then
        Set OutputRng = TermsRng.Offset(0, 1)
    End If
    Dim FTCSheetName As String: FTCSheetName = FieldToCount.Worksheet.Name
    
    Dim TermCell As Range
    For i = 1 To TermsRng.Rows.count
        Set TermCell = TermsRng.Cells(i, 1)
        'Debug.Print (TermCell.Address)
        Dim TermCount As Integer: TermCount = 0
        Dim FirstFound As Range, CurFound As Range
        Dim FirstAddr As String, NextAddr As String
        Set FirstFound = FieldToCount.Find(What:=TermCell.Text, LookAt:=xlPart, MatchCase:=False, SearchDirection:=xlNext)
        If Not FirstFound Is Nothing Then
            FirstAddr = FirstFound.Address
            NextAddr = FirstAddr
            Do
                TermCount = TermCount + 1
                Set CurFound = Worksheets(FTCSheetName).Range(NextAddr)
                NextAddr = FieldToCount.FindNext(after:=CurFound).Address
            Loop While NextAddr <> FirstAddr
        End If
        OutputRng.Cells(i, 1).Value = TermCount
    Next i
End Sub


