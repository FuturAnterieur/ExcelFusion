VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDPSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DP stands for Data Processor.
'It takes only one input sheet and ouputs one output sheet;
'but it does not fuse data between sheets.
'It justs takes various columns from the source sheets
'as input, and processes their data according to instructions.
Public m_SourceSheet As clsDataSheet
Public m_OutputOnOtherSheet As Boolean
Public m_OutputSheetName As String
Public m_LastOutputRow As Integer
Public m_TranslationCharts As Scripting.Dictionary

Public m_DFManager As clsDataFieldsManager

Public Sub Class_Initialize()

    Set m_TranslationCharts = New Scripting.Dictionary
    Set m_DFManager = New clsDataFieldsManager
    m_OutputOnOtherSheet = True

End Sub

Public Function m_GetDFLocation(InputIDF As String) As clsMoreDFLocationInfo

    Dim DFLI As New clsMoreDFLocationInfo

    Dim FoundSuccess As Boolean: FoundSuccess = False
    Dim Provenance As String, FieldName As String
    If InStr(InputIDF, "->") > 0 Then
        Dim idfparts() As String: idfparts = Split(InputIDF, "->", 2, vbTextCompare)
        Provenance = idfparts(0)
        FieldName = idfparts(1)
    Else
        Provenance = ""
        FieldName = InputIDF
    End If
    
    Dim TargetInternalDF As clsDataField
    If m_DFManager.m_DataFieldsChart.Exists(FieldName) Then
        Set TargetInternalDF = m_DFManager.m_DataFieldsChart(FieldName)
        If Provenance = "IDF" Or Provenance = "" Then
            If TargetInternalDF.m_ColOnOutputSheet > 0 Then
            'In this case, the program will look for the value as it appears (corrected) on the output sheet.
                With DFLI
                    .TargetShName = m_OutputSheetName
                    .StartRowNum = 2
                    .EndRowNum = m_LastOutputRow
                    .ColNum = TargetInternalDF.m_ColOnOutputSheet
                End With
                FoundSuccess = True
            Else
                MsgBox "The given IDF name " & FieldName & " does not appear on the output sheet."
            End If
        Else
            Dim ShObj As clsDataSheet
            'in that case, the program will seek the value as it was written on the input sheet named Provenance
            If m_DFManager.m_SheetsChart.Exists(Provenance) Then
                Set ShObj = m_DFManager.m_SheetsChart(Provenance)
            Else
                Set ShObj = m_DFManager.m_SheetsChart.Items(m_DFManager.m_SheetsChart.count - 1)
                If Not Provenance = "Source" And Not Provenance = "Src" Then
                    MsgBox "Warning : Sheet name " & Provenance & " was not recognized, so the source sheet was chosen instead."
                End If
            End If
                
                With DFLI
                    .TargetShName = ShObj.m_SheetName
                    .StartRowNum = ShObj.m_NumHeaderRows + 1
                    .EndRowNum = ShObj.m_FurthestDataRow
                    
                    If ShObj.m_LocalDFChart.Exists(FieldName) Then
                        .ColNum = Val(Replace(ShObj.m_LocalDFChart(FieldName), "_", ""))
                        'take the first colnum that is linked to the IDF
                        FoundSuccess = True
                    End If
                End With
        End If
    End If

    If Not FoundSuccess Then
        MsgBox "Error : the given internal data field location " & InputIDF & " couldn't be found."
        Err.Raise 2002
    Else
       Set m_GetDFLocation = DFLI
    End If

End Function

'm_GetWholeDFColumnAddress is a helper function used
'by m_OutputValuesOnSheet when it has to determine which
'ranges to take as input and/or output of range-wise data field instructions
'accepting internal data field names as arguments.

'I am not sure that such instructions will still be processed in m_OutputValuesOnSheet;
'they will probably be moved to the PostProcessing class module.

Public Function m_GetWholeDFColumnAddress(InputIDF As String) As String
    Dim DFLI As clsMoreDFLocationInfo
    Set DFLI = m_GetDFLocation(InputIDF)
    
    With DFLI
        Dim RangeAddress As String
        RangeAddress = .TargetShName & "!" & Range(Cells(.StartRowNum, .ColNum), Cells(.EndRowNum, .ColNum)).Address
    End With
    
    m_GetWholeDFColumnAddress = RangeAddress
End Function

Public Function m_GetFirstCellInDFColumnAddress(InputIDF As String) As String

    Dim DFLI As clsMoreDFLocationInfo
    Set DFLI = m_GetDFLocation(InputIDF)
    
    With DFLI
        Dim CellAddress As String
        CellAddress = .TargetShName & "!" & Cells(.StartRowNum, .ColNum).Address(RowAbsolute:=False)
    End With
    
    m_GetFirstCellInDFColumnAddress = CellAddress
    
End Function


'The system will go through the IDFs in two passes :
'1- do the IDFs that do direct copies (no modifications) or direct translations (i.e. with the Replace and Fix Date-type instructions);
'   i.e. fields that have a source data fields assigned to them
'2- do the IDFs which have instructions that depend on other fields and as such, were not directly assigned a source field.
'   These are called "calculated" fields.

Public Sub m_DoInstructionsOnChartsSheet(ChartsSheetName As String, OutputSheetName As String)

    Call m_DFManager.m_InitFromChartsSheet(ChartsSheetName)
    Call FindAllTransChartsOnSheet(ChartsSheetName, Me, True)
    
    If Not SheetExists(OutputSheetName) Then
        Sheets.Add(after:=ActiveSheet).Name = OutputSheetName
    End If
    
    m_OutputSheetName = OutputSheetName

    Dim NumSrcSheets As Integer: NumSrcSheets = m_DFManager.m_SheetsChart.count

    Dim SrcSheet As clsDataSheet: Set SrcSheet = m_DFManager.m_SheetsChart.Items(NumSrcSheets - 1)
    Dim SrcShName As String: SrcShName = SrcSheet.m_SheetName
    Dim SrcFirstRow As Integer: SrcFirstRow = SrcSheet.m_NumHeaderRows + 1
    Dim SrcLastRow As Integer: SrcLastRow = SrcSheet.m_FurthestDataRow
    Dim SrcLastCol As Integer: SrcLastCol = SrcSheet.m_FurthestHRColumn
    
    m_LastOutputRow = SrcLastRow - SrcSheet.m_NumHeaderRows + 1
    
    If NumSrcSheets > 1 Then
        MsgBox ("Warning : Many source sheets are present in the data fields chart, but this is a data processing operation. Only the last one, " _
        & SrcShName & ", will be kept.")
    End If
    
    Sheets(OutputSheetName).Select
    ActiveSheet.Cells.ClearContents
    ActiveSheet.Cells.ClearFormats
    Dim OutputCol As Integer: OutputCol = 0
    Dim SrcLocation As clsDFLocationInfo
    
    Dim ArgFinder As Object: Set ArgFinder = CreateObject("VBScript.RegExp")
    ArgFinder.Global = True
    ArgFinder.MultiLine = False
    ArgFinder.IgnoreCase = True
    
    Dim CurExcelFunc As New clsCallFunc
    
    Dim CalculatedIDFs As New Collection
   
    For Each IDF In m_DFManager.m_DataFieldsChart.Items
         
        If Not IDF.m_NameOnOutputSheet = "" Then
            OutputCol = OutputCol + 1
            Cells(1, OutputCol).Value = IDF.m_NameOnOutputSheet
            
            Dim FinalOutRange As Range: Set FinalOutRange = Range(Cells(2, OutputCol), Cells(m_LastOutputRow, OutputCol))
            Dim TempOutRange1 As Range, TempOutRange2 As Range
            
            If SrcSheet.m_LocalDFChart.Exists(IDF.m_InternalOfficialName) Then
                'take the first dfloc from the last sheet
                Dim SrcColNum As Integer: SrcColNum = Val(Replace(SrcSheet.m_LocalDFChart(IDF.m_InternalOfficialName), "_", ""))
                Dim SrcRange As Range: Set SrcRange = Range(Sheets(SrcShName).Cells(SrcFirstRow, SrcColNum), Sheets(SrcShName).Cells(SrcLastRow, SrcColNum))
                
                If IDF.m_Instructions = "" Then
                    SrcRange.Copy Cells(2, OutputCol)
                Else
                    Set TempOutRange1 = Range(Sheets(SrcShName).Cells(SrcFirstRow, SrcLastCol + 1), Sheets(SrcShName).Cells(SrcLastRow, SrcLastCol + 1))
                    Set TempOutRange2 = TempOutRange1.Offset(0, 1)
                    
                    SrcRange.Copy TempOutRange1
                    SrcRange.Copy TempOutRange2
                    
                    Dim AllDFInstr() As String
                    AllDFInstr = Split(IDF.m_Instructions, "&", , vbBinaryCompare)
                    
                    For Each DFInstr In AllDFInstr
                        DFInstr = Trim(DFInstr)
                        Dim InstrIsValid As Boolean: InstrIsValid = False
                        Dim FirstParenth As Integer: FirstParenth = InStr(DFInstr, "(")
                        Dim SecondParenth As Integer: SecondParenth = Len(DFInstr)
                        
                        If DFInstr Like "TextReplace(*)" Or DFInstr Like "CellReplace(*)" Or DFInstr Like "SlowTextReplace(*)" Then
                            TCName = Mid(DFInstr, FirstParenth + 1, SecondParenth - FirstParenth - 1)
                            Dim InstrName As String: InstrName = Left(DFInstr, FirstParenth - 1)
                            If m_TranslationCharts.Exists(TCName) Then
                                If InstrName = "TextReplace" Then
                                    Call m_TranslationCharts(TCName).m_QuickTranslateCells(TempOutRange1, TempOutRange2, 0)
                                ElseIf InstrName = "CellReplace" Then
                                    Call m_TranslationCharts(TCName).m_QuickTranslateCells(TempOutRange1, TempOutRange2, 1)
                                ElseIf InstrName = "SlowTextReplace" Then
                                    Call m_TranslationCharts(TCName).m_TranslateCells(TempOutRange1, TempOutRange2, 0)
                                End If
                            End If
                            InstrIsValid = True
                        ElseIf DFInstr = "FixDate()" Then
                            Call FixAllDatesInRange(TempOutRange1, TempOutRange2)
                            InstrIsValid = True
                        Else
                            MsgBox ("Warning : Instruction " & DFInstr & " for field " & IDF.m_InternalOfficialName & " was not recognized and will be ignored.")
                        End If
                        
                        If InstrIsValid Then
                            TempOutRange2.Copy TempOutRange1
                        End If
                    Next DFInstr
                    
                    TempOutRange2.Copy FinalOutRange
                End If
            ElseIf Not IDF.m_Instructions = "" Then
                CalculatedIDFs.Add IDF
            End If
        End If
    Next IDF
    
    For Each IDF In CalculatedIDFs
        DFInstr = IDF.m_Instructions
        FirstParenth = InStr(DFInstr, "(")
        SecondParenth = Len(DFInstr)
        Dim InputIDFName As String
        Dim SrcRangeString As String
        OutputCol = IDF.m_ColOnOutputSheet
        Set FinalOutRange = Range(Cells(2, OutputCol), Cells(m_LastOutputRow, OutputCol))
        
        If DFInstr Like "TextToMultiChoice(*)" Then
            Dim Args As String
            Args = Mid(DFInstr, FirstParenth + 1, SecondParenth - FirstParenth - 1)
            Dim ArgsArr() As String: ArgsArr = Split(Args, ";")
            InputIDFName = Trim(ArgsArr(0)): TCName = Trim(ArgsArr(1))
            If m_TranslationCharts.Exists(TCName) Then
                SrcRangeString = m_GetWholeDFColumnAddress(InputIDFName)
                Call m_TranslationCharts(TCName).m_ConvertTextToMCRForColumn(Range(SrcRangeString), Range(Cells(2, OutputCol), Cells(m_LastOutputRow, OutputCol)))
            End If
        
        ElseIf DFInstr Like "ExcelFormula*" Then
            Dim FuncText As String: FuncText = Right(DFInstr, Len(DFInstr) - Len("ExcelFormula"))
            
            Call CurExcelFunc.m_SetOrigFuncStr(FuncText, 1)
            ArgFinder.Pattern = "_(.*?)_" 'the ? makes .* lazy; it tries to get matches that are as short as possible
            Set ArgMatches = ArgFinder.Execute(FuncText)
            For Each ArgMatch In ArgMatches
                Call CurExcelFunc.m_AddNewArg(CStr(ArgMatch))
                InputIDFName = ArgMatch.submatches(0)
                Dim FirstSrcCellAddress As String
                SrcRangeString = m_GetFirstCellInDFColumnAddress(InputIDFName)
                Call CurExcelFunc.m_SetArgValue(CStr(ArgMatch), SrcRangeString)
            Next ArgMatch
                        
            CurExcelFunc.m_ReplaceIDFNamesWithVals
            FinalOutRange.Formula = CurExcelFunc.m_CurFuncStr
        End If
    Next IDF
    
    Set CalculatedIDFs = Nothing
    
    TempOutRange1.Clear
    TempOutRange2.Clear
    
End Sub

