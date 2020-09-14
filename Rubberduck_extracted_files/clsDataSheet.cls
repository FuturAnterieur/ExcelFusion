VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public m_SheetNum As Integer
Public m_SheetName As String
Public m_InternalIndex As Integer
Public m_ColOnDFChart As Integer
Public m_Qualifiers As String
Public m_DataFusionPriority As Integer
Public m_LocalEntrySpecifiers As Collection
Public m_ECMNode As String
Public m_LocalDFChart As Scripting.Dictionary 'of internal data field names
Public m_NormHeaderNotAddedYet As Boolean

Public m_NormHeaderRow As Integer
Private m_AreDataFieldsNorm As Boolean
Private m_IsInit As Boolean

Public m_NumHeaderRows As Integer, m_FurthestHRColumn As Integer, m_FurthestDataRow As Integer
Public m_NumTempColumns As Integer
Public Sub Class_Initialize()

Set m_LocalDFChart = New Scripting.Dictionary

m_FurthestHRColumn = 0
m_NormHeaderNotAddedYet = True
m_InternalIndex = 0
m_ColOnDFChart = 0
m_IsInit = False
m_AreDataFieldsNorm = False
m_Qualifiers = ""
m_NumHeaderRows = 1
m_DataFusionPriority = 0

m_NumTempColumns = 0

Set m_LocalEntrySpecifiers = New Collection
End Sub

Public Sub m_DoInit(SheetIndex As Integer, ColOnDFChart As Integer, SheetName As String, Optional Qualifiers As String = "")
    
    If (m_IsInit = False) Then
        m_SheetName = SheetName
        m_SheetNum = Sheets(SheetName).Index
        m_InternalIndex = SheetIndex
        m_ColOnDFChart = ColOnDFChart
        
        If Not Qualifiers = "" Then
            m_Qualifiers = Qualifiers
            Dim QualifFinder As Object
            Set QualifFinder = CreateObject("VBScript.RegExp")
            QualifFinder.Global = False
            QualifFinder.MultiLine = False
            QualifFinder.IgnoreCase = True
            QualifFinder.Pattern = "(NumHR|Header Rows)\s*=\s*(\d+)"
            If QualifFinder.Test(m_Qualifiers) Then
                Set Games = QualifFinder.Execute(m_Qualifiers)
                m_NumHeaderRows = Games(0).submatches(1)
            End If
            QualifFinder.Pattern = "DFP\s*=\s*(\d+)"
            If QualifFinder.Test(m_Qualifiers) Then
                Set Games = QualifFinder.Execute(m_Qualifiers)
                m_DataFusionPriority = Games(0).submatches(0)
            End If
        End If
        
    
        Dim StartOfHeaderRow As Range
        Dim i As Integer, CurHRColumn As Integer
    
    
        For i = 1 To m_NumHeaderRows
            Set EndOfHeaderRow = Sheets(m_SheetName).Rows(i).Find("*", after:=Cells(i, 1), SearchDirection:=xlPrevious, LookAt:=xlPart, LookIn:=xlValues)
            If Not EndOfHeaderRow Is Nothing Then
            
                CurHRColumn = EndOfHeaderRow.Column
                If CurHRColumn > m_FurthestHRColumn Then
                    m_FurthestHRColumn = CurHRColumn
                End If
            End If
        Next i
    
        Dim CurColEndRow As Integer
        Dim EndOfCurCol As Range
        For i = 1 To m_FurthestHRColumn
            Set EndOfCurCol = Sheets(m_SheetName).Columns(i).Find("*", after:=Cells(1, i), SearchDirection:=xlPrevious, LookAt:=xlPart, LookIn:=xlValues)
            If Not EndOfCurCol Is Nothing Then
                CurColEndRow = EndOfCurCol.Row
                
                If CurColEndRow > m_FurthestDataRow Then
                    m_FurthestDataRow = CurColEndRow
                End If
            End If
        Next i
        m_IsInit = True
        
        ' "Parse" qualifiers.
        'data fusion priority is explained in the notes to clsDFSystem
        
        
    End If
End Sub

Public Sub m_ExtractDataFields(OutputCell As Range, Optional InSeparateColumns As Boolean = False, Optional Separator As String = " _or_ ")
    
    'only the version that is NOT in separate columns is programmed right now.
    
    Dim i As Integer, j As Integer
    Dim FirstOutputCell As Range, CurCell As Range
    Dim OutputStr As String
    
    Set FirstOutputCell = OutputCell.Cells(1, 1)
    Dim OutputRng As Range: Set OutputRng = Range(FirstOutputCell, OutputCell.Cells(m_FurthestHRColumn, 1))
    OutputRng.Clear
    Dim DictOfExistingNames As Scripting.Dictionary
    Dim ColToDFName As Scripting.Dictionary
    Set DictOfExistingNames = New Scripting.Dictionary
    Set ColToDFName = New Scripting.Dictionary
    Dim ExistingCols As String
    Dim ColsForThisTerm() As String
    Dim Col As Integer
    
    'so first : scan the data fields
    For i = 1 To m_FurthestHRColumn
        For j = 1 To m_NumHeaderRows
            Set CurCell = Sheets(m_SheetName).Cells(j, i)
            If Not CurCell.Text = "" Then
                If Not DictOfExistingNames.Exists(CurCell.Text) Then
                    DictOfExistingNames.Add Key:=CurCell.Text, Item:=i
                Else
                    ExistingCols = DictOfExistingNames(CurCell.Text)
                    ColsForThisTerm = Split(ExistingCols, ",")
                    If Val(ColsForThisTerm(UBound(ColsForThisTerm))) <> i Then
                        DictOfExistingNames(CurCell.Text) = ExistingCols & "," & i
                    End If
                End If
            End If
        Next j
    Next i
    
    For Each DFTerm In DictOfExistingNames.Keys
        Dim ManyCols As Boolean: ManyCols = False
        If InStr(DictOfExistingNames(DFTerm), ",") > 0 Then
            ManyCols = True
        End If
        
        ColsForThisTerm = Split(DictOfExistingNames(DFTerm), ",")
        For Each StrCol In ColsForThisTerm
            Col = Val(StrCol)
            Dim TermToOutput As String
            If Not ManyCols Then
                TermToOutput = DFTerm
            Else
                TermToOutput = DFTerm & "_inCol" & Col
            End If
            Dim TargetCell As Range
            If Not InSeparateColumns Then
                Set TargetCell = FirstOutputCell.Offset(Col - 1, 0)
                TargetCell.Value = TargetCell.Text & Separator & TermToOutput
            End If
        Next StrCol
    Next DFTerm
    
    For Each EndCell In OutputRng
        If Not EndCell.Value = "" Then
            EndCell.Value = Right(EndCell.Text, Len(EndCell.Text) - Len(Separator))
        End If
    Next EndCell
    
End Sub





