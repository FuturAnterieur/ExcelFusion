VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_OrigFuncStr As String
Public m_FunctionType As Integer '0 = generic, local macro mode; 1 = Evaluate Worksheetfunction string
Public m_CurFuncStr As String
Public m_FuncName As String
Public m_ArgSep As String
Public m_ArgsString As String 'i.e. what goes inside the parenthesis, split by commas (at least in the English version of excel)
Private m_IDFValDict As Object 'IDF names must be preceded by underscores to avoid writing over text that is not meant to be replaced.
Public m_ArgNamesList As Collection
Public m_ProcessedArgs As Collection

Public Sub Class_Initialize()
    Set m_IDFValDict = CreateObject("Scripting.Dictionary")
    Set m_ProcessedArgs = New Collection
    Set m_ArgNamesList = New Collection
    m_ArgSep = ","
End Sub
Public Sub m_SetOrigFuncStr(ofs As String, Optional FunctionType As Integer = 1)
    m_OrigFuncStr = ofs
    m_FunctionType = FunctionType
End Sub
Public Sub m_AddNewArg(ArgName As String, Optional IsString As Boolean = False, Optional ConvertToInt As Boolean = False)
    If Not m_IDFValDict.Exists(ArgName) Then
        m_IDFValDict.Add Key:=ArgName, Item:=New clsWSArg
        m_IDFValDict(ArgName).m_IsString = IsString
        m_IDFValDict(ArgName).m_ConvertToInt = ConvertToInt
        m_ArgNamesList.Add ArgName
    End If
End Sub

Public Sub m_SetArgValue(ArgName As String, ArgVal As Variant)
    If m_IDFValDict.Exists(ArgName) Then
        m_IDFValDict(ArgName).m_ArgVal = ArgVal
        'm_IDFValDict(ArgName).m_IsString = IsString
    End If
End Sub

Public Sub m_ReplaceIDFNamesWithVals()
    m_CurFuncStr = m_OrigFuncStr
    For Each IDFNameWithUS In m_IDFValDict.Keys
        If m_IDFValDict(IDFNameWithUS).m_IsString Then
            m_CurFuncStr = Replace(m_CurFuncStr, IDFNameWithUS, Chr(34) & m_IDFValDict(IDFNameWithUS).m_ArgVal & Chr(34))
        ElseIf m_IDFValDict(IDFNameWithUS).m_ConvertToInt Then
            m_CurFuncStr = Replace(m_CurFuncStr, IDFNameWithUS, CLng(m_IDFValDict(IDFNameWithUS).m_ArgVal))
        Else
            m_CurFuncStr = Replace(m_CurFuncStr, IDFNameWithUS, m_IDFValDict(IDFNameWithUS).m_ArgVal)
        End If
    Next IDFNameWithUS
End Sub

Public Sub m_ExtractNameAndArgs()
    Dim SplitFuncName() As String
    SplitFuncName = Split(m_CurFuncStr, "(", 2, vbBinaryCompare)
    m_FuncName = SplitFuncName(0)
    m_ArgsString = Left(SplitFuncName(1), Len(SplitFuncName(1)) - 1) 'remove closing parenthesis
End Sub

Public Sub m_ExtractProcessedArgs()
    
    Dim ArgFinder As Object
    Set ArgFinder = CreateObject("VBScript.RegExp")
    
    ArgFinder.Global = True
    ArgFinder.MultiLine = False
    ArgFinder.IgnoreCase = True
    ArgFinder.Pattern = "(\w(\w|\s)+)\s*(,|$)"
    'maybe some spaces, then at least a wordchar followed by a series of other wordchars or spaces and closed by a wordchar,
    'this sequence  being the full argument (i.e. "Age" or "OpType = Lobectomy"), ended by a comma or a closing parenthesis.
    'the next part is to replace any IDFName found in these arguments by the corresponding value from the desired sheet.
    'but I could also do it before Arg Splitting, in ReplaceIDFNamesWithVals
    Set Allumettes = ArgFinder.Execute(m_ArgsString)
    For Each Allumette In Allumettes
        m_ProcessedArgs.Add Allumette.submatches(0)
    Next Allumette
    
End Sub

Public Function m_EvalExcelExpression() As Variant

    m_EvalExcelExpression = Application.Evaluate(m_CurFuncStr)
End Function

Public Function m_CallLocalMacro() As Variant 'still doesn't seem to work properly with many Variant arguments

    ReDim Args(1 To m_ProcessedArgs.count) As Variant
    For i = 1 To m_ProcessedArgs.count
        Args(i) = m_ProcessedArgs(i)
        Debug.Print (Args(i))
    Next i
    
    m_CallLocalMacro = CallByName(Me, m_FuncName, VbMethod, Args)
End Function

Public Function GetOutput() As Variant
    
    If m_FunctionType = 0 Then
        GetOutput = m_CallLocalMacro
    ElseIf m_FunctionType = 1 Then
        GetOutput = m_EvalExcelExpression
    End If
    
End Function

