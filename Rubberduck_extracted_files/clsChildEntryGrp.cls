VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChildEntryGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_Format As String
Public m_Members As Object 'dico of (references to) clsEntry objects
'Public m_IGCCompliantMembers As Object 'dico of clsEntry objects that respect the InterGroup Criteria of the parent clsRelationalFilter
Public m_ConditionEvaluator As clsCallFunc 'the condition(s) that are to be met by members of the group.
Public m_ParentDFSys As clsDFSystem
Public m_IndexedTable As Collection

Public Sub Class_Initialize()
    Set m_Members = CreateObject("Scripting.Dictionary")
    Set m_IndexedTable = New Collection
    'Set m_IGCCompliantMembers = CreateObject("Scripting.Dictionary")
End Sub

Public Sub m_SetFormat(frmt As String)
    m_Format = frmt
End Sub

Public Function m_GetFormat()
    m_GetFormat = m_Format
End Function

Public Function m_GetName()
    m_GetName = m_Name

End Function
Public Sub m_SetName(Name As String)
    m_Name = Name
End Sub

'Public Sub m_SetConditionEvaluator(condeval As clsCallFunc)
'    Set m_ConditionEvaluator = condeval
'End Sub

Public Function m_PassTestForApplicant(Applicant As clsEntry) As Boolean
    'Debug.Print ("now testing " & Applicant.m_UIDString & " for child entry group " & m_Name)
    Dim Res As Boolean: Res = False
    If Applicant.m_Format = m_Format Then
        If m_ConditionEvaluator.m_ArgNamesList.count > 0 And Not Applicant.m_MainValidMatchGrp Is Nothing Then
            Dim TargetRowOnOS As Integer, TargetColOnOS As Integer
            For Each IDFNameWithUS In m_ConditionEvaluator.m_ArgNamesList
                Dim IDFName As String: IDFName = Right(IDFNameWithUS, Len(IDFNameWithUS) - 1)
                TargetRowOnOS = Applicant.m_MainValidMatchGrp.m_RowOnOutputSheet
                TargetColOnOS = m_ParentDFSys.m_DFManager.m_DataFieldsChart(IDFName).m_ColOnOutputSheet
                Dim TargetCell As Range: Set TargetCell = Sheets(m_ParentDFSys.m_OutputSheetName).Cells(TargetRowOnOS, TargetColOnOS)
                Call m_ConditionEvaluator.m_SetArgValue(CStr(IDFNameWithUS), TargetCell.Address)
            Next IDFNameWithUS
        
            Call m_ConditionEvaluator.m_ReplaceIDFNamesWithVals
           ' Debug.Print "Evaluating " & m_ConditionEvaluator.m_CurFuncStr
            
            Dim Val As Integer
            Val = CInt(m_ConditionEvaluator.m_EvalExcelExpression)
            If Val > 0 Then
                Res = True
                m_Members.Add Key:=Applicant.m_UIDString, Item:=Applicant
                m_IndexedTable.Add Applicant.m_UIDString
            End If
        Else
            'If no arguments (and by extension, any function depending on the Applicant's Data) have been specified,
            'then any Applicant with the correct format and the conditions below will be admitted into the group.
        
            Dim WithValidInstances As Boolean: WithValidInstances = Not Applicant.m_MainValidMatchGrp Is Nothing
            Dim WithoutInstances As Boolean: WithoutInstances = Applicant.m_HasInstancesOnSheets = False
            
            'Not having a Valid Match Group can mean two things :
            'either Case 1 : the Applicant has instances on the source sheets, but they all are duplicates, or
            'either Case 2 : the Applicant just doesn't have any instances on the source sheets, and was thus
            'created as a parent to one or more child entries that do have instances on the source sheets.
            'We want to include all Applicants from Case 2 (even if arguments were given), but none from Case 1.
            If WithValidInstances Or WithoutInstances Then
                m_Members.Add Key:=Applicant.m_UIDString, Item:=Applicant
                m_IndexedTable.Add Applicant.m_UIDString
            End If
        End If
    End If
    m_PassTestForApplicant = Res
End Function


