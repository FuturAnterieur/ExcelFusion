VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsECMNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_Name As String 'basically, an ECMNode is a group of at least one Entry Specifier(s) name(s)
Public m_DirectParents As Scripting.Dictionary 'i.e. direct parents
Public m_DirectChildren As Scripting.Dictionary
Public m_ParentNodesToCheck As Scripting.Dictionary 'run-thru of the tree in the upwards direction
Public m_ChildNodesToCheck As Scripting.Dictionary
Public m_NumberOfChildLayers As Integer 'i.e. the length of the deepest tree branch starting from that node going downwards
Public m_NumberOfParentLayers As Integer 'i.e. the length of the deepest tree branch starting from that node going upwards

Public Sub Class_Initialize()
    Set m_DirectParents = New Scripting.Dictionary
    Set m_DirectChildren = New Scripting.Dictionary
    Set m_ParentNodesToCheck = New Scripting.Dictionary
    Set m_ChildNodesToCheck = New Scripting.Dictionary
    m_NumberOfChildLayers = -1
    m_NumberOfParentLayers = -1
End Sub

Public Function m_IsolateInIDString(FullIDStr As String) As String
    Dim Res As String
    Res = ""
    
    Dim ESNames() As String
    ESNames = Split(m_Name, " & ", , vbTextCompare)
    
    Dim SubIDStrs() As String
    SubIDStrs = Split(FullIDStr, " & ", , vbTextCompare)
    
    For Each SubIDStr In SubIDStrs
        For Each ESName In ESNames
            If InStr(SubIDStr, ESName & " = ") > 0 Then
                Res = Res & " & " & SubIDStr
                Exit For
            End If
        Next ESName
    Next SubIDStr
    
'    For Each ESName In ESNames
'        For Each SubIDStr In SubIDStrs
'            If InStr(SubIDStr, ESName & " = ") > 0 Then
'                Res = Res & " & " & SubIDStr
'                Exit For
'            End If
'        Next SubIDStr
'    Next ESName

    If Res <> "" Then
        Res = Right(Res, Len(Res) - Len(" & "))
        
    End If
    m_IsolateInIDString = Res

End Function

