VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntrySpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_IDFName As String



Public Function m_IsolateInIDString(FullIDStr As String) As String
    Debug.Print FullIDStr
    Dim Res As String
    Res = ""
    
    Dim SubIDStrs() As String
    SubIDStrs = Split(FullIDStr, " & ", , vbTextCompare)
    For Each SubIDStr In SubIDStrs
        If InStr(SubIDStr, m_IDFName & " = ") > 0 Then
            Res = SubIDStr
            Exit For
        End If
    Next SubIDStr
    
    m_IsolateInIDString = Res
End Function
