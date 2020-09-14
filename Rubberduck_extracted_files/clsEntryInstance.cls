VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntryInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_ShNum As Integer
Public m_ShName As String
Public m_RowNum As Integer
Public m_RowOnOutputSheet As Integer
Public m_OwnerUID As String


Public Function m_GetLocationString() As String
    Dim Indic As String: Indic = " : "
    m_GetLocationString = m_ShName & Indic & m_RowNum
End Function

