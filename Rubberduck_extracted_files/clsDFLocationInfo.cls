VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDFLocationInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_ShNum As Integer 'the sheet's index in the clsDFSystem m_SheetsChart dictionary.
Public m_ColNum As Integer
Public m_DFPValue As Integer
Sub Class_Initialize()
    
    m_DFPValue = 0

End Sub

Public Function m_ConvertToString() As String

    m_ConvertToString = m_ShNum & "&" & m_ColNum

End Function
