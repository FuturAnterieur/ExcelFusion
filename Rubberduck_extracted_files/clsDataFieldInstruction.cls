VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataFieldInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_CallingObject As Object
Public m_FuncStr As String
Public m_Type As Integer
Public m_OtherArgs As Collection

Public Sub Class_Initialize()
    m_FuncStr = ""
    Set m_OtherArgs = New Collection
    m_Type = 0
    'DF Instruction Types :
    '0 : the only arg is the CurValue of the cell being evaluated, and the function returns a single value (which becomes the new CurValue)
    '1 : the args are the CurValue, followed by a single target cell address; to be called before other instructions, with the value
    'placed somewhere on the sheet.

End Sub
