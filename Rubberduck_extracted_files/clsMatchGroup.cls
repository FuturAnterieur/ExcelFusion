VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public m_OwnerUID As String 'the UID of the main entry that possesses the group. Always a parent
Public m_Participants As Collection 'finally, this will be a collection clsEntryInstance objects
Public m_RowOnOutputSheet As Integer
Public m_IsDuplicate As Boolean 'indicates whether or not the MatchGrp is dedicated to an invalid double, or to a group of valid matching
'entry instances (i.e. a Solo-Sheets match group)


Public Sub Class_Initialize()

    Set m_Participants = New Collection

End Sub

Public Function m_GetLocationString() As String
    Dim LocationStr As String
    
    Dim LocSep As String
    LocSep = ", "
    LenLocSep = Len(LocSep)
    
    If m_Participants.count = 0 Then
        LocationStr = ""
    Else
        For Each InstanceLoc In m_Participants
            LocationStr = LocationStr & LocSep & InstanceLoc.m_GetLocationString()
        Next InstanceLoc
        LocationStr = Right(LocationStr, Len(LocationStr) - Len(LocSep))
    End If
    m_GetLocationString = LocationStr
End Function
