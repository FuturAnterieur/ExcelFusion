VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_UIDString As String   'the criteria mashed together, I guess

Public m_WhereCanIBeFound As Collection 'of clsEntryInstance objects
Public m_NbOfRowsOnEachSheet As Object
Public m_NbOfPolyRowSheets As Integer 'number of sheets where the entry appears more than once i.e. has doubles.
Public m_ValidInstances As Collection 'collection of m_WhereICanBeFound indices
Public m_InvalidInstances As Collection
Public m_Parents As Object
Public m_Children As Object

Public m_MainValidMatchGrp As clsMatchGroup
Public m_HasInstancesOnSheets As Boolean

Public m_Format As String 'the entry specifiers that form its UIDString.


'Public m_MatchList As Collection 'a collection of several collections; each subcollection represents a matching group and contains
''indexes to the participating m_WhereCanIBeFound entries. I did this in case many separate matching groups have to be established for a same entry.
''(which would be kinda crazy but, you know, just in case).

Public Sub Class_Initialize()

    Set m_WhereCanIBeFound = New Collection
    
    'Set m_MatchList = New Collection
    
    Set m_ValidInstances = New Collection
    Set m_InvalidInstances = New Collection
    
    Set m_NbOfRowsOnEachSheet = CreateObject("Scripting.Dictionary")
    m_NbOfPolyRowSheets = 0

    Set m_Parents = CreateObject("Scripting.Dictionary")
    Set m_Children = CreateObject("Scripting.Dictionary")
    
    Set m_MainValidMatchGrp = Nothing
    m_HasInstancesOnSheets = False 'This variable indicates whether or not the entry has actual instances, associated with sheet data, or is just a "virtual" entry
    'whose existence was deducted from instances of its children
End Sub

Public Function m_GetMainLocationString() As String
    Dim LocationStr As String
    If m_MainValidMatchGrp Is Nothing Then
        LocationStr = ""
    Else
        LocationStr = m_MainValidMatchGrp.m_GetLocationString
    End If
    m_GetMainLocationString = LocationStr

End Function
