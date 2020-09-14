VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsECMTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public m_NumberOfLayers As Integer

Public m_ConnectMap As Scripting.Dictionary
Public m_NodesOnTopLayer As Scripting.Dictionary 'nodes without parents, but with children
Public m_NodesOnBottomLayer As Scripting.Dictionary 'nodes without children, but with parents
Public m_OrphanedNodes As Scripting.Dictionary 'without a family

Public Sub Class_Initialize()
    
    Set m_ConnectMap = New Scripting.Dictionary
    Set m_NodesOnTopLayer = New Scripting.Dictionary
    Set m_NodesOnBottomLayer = New Scripting.Dictionary
    Set m_LonesomeNodes = New Scripting.Dictionary

End Sub

   'EXPLANATIONS
    
    'Now, entries will get assigned to child or parent roles based on their format,
    'and their format corresponds to the list of entry specifiers that are used
    'to define them. For example, if the IDFs PNum, OpDate and OpType have the ES qualifier
    'in the data fields chart, and sheet PatientDataA only has PNum and OpDate, then all entries
    'that physically exist (i.e. have a row) on PatientDataA will share the "PNum & OpDate" format.
    
    'That is why the reference tree (m_ECMTree, a clsECMTree object) used to link entries together has clsECMNode objects (and not clsEntrySpecifier) as nodes.
    'An ECMNode is mostly defined by its name, which corresponds to an entry format : a series of (at least one) entry specifier names.
    
    'The idea in m_EstablishEssentialData is to build the m_ECMTree so that we have a reference for "which type of entry is
    'a parent to which other type of entry" when we will match entries together in m_MatchEntriesAcrossSheets.
    
    'HOW TO ESTALBISH A LIST OF PARENT-CHILD LINKS
    '
    '1. Write "ECM Graph" (and only "ECM Graph") in any cell with free space under it on the charts sheet.
    '2. Under the "ECM Graph" cell, each cell will be interpreted as estalbishing a single parent-child link.
    '3. Each parent-child link follows the following format : [parent node name] -> [child node name]
    '   where [parent node name] and [child node name] are valid ECMNode names, i.e. they only contain
    '   sets of entry specifier names split by ampersands ("&" characters). Writing a node name in the ECM Graph
    '   creates the actual ECMNode object, in addition to linking it to its parents/children.
    '   Example : PNum -> PNum & OpDate
    '
    '   Now, specifiyng ECMNode connections follows some rules and restrictions :
    '   1. As was said above, every node name has to be exclusively made of valid entry specifier names.
    '       Not respecting this will cause an error. Also, these entry specifier names have to follow
    '       the order of appearance of entry specifiers in the data fields chart; maybe TODO : a ECM Node name reorderer.
    
    '
    '   2. For any given Parent - Child link, the Child node's name has to contain all the entry specifiers found in
    '        the parent node's name.
    '        This may sound unjustified at first, but without this, it would have been way more complex to
    '        match parent and child entries together. In the current system, for a child entry to properly recognize its parent,
    '        the child entry has to contain the entry specifiers that will be used to identify said parent.
    '        Failure to comply to this will cause an error. Its effect on data fusion would be limited
    '        (it would just prevent parent-to-child data fusion from working), but it
    '         can cause the relational filtering part to crash if it is left running.
    '   3. It is entirely possible, however, to create tree nodes whose format does not exist
    '       among the entries present on the source sheets. For example, it is possible to specify
    '       a "PNum -> PNum & OpDate" link even if all the source sheets have both PNum and OpDate as entry specifiers.
    '       Doing so can be useful for relational filtering later on.
    '   4. The tree cannot be circular. Between 2 nodes on the tree, many paths may exist,
    '       but they cannot be circular i.e. a parent becoming the child of its child.
    '       This condition is logically enforced by the m_NodesMetInPassing dictionary,
    '       found in the recursive function m_SetParentCollectionForNode. However,
    '       the system first checks for condition 2, and if condition 2 is respected,
    '       then it is pretty much impossible for a parent to become its child's child in the first place.
    '
    '   4. In the ECM Graph, the user may specify any number of valid parent-child links in any order.
    '       The m_CreateECMParentChildLink function keeps track of which nodes are on top and
    '       on the bottom of the tree, and which links exist between them. This is what
    '       is necessary for the m_SetAllParentCollections function to do its job :
    '       to know the list of all parents and all children for each node.
    '
    '   WHY ALL OF THIS
    '   Creating parent-child relationships has two main uses: 1. fusing data from parent to child entries,
    '   and 2. applying relational filtering to the list of entries once it has been outputted.
    '   Behaviour in those two use cases can be explained through clsDFSystem's m_MatchEntriesAcrossSheets function,
    '   so go see this function if you want more explanation. But to put it in a basic way,
    '   create parent-child links if you want data from Entries following the Parent format
    '   to be matched with the corresponding Entries of the Child Format.
    '   Also, remember that an Entry Specifier doesn't have to have an ECM node
    '   named after it to function properly.


Public Sub m_ReadECMGraphOnChartsSheet(ChartsSheetName As String, EntrySpecifierDict As Scripting.Dictionary, Optional SheetExistenceChecked As Boolean = False)

    If Not SheetExistenceChecked Then
        If Not SheetExists(ChartsSheetName) Then
            MsgBox ("Specified trans chart sheet name " & ChartsSheetName & " does not exist in this workbook. Aborting.")
            Err.Raise 1998
        End If
    End If
    
    Set ChartsSheet = Worksheets(ChartsSheetName)

    Dim ECMGraphCell As Range
    Set ECMGraphCell = ChartsSheet.Cells.Find("ECM Graph", after:=Cells(1, 1), LookAt:=xlWhole, MatchCase:=False, SearchDirection:=xlNext)
    If Not ECMGraphCell Is Nothing Then
       Dim ECMGraphRange As Range: Set ECMGraphRange = Range(ECMGraphCell.Offset(1, 0), ECMGraphCell.End(xlDown))
       
       For Each ECMLinkCell In ECMGraphRange
           Dim LinkMembers() As String
           LinkMembers = Split(ECMLinkCell.Text, "->", , vbTextCompare)
           Call m_CreateECMParentChildLink(Trim(LinkMembers(0)), Trim(LinkMembers(1)), EntrySpecifierDict)
       Next ECMLinkCell
       
       'Once the ECM graph has been read, we are only interested, for the sake of this program,
       'to know the list of all parents and all children for each node. This is what
       'm_SetAllParentCollections does.
       Call m_SetAllParentCollections
    End If

End Sub


Public Sub m_CreateECMParentChildLink(ParentName As String, ChildName As String, EntrySpecifierDict As Scripting.Dictionary)

    Dim ESsInParent() As String
    ESsInParent = Split(ParentName, "&", , vbBinaryCompare)
    For Each ESInParent In ESsInParent
        If Not InStr(ChildName, ESInParent) > 0 Then
            MsgBox "Error. The given child name " & ChildName & " does not contain the given parent component " & ESInParent & "."
            Err.Raise 1000, "m_CreateECMParentChildLink", "parent name not in child name"
        End If
    Next ESInParent

    If ChildName = ParentName Then
        MsgBox "Error. The name " & ChildName & " was given as its own parent."
        Err.Raise 1000, "m_CreateECMParentChildLink", "parent is same as child"
    End If

    Dim NewParent As Boolean, NewChild As Boolean
    NewParent = False: NewChild = False
    If Not m_ConnectMap.Exists(ParentName) Then
        Call m_CreateECMNode(ParentName, EntrySpecifierDict)
        m_NodesOnTopLayer.Add Key:=ParentName, Item:=ParentName
        NewParent = True
    End If
    If Not m_ConnectMap.Exists(ChildName) Then
        Call m_CreateECMNode(ChildName, EntrySpecifierDict)
        NewChild = True
        m_NodesOnBottomLayer.Add Key:=ChildName, Item:=ChildName
    End If
    
    If m_NodesOnBottomLayer.Exists(ParentName) And m_NodesOnTopLayer.Exists(ChildName) Then
        MsgBox ("Error : A circular link has been mistakenly set between a top node and a bottom node in the ECM Graph.")
        Err.Raise 1004, "m_SetAllParentCollections", "circular link between top and bottom"
        'other circular links can be caught in m_SetAllParentCollections
        '(not counting the "parent name inside child name" condition above, which should just
        'catch all these circular link cases).
    Else
        If m_NodesOnBottomLayer.Exists(ParentName) Then
            m_NodesOnBottomLayer.Remove (ParentName)
        ElseIf m_NodesOnTopLayer.Exists(ChildName) Then
            m_NodesOnTopLayer.Remove (ChildName)
        End If
        
        
        If Not m_ConnectMap(ParentName).m_DirectChildren.Exists(ChildName) And _
           Not m_ConnectMap(ChildName).m_DirectParents.Exists(ParentName) Then
            m_ConnectMap(ParentName).m_DirectChildren.Add Key:=ChildName, Item:=0
            m_ConnectMap(ChildName).m_DirectParents.Add Key:=ParentName, Item:=0
        End If
    
    End If

End Sub
Public Sub m_CreateECMNode(Name As String, EntrySpecifierDict As Scripting.Dictionary)

    Dim Okay As Boolean
    Okay = True
    Dim FaultyESName As String
    Dim ESNames() As String
    ESNames = Split(Name, "&", , vbBinaryCompare)

    For Each ESName In ESNames
        If Not EntrySpecifierDict.Exists(Trim(ESName)) Then
            Okay = False
            MsgBox "Error : the given " & ESName & " does not belong to a recognized entry specifier."
            FaultyESName = ESName
            Exit For
        End If
    Next ESName
    
    If Okay Then
        m_ConnectMap.Add Key:=Name, Item:=New clsECMNode
        m_ConnectMap(Name).m_Name = Name
    Else
        Err.Raise 1001, "clsDFSystem - m_CreateECMNode", "An ESName was not valid; it was " & FaultyESName & "."
    End If
    
End Sub

Public Sub m_SetAllParentCollections()
    Dim UnusedResult As Collection
    For Each BottomNodeName In m_NodesOnBottomLayer.Keys
        Dim StartNodeList As New Scripting.Dictionary
        StartNodeList.RemoveAll
        Debug.Print "New bottom node : " & BottomNodeName
       Set UnusedResult = m_SetParentCollectionForNode(CStr(BottomNodeName), StartNodeList, 1)
    Next BottomNodeName
End Sub
Public Function m_SetParentCollectionForNode(CurNodeName As String, NodesMetInPassing As Scripting.Dictionary, RelativeTreeLevel As Integer) As Collection
    Debug.Print CurNodeName
    Dim CurNodeObj As clsECMNode, PNObj As clsECMNode
    Set CurNodeObj = m_ConnectMap(CurNodeName)
    Dim ReturnedCollection As New Collection
    Set ReturnedCollection = Nothing

    If Not NodesMetInPassing Is Nothing Then
        If Not NodesMetInPassing.Exists(CurNodeName) Then
            NodesMetInPassing.Add Key:=CurNodeName, Item:=RelativeTreeLevel
            Debug.Print ("New node " & CurNodeName & " at level " & RelativeTreeLevel)
        Else
            MsgBox "Error : Node " & CurNodeName & " is being linked reciprocally to the same node."
            Err.Raise 1003, "clsDFSystem - m_SetParentCollectionForNode", "Reciprocal parent-child link between two nodes"
        End If
    End If

    If CurNodeObj.m_ParentNodesToCheck.count > 0 Then
        'Debug.Print "case 1 for " & CurNodeName
        ' this means we stumbled upon a node for which the recursive process
        ' in the corresponding Else clause was already done
        For Each ExistingPNodeName In CurNodeObj.m_ParentNodesToCheck.Keys
            ReturnedCollection.Add ExistingPNodeName
        Next ExistingPNodeName
    Else
        'Debug.Print "case 2 step 1 for " & CurNodeName
        '1. Recursively find all parents
        Set ReturnedCollection = New Collection
        For Each PNodeName In CurNodeObj.m_DirectParents.Keys
                RelativeTreeLevel = RelativeTreeLevel + 1
                ReturnedCollection.Add PNodeName
                For Each GrandpaName In m_SetParentCollectionForNode(CStr(PNodeName), NodesMetInPassing, RelativeTreeLevel)
                    ReturnedCollection.Add GrandpaName
                    If CurNodeObj.m_DirectParents.Exists(GrandpaName) Then
                        MsgBox "Warning : Node " & GrandpaName & " is repetitively linked to node " & CurNodeName & "."
                    End If
                Next GrandpaName
                NodesMetInPassing.Remove (PNodeName)
                RelativeTreeLevel = RelativeTreeLevel - 1
        Next PNodeName
        
        'Debug.Print "case 2 step 2 for " & CurNodeName
        '2. For each parent found in the recursive process,
        For Each PNodeName In ReturnedCollection
            If Not CurNodeObj.m_ParentNodesToCheck.Exists(PNodeName) Then
                'make sure the m_ParentNodesToCheck dictionary doesn't get confused
                'in case of repetitive links. Repetitive links are not errors, especially in the current context,
                'where "every node above you gets to be your parent"
                Set PNObj = m_ConnectMap(PNodeName)
                CurNodeObj.m_ParentNodesToCheck.Add Key:=PNodeName, Item:=0
                PNObj.m_ChildNodesToCheck.Add Key:=CurNodeName, Item:=0
            'Else
                
            End If
        Next PNodeName
    End If
    
    Set m_SetParentCollectionForNode = ReturnedCollection

End Function

