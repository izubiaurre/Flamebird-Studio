VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Zubiaurre:  izubiaurre@users.sourceforge.net (Imanol Zubiaurre)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

Option Explicit
Public Enum TrackerColumns
    tcCheckBox = &H1
    tcIcons = &H2
    tcSummary = &H4
    tcCategory = &H8
    tcModule = &H10
    tcAssignedTo = &H20
    tcCompleted = &H40
    tcCreated = &H80
    tcModified = &H100
    tcPriority = &H200
    tcDetailedDesc = &H400
    tcSubmittedBy = &H800
    tcDateclosing = &H1000
    tcGeneral = &H1FFF '&H3BF Or tcCompleted Or tcDATECLOSING Or tcSubmittedBy Xor tcPriority   ' All except Completed and detailed desc
End Enum

Private m_Col As Collection             ' Item collection of the tracker (main)
Private m_colCategory As Collection     ' Tracker categories

Private m_ID As Long
Private m_Name As String                ' Name of the tracker
Private m_Description As String
Private m_IconIndex As Long             ' Icon in the icon-list
Private m_VisibleColumns As TrackerColumns      ' One or more elements that specifies the visible columns
Private m_SortColumn As Long            ' Column to sort by
Private m_SortOrder As ECGSortOrderConstants
Private m_ColsArrangement(1 To 10)      ' Columns disposed in the grid
Private m_AutoExpandSelItems As Boolean ' If it's true, the items show its desc when selected

' Properties to be added
Public ColorItemsByPriority As Boolean
Public ColorOldItems As Boolean
Private m_OldItemsDays As Integer
'Private m_HiddenItemsDays As Integer

Private counter As Long ' Counter to asign IDs to the elementes of teh tracker
                        ' Always growing and it's the colecction key of the
                        ' cTracker. Starts in 1
                        
'-----------------------------------------------------------------------------
' PROPERTIES AND GENERAL METHOD IMPLEMENTATION
'-----------------------------------------------------------------------------
Friend Property Let Id(newID As Long)
    m_ID = newID
End Property
Public Property Get Id() As Long
    Id = m_ID
End Property

Public Property Get Key() As String
    Key = CStr(hex(m_ID))
End Property

Public Property Let name(newName As String)
    m_Name = newName
End Property
Public Property Get name() As String
    name = m_Name
End Property

Public Property Let description(newDesc As String)
    m_Description = newDesc
End Property
Public Property Get description() As String
    description = m_Description
End Property

Public Property Let IconIndex(newIcon As Long)
    m_IconIndex = newIcon
End Property
Public Property Get IconIndex() As Long
    IconIndex = m_IconIndex
End Property

Public Property Let VisibleColumns(newVC As TrackerColumns)
    m_VisibleColumns = newVC
End Property
Public Property Get VisibleColumns() As TrackerColumns
    VisibleColumns = m_VisibleColumns
End Property

Public Property Let SortColumn(newCol As Long)
    m_SortColumn = newCol
End Property
Public Property Get SortColumn() As Long
    SortColumn = m_SortColumn
End Property

Public Property Let SortOrder(newOrder As ECGSortOrderConstants)
    m_SortOrder = newOrder
End Property
Public Property Get SortOrder() As ECGSortOrderConstants
    SortOrder = m_SortOrder
End Property

' Keep atention with this function, cause doesn't check the given parameters
Public Sub SetColsArrangement(ParamArray iCols() As Variant)
    Dim i As Integer
    
    For i = 0 To UBound(iCols)
        m_ColsArrangement(i) = iCols(i)
    Next
End Sub

Public Property Let ShowDescription(newBool As Boolean)
    If newBool = True Then
        m_VisibleColumns = m_VisibleColumns Or tcDetailedDesc
    Else
        If Not (m_VisibleColumns And tcDetailedDesc) Then
            m_VisibleColumns = m_VisibleColumns Xor tcDetailedDesc
        End If
    End If
End Property
Public Property Get ShowDescription() As Boolean
    ShowDescription = IIf((m_VisibleColumns And tcDetailedDesc), True, False)
End Property

Public Property Let ColumnVisible(newColumn As TrackerColumns, newBool As Boolean)
    If newBool = True Then
        m_VisibleColumns = m_VisibleColumns Or newColumn
    Else
        If Not (m_VisibleColumns And newColumn) Then
            m_VisibleColumns = m_VisibleColumns Xor newColumn
        End If
    End If
End Property
Public Property Get ColumnVisible(newColumn As TrackerColumns) As Boolean
    ColumnVisible = IIf((m_VisibleColumns And newColumn), True, False)
End Property

Public Property Let AutoExpandSelItems(newExpand As Boolean)
    m_AutoExpandSelItems = newExpand
    If newExpand = True Then ShowDescription = True
End Property
Public Property Get AutoExpandSelItems() As Boolean
    ' If colum detailed desc is visible, return false
    'AutoExpandSelItems = IIf((m_VisibleColumns And tcDetailedDesc), False, m_AutoExpandSelItems)
    AutoExpandSelItems = m_AutoExpandSelItems
End Property

Public Property Let OldItemsDays(newDays As Integer)
    If newDays < 1 Then newDays = 1
    m_OldItemsDays = newDays
End Property
Public Property Get OldItemsDays() As Integer
    OldItemsDays = m_OldItemsDays
End Property

'CATEGORIES
Public Sub AddCategory(catName As String)
    m_colCategory.Add catName, catName
End Sub

Public Sub RemoveCategory(Index As Variant)
    m_colCategory.Remove (Index)
End Sub

Public Function CategoryCount() As Long
    CategoryCount = m_colCategory.count
End Function

Public Property Get CategoryCol() As Collection
    Set CategoryCol = m_colCategory
End Property

Public Sub CategoryClear()
    Dim p As Variant
    For Each p In m_colCategory
        m_colCategory.Remove (CStr(p))
    Next
End Sub


'-----------------------------------------------------------------------------
' IMPLEMENTATION OF THE COLECTION
'-----------------------------------------------------------------------------
' Add a element to the tracker from cTrackerItem,
' asigns the Id as key for the collection (Hex type)
Public Sub AddIndirect(ByVal TrackerItem As cTrackerItem)
    Dim it As cTrackerItem
    
    counter = counter + 1
    
    ' We need a copy of the obj cause ByVal doesn't prevent to change the given obj as argument
    Set it = TrackerItem.Copy
    it.Id = counter
    
    m_Col.Add it, hex(counter) ' Add the elemnt to the collection
    
    Set it = Nothing
End Sub

Public Property Get item(ByVal ItemIndex As Variant) As cTrackerItem
Attribute item.VB_UserMemId = 0
    Set item = m_Col.item(ItemIndex)
End Property

Public Property Get count() As Long
    count = m_Col.count
End Property

Public Sub Remove(ByVal ItemIndex As Variant)
    m_Col.Remove ItemIndex
End Sub

Public Sub Clear()
    Dim p As Variant
    For Each p In m_Col
        m_Col.Remove (hex(p.Id))
    Next
    counter = 0
End Sub

' To use For... Each...Next. Id de proc=-4, Needed hidden
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

'-----------------------------------------------------------------------------
' INITIALIZE AND TERMINATE
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Col = New Collection
    Set m_colCategory = New Collection
    m_IconIndex = -1            ' Without icon by default
    m_VisibleColumns = tcGeneral
    ColorItemsByPriority = True ' Priorities with colors by default
    m_AutoExpandSelItems = True
    ColorOldItems = True
    m_OldItemsDays = 30
    counter = 0                 ' Counter at 0
    Id = -1                     ' Tracker is not in any colection
End Sub

Private Sub Class_Terminate()
    Set m_Col = Nothing
    Set m_colCategory = Nothing
End Sub

' returns a copy-object
Public Function Copy() As cTracker
    Dim ret As cTracker
    Set ret = New cTracker
    
    ret.AutoExpandSelItems = Me.AutoExpandSelItems
    ret.ColorItemsByPriority = Me.ColorItemsByPriority
    ret.description = Me.description
    ret.IconIndex = Me.IconIndex
    ret.name = Me.name
    ret.SortColumn = Me.SortColumn
    ret.SortOrder = Me.SortOrder
    ret.VisibleColumns = Me.VisibleColumns
    ret.ColorOldItems = Me.ColorOldItems
    ret.OldItemsDays = Me.OldItemsDays
    
    ' Copy the categories
    Dim s As Variant
    For Each s In m_colCategory
        ret.AddCategory CStr(s)
    Next
    
    ' Copy the elements of the tracjer
    Dim it As cTrackerItem
    For Each it In Me
        ret.AddIndirect it
    Next
    
    Set Copy = ret
    Set ret = Nothing
End Function
