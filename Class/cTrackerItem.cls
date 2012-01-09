VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrackerItem"
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
Public Enum PriorityLevel
    plHighest = 5
    plHigh = 4
    plMedium = 3
    plLow = 2
    plLowest = 1
End Enum

Private m_ID As Long                ' Unique ID automaticaly given when adding to the tracker
                                    ' -1 means that it has been not added to any tracker
Private m_Summary As String         ' Short description
Private m_DetailedDesc As String
Private m_Module As String          ' File that references the item
Private m_Category As String
Private m_Priority As PriorityLevel
Private m_Completed As Integer      ' Completed percent
Private m_AssignedTo As String
Private m_SubmittedBy As String
Private m_DateCreated As Date
Private m_DateModified As Date
Private m_DateClosing As Date
Private m_Closed  As Boolean        ' Shows if the item is closed or not
Private m_Hidden As Boolean         ' Shows if the item is deleted or not
Private m_Locked As Boolean         ' Shows if the item is blocked or not
'-----------------------------------------------------------------------------
' PROPERTIES IMPLEMENTATION
'-----------------------------------------------------------------------------
Friend Property Let Id(newID As Long)
    m_ID = newID
End Property
Public Property Get Id() As Long
    Id = m_ID
End Property

Public Property Let Summary(newSummary As String)
    m_Summary = newSummary
End Property
Public Property Get Summary() As String
    Summary = m_Summary
End Property

Public Property Let DetailedDesc(newDetailedDesc As String)
    m_DetailedDesc = newDetailedDesc
End Property
Public Property Get DetailedDesc() As String
    DetailedDesc = m_DetailedDesc
End Property

Public Property Let Category(newCategory As String)
    m_Category = newCategory
End Property
Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Let module(newModule As String)
    m_Module = newModule
End Property
Public Property Get module() As String
    module = m_Module
End Property

Public Property Let Priority(newPriority As PriorityLevel)
    m_Priority = newPriority
End Property
Public Property Get Priority() As PriorityLevel
    Priority = m_Priority
End Property

Public Property Let Completed(newCompleted As Integer)
    If newCompleted < 0 Then newCompleted = 0
    If newCompleted > 100 Then newCompleted = 100
    m_Completed = newCompleted
End Property
Public Property Get Completed() As Integer
    Completed = m_Completed
End Property

Public Property Let AssignedTo(newAssignedTo As String)
    m_AssignedTo = newAssignedTo
End Property
Public Property Get AssignedTo() As String
    AssignedTo = m_AssignedTo
End Property

Public Property Let SubmittedBy(newSubmitted As String)
     m_SubmittedBy = newSubmitted
End Property
Public Property Get SubmittedBy() As String
     SubmittedBy = m_SubmittedBy
End Property

Public Property Let DateCreated(newDate As Date)
    m_DateCreated = newDate
End Property
Public Property Get DateCreated() As Date
    DateCreated = m_DateCreated
End Property

Public Property Let DateModified(newDate As Date)
    m_DateModified = newDate
End Property
Public Property Get DateModified() As Date
    DateModified = m_DateModified
End Property

Public Property Let DateClosing(newDate As Date)
    m_DateClosing = newDate
End Property
Public Property Get DateClosing() As Date
    DateClosing = m_DateClosing
End Property

Public Property Let Closed(newClosed As Boolean)
    m_Closed = newClosed
End Property

Public Property Get Closed() As Boolean
    Closed = m_Closed
End Property

Public Property Let Hidden(newHidden As Boolean)
    m_Hidden = newHidden
End Property

Public Property Get Hidden() As Boolean
    Hidden = m_Hidden
End Property

Public Property Let Locked(newLocked As Boolean)
    m_Locked = newLocked
End Property

Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

'-----------------------------------------------------------------------------
' INITIALIZE AND TERMINATE
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    m_ID = -1       ' It's not in any tracker
    m_Completed = 0
    m_Closed = False
    m_Priority = plLowest
   ' m_AssignedTo = New cDeveloper
   ' m_SubmittedBy = New cDeveloper
End Sub

Private Sub Class_Terminate()
   ' Set m_AssignedTo = Nothing
   ' Set m_SubmittedBy = Nothing
End Sub

' Returns the copy of the object
Public Function Copy() As cTrackerItem
    Dim ret As cTrackerItem
    Set ret = New cTrackerItem
    
    ret.AssignedTo = Me.AssignedTo
    ret.Category = Me.Category
    ret.Completed = Me.Completed
    ret.DateClosing = Me.DateClosing
    ret.DateCreated = Me.DateCreated
    ret.DateModified = Me.DateModified
    ret.DetailedDesc = Me.DetailedDesc
    ret.Id = Me.Id
    ret.module = Me.module
    ret.Priority = Me.Priority
    ret.SubmittedBy = Me.SubmittedBy
    ret.Summary = Me.Summary
    ret.Closed = Me.Closed
    ret.Hidden = Me.Hidden
    ret.Locked = Me.Locked
    
    Set Copy = ret
    Set ret = Nothing
End Function
