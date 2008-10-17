VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCatViewFolder"
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

Private m_Name As String
Private m_Ext As String

Public Property Get name() As String
    name = m_Name
End Property

Public Property Let name(newName As String)
    m_Name = newName
End Property

Public Property Get Extensions() As String
    Extensions = m_Ext
End Property
Public Property Let Extensions(newExt As String)
    m_Ext = newExt
End Property

Public Function ExtCount() As Integer
    Dim s() As String
    s = Split(m_Ext, "|", , vbTextCompare)
    ExtCount = UBound(s) + 1
End Function

Public Function Ext(Index As Integer) As String
    Dim s() As String
    s = Split(m_Ext, "|", , vbTextCompare)
    If Not Index > UBound(s) Then
        Ext = Trim(s(Index))
    End If
End Function
