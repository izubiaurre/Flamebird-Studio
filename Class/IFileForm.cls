VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com(JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Izubiaurre: izubiaurre@users.sourceforge.net (Imanol Izubiaurre)
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

Public Enum EFileFormConstants
    FF_SOURCE
    FF_MAP
    FF_FPG
End Enum
Public Property Get IsDirty() As Boolean
End Property
Public Property Get filename() As String
End Property
Public Property Get FilePath() As String
End Property
Public Function Load(ByVal sFile As String) As Long
End Function
Public Function Save(ByVal sFile As String) As Long
End Function
'The letter W is because of prohibition of VB to call this functions Close and New
Public Function CloseW() As Long
End Function
Public Function NewW(ByVal iUntitledCount As Integer) As Long
End Function
Public Function Identify() As EFileFormConstants
End Function
Public Property Get AlreadySaved() As Boolean 'To know when the Save As dialog must be shown
End Property
Public Property Get Title() As String
End Property
Private Property Let IsDirty(ByVal newval As Boolean)
    m_IsDirty = newval
    'Put an * to the caption if dirty
    Caption = IIf(newval, "*", "") & IFileForm_Title
End Property



