VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeveloper"
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

Option Explicit

Private m_Name As String
Private m_RealName As String
Private m_Mail As String
Private m_ID As Long

'-----------------------------------------------------------------------------
'IMPLEMENTACIÓN PROPIEDADES Y MÉTODOS GENERALES
'-----------------------------------------------------------------------------
Friend Property Let id(newID As Long)
    m_ID = newID
End Property
Public Property Get id() As Long
    id = m_ID
End Property

Public Property Get Key() As String
    Key = CStr(Hex(m_ID))
End Property

'Name es la propiedad por defecto
Public Property Let name(newName As String)
    m_Name = newName
End Property
Public Property Get name() As String
Attribute name.VB_UserMemId = 0
    name = m_Name
End Property

Public Property Let RealName(newRealName As String)
    m_RealName = newRealName
End Property
Public Property Get RealName() As String
    RealName = m_RealName
End Property

Public Property Let Mail(newMail As String)
    m_Mail = newMail
End Property
Public Property Get Mail() As String
    Mail = m_Mail
End Property

'-----------------------------------------------------------------------------
' INICIALIZACIÓN Y SALIDA
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    id = -1 'El developer no está en ninguna colección
End Sub
