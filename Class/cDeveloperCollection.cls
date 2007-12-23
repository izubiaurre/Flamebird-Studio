VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeveloperCollection"
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

Private m_Col As Collection
Private counter As Long 'Contador para asignar IDs a los developers de la coleccón
Private m_DefaultDev As String

Public Property Let defaultDev(DevName As String)
    If IDForName(DevName) > 0 Then
        m_DefaultDev = DevName
    Else
        m_DefaultDev = ""
    End If
End Property

Public Property Get defaultDev() As String
    defaultDev = m_DefaultDev
End Property

Public Sub Add(name As String, RealName As String, Mail As String)
    Dim dev As New cDeveloper
    
    dev.name = name
    dev.RealName = RealName
    dev.Mail = Mail
    
    counter = counter + 1 'incrementamos el contador
    dev.id = counter
    m_Col.Add dev, Hex(counter)
End Sub

Public Property Get item(ByVal ItemIndex As Variant) As cDeveloper
Attribute item.VB_UserMemId = 0
    If ItemIndex = "" Then 'Si es nulo seleccionamos el primer elemento
        Set item = m_Col.item(1)
    Else
        Set item = m_Col.item(ItemIndex)
    End If
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
        m_Col.Remove (p.Key)
    Next
    counter = 0
End Sub

Public Function IDForName(ByVal sName As String) As Long
    Dim dev As cDeveloper, i As Long
    For Each dev In Me
        If LCase(dev.name) = LCase(sName) Then
            IDForName = dev.id
            Exit Function
        End If
    Next
    IDForName = 0
End Function

'Permite usar For...Each...Next. Id proc=-4, Atribut=Oculto
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'permite ser utilizado con For ...Each
    Set NewEnum = m_Col.[_NewEnum]
End Function

'-----------------------------------------------------------------------------
' INICIALIZACIÓN Y SALIDA
'-----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Col = New Collection
    counter = 0
End Sub

Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub
