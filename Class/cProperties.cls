VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Private m_Col As Collection

Public Function Add(ByVal Caption As String, Key As String, ByVal TypeOfProp As PropertyType, _
                    ByRef CallingObject As Object, ByVal CallBackFunction As String, ByVal Value As String, _
                    ByVal Editable As Boolean, Optional max As Integer, _
                    Optional min As Integer = 0, Optional CanBeEmpty As Boolean = False)
    
    On Error GoTo ErrHandle
    
    'Creamos la propiedad y establecemos sus propiedades
    Dim Prop As New cProperty
    Prop.name = Caption
    Prop.Key = Key
    Prop.TypeOfProp = TypeOfProp
    Set Prop.CallingObject = CallingObject
    Prop.CallBackFunction = CallBackFunction
    Prop.Value = Value
    Prop.Editable = Editable
    If Not IsMissing(max) And Not IsMissing(min) Then Prop.SetLimits max, min
    Prop.CanBeEmpty = CanBeEmpty
    
    m_Col.Add Prop, Prop.Key 'Agrega la propiedad a la colección
    
    Set Prop = Nothing
    Exit Function
    
ErrHandle:
    MsgBox Err.description
    Set Prop = Nothing
End Function

Public Property Get item(ByVal ItemIndex As Variant) As cProperty
Attribute item.VB_UserMemId = 0
    Set item = m_Col.item(ItemIndex)
End Property

Public Property Get count() As Integer
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
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'permite ser utilizado con For ...Each
    Set NewEnum = m_Col.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub
