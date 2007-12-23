VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Enum PropertyType
    ptText = 1
    ptInteger = 2
    ptNumeric = 3
    ptCombo = 4
    ptLink = 5
End Enum

Private m_Name As String
Private m_Key As String
Private m_Type As PropertyType
Private m_CallBackFunction As String 'Función a la que se llama cuando se modifica una propiedad
Private m_CallingObject As Object 'Objeto donde se ejecuta la función callback
Private m_Value As Variant
Private m_Editable As Boolean
Private m_Description As String
'Las propiedades max y min son polivalentes,
'para ptText significa máximo y mínimo de caracteres
'mientras que para números significa valores límites
Private m_CanBeEmpty As Boolean
Private m_IsLimited As Boolean
Private m_Max As Integer
Private m_Min As Integer

Private m_colOptions As Collection

Public Property Let name(ByVal newName As String)
    m_Name = newName
End Property

Public Property Get name() As String
    name = m_Name
End Property

Public Property Let Key(ByVal newKey As String)
    m_Key = newKey
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Set CallingObject(newObject As Object)
    Set m_CallingObject = newObject
End Property

Public Property Get CallingObject() As Object
    Set CallingObject = m_CallingObject
End Property

Public Property Let TypeOfProp(ByVal newType As PropertyType)
    m_Type = newType
    Set m_colOptions = Nothing
    If TypeOfProp = ptCombo Then
        Set m_colOptions = New Collection
    End If
    
    m_Max = 0
    m_Min = 0
    m_IsLimited = False
End Property

Public Property Get TypeOfProp() As PropertyType
    TypeOfProp = m_Type
End Property

Public Property Let CallBackFunction(newCallBackFunc As String)
    m_CallBackFunction = newCallBackFunc
End Property

Public Property Get CallBackFunction() As String
    CallBackFunction = m_CallBackFunction
End Property

Public Property Let value(newValue As Variant)
    'Comprobamos el tipo de propiedad
    Select Case m_Type
    Case ptInteger 'Entero
        m_Value = newValue
    
    Case ptNumeric 'Número cualquiera (admite decimales)
        m_Value = newValue 'cambiamos la propiedad
    
    Case ptCombo 'ComboBox
        m_Value = CInt(newValue)
    
    Case Else
        m_Value = newValue
    End Select
    
End Property

Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
    value = m_Value
End Property

Public Property Let Editable(newEditable As Boolean)
    m_Editable = newEditable
End Property

Public Property Get Editable() As Boolean
    Editable = m_Editable
End Property

Public Property Let Description(newDesc As String)
    m_Description = newDesc
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let IsLimited(newLimited As Boolean)
    If Not m_Type = ptCombo Then m_IsLimited = newLimited
End Property

Public Property Get IsLimited() As Boolean
    IsLimited = m_IsLimited
End Property
Public Property Let CanBeEmpty(newEmpty As Boolean)
    If Not m_Type = ptCombo Then m_CanBeEmpty = newEmpty
End Property

Public Property Get CanBeEmpty() As Boolean
    CanBeEmpty = m_CanBeEmpty
End Property
Public Function SetLimits(ValMax As Integer, ValMin As Integer)
    m_IsLimited = True
    If Not m_Type = ptCombo Then
        m_Max = ValMax
        m_Min = ValMin
    End If
End Function
Public Property Let max(newMax As Integer)
    If Not m_Type = ptCombo Then
        m_Max = newMax
    End If
End Property

Public Property Get max() As Integer
    max = m_Max
End Property

Public Property Let min(newMin As Integer)
    If Not m_Type = ptCombo Then
        m_Min = newMin
    End If
End Property

Public Property Get min() As Integer
    min = m_Min
End Property

Public Function AddOption(OpName As String)
    'Error si no es un combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Function
    
    m_colOptions.Add OpName, OpName
    'm_Max = m_colOptions.Count - 1 'Establece el máximo elemento
End Function

Public Property Get OptionItem(ByVal ItemIndex As Variant) As String
    'Error si no es un combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Property
    OptionItem = m_colOptions.item(ItemIndex + 1)
End Property

Public Property Get OptionCount() As Integer
    'Error si no es un combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Property
    
    OptionCount = m_colOptions.count
End Property

Public Sub OptionRemove(ByVal ItemIndex As Variant)
    'Error si no es un combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Sub
    
    m_colOptions.Remove ItemIndex
    'm_Max = m_colOptions.Count - 1 'Establece el máximo elemento
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'permite ser utilizado con For ...Each
    Set NewEnum = m_colOptions.[_NewEnum]
End Function

Private Sub Class_Initialize()
    m_Max = 0
    m_Min = 0
    m_IsLimited = False
    m_CanBeEmpty = False
    m_Description = "Description not available"
End Sub

Private Sub Class_Terminate()
    Set m_colOptions = Nothing
End Sub
