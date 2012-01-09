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
Private m_CallBackFunction As String    ' Function that is called when property is changed
Private m_CallingObject As Object       ' Object where is executed the callback function
Private m_Value As Variant
Private m_Editable As Boolean
Private m_Description As String
' Max and min properties are multipurpose
' for ptText mean maximum and minimum characters
' whereas for numbers mean limit values
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

Public Property Let Value(newValue As Variant)
    ' Check property type
    Select Case m_Type
    Case ptInteger ' Integer
        m_Value = newValue
    
    Case ptNumeric ' Numeric (can be decimal)
        m_Value = newValue ' change the property
    
    Case ptCombo ' ComboBox
        m_Value = CInt(newValue)
    
    Case Else
        m_Value = newValue
    End Select
    
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Editable(newEditable As Boolean)
    m_Editable = newEditable
End Property

Public Property Get Editable() As Boolean
    Editable = m_Editable
End Property

Public Property Let description(newDesc As String)
    m_Description = newDesc
End Property

Public Property Get description() As String
    description = m_Description
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
    ' Error if not a combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Function
    
    m_colOptions.Add OpName, OpName
    'm_Max = m_colOptions.Count - 1 ' Sets max element
End Function

Public Property Get OptionItem(ByVal ItemIndex As Variant) As String
    ' Error if not a combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Property
    OptionItem = m_colOptions.item(ItemIndex + 1)
End Property

Public Property Get OptionCount() As Integer
    ' Error if not a combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Property
    
    OptionCount = m_colOptions.count
End Property

Public Sub OptionRemove(ByVal ItemIndex As Variant)
    ' Error if not a combo
    If m_Type <> ptCombo Then Err.Raise vbObjectError + 513, , "The property's type is not PTCOMBO": Exit Sub
    
    m_colOptions.Remove ItemIndex
    'm_Max = m_colOptions.Count - 1 ' Sets max element
End Sub

' To use For...Each...Next.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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
