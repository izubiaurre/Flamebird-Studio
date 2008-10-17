VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrackerCollection"
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
Private m_Col As Collection

Private counter As Long 'Contador para asignar IDs a los trackers de la coleccón

'-----------------------------------------------------------------------------
' IMPLEMENTACIÓN DE LA COLECCIÓN
'-----------------------------------------------------------------------------

'Añade un tracker a la colección a partir de los parámetros que se especifican
Public Sub Add(name As String, Optional description As String = "", _
                    Optional IconIndex As Long = -1, Optional VisibleColumns _
                    As TrackerColumns = tcGeneral, Optional SortColumn = 5, _
                    Optional SortOrder As ECGSortOrderConstants = CCLOrderDescending, _
                    Optional AutoExpandSelItems As Boolean = False, Optional ColorByPriority As Boolean = True, _
                    Optional ColorOldItems As Boolean = True, Optional OldItemsDays As Integer = 30)
                    
    
    Dim tr As New cTracker
    'Propiedades del tracker
    tr.name = name
    tr.IconIndex = IconIndex
    tr.VisibleColumns = VisibleColumns
    tr.description = description
    tr.SortColumn = SortColumn
    tr.SortOrder = SortOrder
    tr.AutoExpandSelItems = AutoExpandSelItems
    tr.ColorItemsByPriority = ColorByPriority
    tr.ColorOldItems = ColorOldItems
    tr.OldItemsDays = OldItemsDays
    
    counter = counter + 1 'incrementamos el contador
    tr.Id = counter
    m_Col.Add tr, hex(counter) 'Añade el tracker a la colección
    
    Set tr = Nothing
End Sub

'Añade un tracker a la colección a partir de un objeto cTracker
Public Sub AddIndirect(tracker As cTracker)
    'Añadimos el tracker
    With tracker
        Add .name, .description, .IconIndex, .VisibleColumns, .SortColumn, .SortOrder, _
            .AutoExpandSelItems, .ColorItemsByPriority, .ColorOldItems, .OldItemsDays
    End With
    
    'Copiamos las categorías
    Dim s As Variant
    For Each s In tracker.CategoryCol
        item(m_Col.count).AddCategory CStr(s)
    Next
    
    'Copiamos los items del tracker pasado al tracker que hemos añadido a
    'la colección
    Dim ti As cTrackerItem
    For Each ti In tracker
        item(m_Col.count).AddIndirect ti
    Next
End Sub

Public Property Get item(ByVal ItemIndex As Variant) As cTracker
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

'Permite usar For...Each...Next. Id proc=-4, Atribut=Oculto
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'permite ser utilizado con For ...Each
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Function IndexForName(ByVal sName As String) As Long
    Dim tr As cTracker, i As Long
    For Each tr In Me
        If LCase(tr.name) = LCase(sName) Then
            IndexForName = tr.Id
            Exit Function
        End If
    Next
    IndexForName = 0
End Function

Public Function KeyForName(ByVal sName As String) As String
    Dim tr As cTracker, i As Long
    For Each tr In Me
        If LCase(tr.name) = LCase(sName) Then
            KeyForName = tr.Key
            Exit Function
        End If
    Next
    KeyForName = ""
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

'Devuelve un objeto copia
Public Function Copy() As cTrackerCollection
    Dim ret As cTrackerCollection
    Set ret = New cTrackerCollection
    
    'Copiamos todos los trackers
    Dim tr As cTracker
    For Each tr In Me
        ret.AddIndirect tr.Copy
    Next
    
    Set Copy = ret
    Set ret = Nothing
End Function

'Copia el contenido del objeto en otro en otro
Public Sub CopyIn(colTrackers As cTrackerCollection)
    colTrackers.Clear
    'Copiamos todos los trackers
    Dim tr As cTracker
    For Each tr In Me
        colTrackers.AddIndirect tr.Copy
    Next
End Sub
