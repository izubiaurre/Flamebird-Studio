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
    tcGeneral = &H1FFF '&H3BF Or tcCompleted Or tcDATECLOSING Or tcSubmittedBy Xor tcPriority   'Todas menos Completed y detailed desc
End Enum

Private m_Col As Collection 'Colección de elementos del tracker (principal)
Private m_colCategory As Collection 'Categorias del tracker

Private m_ID As Long
Private m_Name As String 'Nombre del tracker
Private m_Description As String
Private m_IconIndex As Long 'Icono en la lista de iconos
Private m_VisibleColumns As TrackerColumns  'Uno o mas elementos que especifican las columnas visibles
Private m_SortColumn As Long 'Columna por la cual se ordena
Private m_SortOrder As ECGSortOrderConstants
Private m_ColsArrangement(1 To 10) 'Disposición de las columnas en el grid
Private m_AutoExpandSelItems As Boolean 'Si es true, los elementos muestran su descripción al seleccionarse

'Propiedades a ser añadidas
Public ColorItemsByPriority As Boolean
Public ColorOldItems As Boolean
Private m_OldItemsDays As Integer
'Private m_HiddenItemsDays As Integer

Private counter As Long 'Contador para asignar IDs a los elementos del tracker
                        'Es siempre creciente y sirve como key para una colección
                        'cTracker. Empieza en 1
                        
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

Public Property Let name(newName As String)
    m_Name = newName
End Property
Public Property Get name() As String
    name = m_Name
End Property

Public Property Let Description(newDesc As String)
    m_Description = newDesc
End Property
Public Property Get Description() As String
    Description = m_Description
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

'Cuidado con esta función, pues no comprueba el tipo de argumentos pasados
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
    'Si la columna detailed description es visible, devolvemos false
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

'CATEGORIAS
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
' IMPLEMENTACIÓN DE LA COLECCIÓN
'-----------------------------------------------------------------------------
'Añade un elemento al tracker desde un objeto cTrackerItem, asignado el id
'Establece como clave para la colección un valor hexadecimal guardado como
'cadena
Public Sub AddIndirect(ByVal TrackerItem As cTrackerItem)
    Dim it As cTrackerItem
    
    counter = counter + 1 'incrementamos el contador
    
    'Necesitamos una copia del objeto porque ByVal no nos previene de modificar el
    'objeto pasado como argumento
    Set it = TrackerItem.Copy
    it.id = counter
    
    m_Col.Add it, Hex(counter) 'Agregamos el elemento a la colección
    
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
        m_Col.Remove (Hex(p.id))
    Next
    counter = 0
End Sub

'Permite usar For... Each...Next. Id de procedimiento=-4, Necesario oculto
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
    Set m_colCategory = New Collection
    m_IconIndex = -1 'Sin icono por defecto
    m_VisibleColumns = tcGeneral
    ColorItemsByPriority = True 'Por defecto se usan colores para la prioridad
    m_AutoExpandSelItems = True
    ColorOldItems = True
    m_OldItemsDays = 30
    counter = 0 'Contador a 0
    id = -1 'El tracker no está en ninguna colección
End Sub

Private Sub Class_Terminate()
    Set m_Col = Nothing
    Set m_colCategory = Nothing
End Sub

'Devuelve un objeto copia
Public Function Copy() As cTracker
    Dim ret As cTracker
    Set ret = New cTracker
    
    ret.AutoExpandSelItems = Me.AutoExpandSelItems
    ret.ColorItemsByPriority = Me.ColorItemsByPriority
    ret.Description = Me.Description
    ret.IconIndex = Me.IconIndex
    ret.name = Me.name
    ret.SortColumn = Me.SortColumn
    ret.SortOrder = Me.SortOrder
    ret.VisibleColumns = Me.VisibleColumns
    ret.ColorOldItems = Me.ColorOldItems
    ret.OldItemsDays = Me.OldItemsDays
    
    'Copiamos las categorías
    Dim s As Variant
    For Each s In m_colCategory
        ret.AddCategory CStr(s)
    Next
    
    'Copiamos los elementos del tracker
    Dim it As cTrackerItem
    For Each it In Me
        ret.AddIndirect it
    Next
    
    Set Copy = ret
    Set ret = Nothing
End Function
