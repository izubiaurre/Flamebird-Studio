VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Begin VB.Form frmTodoList 
   Caption         =   "Fire Tracker"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "frmToDoList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   7050
   WindowState     =   2  'Maximized
   Begin VB.Frame grbNoTrackers 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label lblLinkTM 
         AutoSize        =   -1  'True
         Caption         =   "Tracker Manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   2280
         MouseIcon       =   "frmToDoList.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label lblNT2 
         AutoSize        =   -1  'True
         Caption         =   "To add new trackers go to the "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label lblNT1 
         AutoSize        =   -1  'True
         Caption         =   "There are no available trackers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   2565
      End
   End
   Begin vbalIml6.vbalImageList ilstMenus 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   3444
      Images          =   "frmToDoList.frx":06F6
      Version         =   131072
      KeyCount        =   3
      Keys            =   "ÿÿ"
   End
   Begin vbalIml6.vbalImageList ilstTabs 
      Left            =   4800
      Top             =   1680
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   115948
      Images          =   "frmToDoList.frx":148A
      Version         =   131072
      KeyCount        =   101
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdTracker 
      Height          =   2295
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4048
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      NoVerticalGridLines=   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14074812
      GridLineColor   =   12434877
      GridFillLineColor=   14074812
      HighlightBackColor=   15523803
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      SelectionAlphaBlend=   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin vbalDTab6.vbalDTabControl tabTracker 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4101
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
   End
   Begin vbalIml6.vbalImageList ilstGrid 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   10332
      Images          =   "frmToDoList.frx":1D996
      Version         =   131072
      KeyCount        =   9
      Keys            =   "ÿÿÿÿÿÿÿÿ"
   End
End
Attribute VB_Name = "frmTodoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'Constantes de color
Private Const CLR_GRIDTEXT = &H9C3000
Private Const CLR_OLDDATE = &H1B11CC
Private Const CLR_GRIDLINES = &HC0C0C0 '&HBDBDBD
Private Const CLR_BACKGRID = &HF0F0EC '&HD6C3BC

'Constantes de posición de columnas
Private Const COL_CHECKBOX = 1
Private Const COL_ICONS = 2
Private Const COL_SUMMARY = 3
Private Const COL_CATEGORY = 4
Private Const COL_MODULE = 5
Private Const COL_ASSIGNEDTO = 6
Private Const COL_COMPLETED = 7
Private Const COL_PRIORITY = 8
Private Const COL_DATECREATED = 9
Private Const COL_DATEMODIFIED = 10
Private Const COL_DATECLOSING = 11
Private Const COL_SUBMITTEDBY = 12
Private Const COL_DETAILEDDESC = 13

'Constantes de iconos
'Iconos del grid
'Headers
Private Const ICD_CHECKHEADER = 0
Private Const ICD_SORTASCENDING = 3
Private Const ICD_SORTDESCENDING = 4
Private Const ICD_PRIORITY = 5
Private Const ICD_STATUS = 6
'Celdas
Private Const ICD_UNCHECK = 1
Private Const ICD_CHECK = ICD_UNCHECK + 1
Private Const ICD_LOCKED = 7
Private Const ICD_HIDDEN = ICD_LOCKED + 1

'Iconos de los tabs
Private Const ICD_BUGS = 0
Private Const ICD_RFE = 1
Private Const ICD_QUESTIONS = 2

Public colTrackers As cTrackerCollection
Public AT As cTracker 'Active tracker
Public ai As cTrackerItem 'Active item

Private bNoTrackers As Boolean 'Store if there isn't any tracker
Private devcol As cDeveloperCollection

Private m_ShowTabs As Boolean 'Ver la barra de lengüetas
Private m_ShowHiddenItems As Boolean 'Ver elementos ocultos

Implements IGridCellOwnerDraw
'Implements ITDockMoveEvents



Private Property Let NoTrackers(newBool As Boolean)
    bNoTrackers = newBool
    If newBool = True Then 'No trackers
        'Disable some controls and Show the grbNotrackers
        grdTracker.Visible = False
        tabTracker.Visible = False
        grbNoTrackers.ZOrder 0 'Move to the foreground
        grbNoTrackers.Visible = True
    Else
        'Enable some controls and hide the grdbnotrackers
        grdTracker.Visible = True
        tabTracker.Visible = ShowTabs
        grbNoTrackers.ZOrder 1 'Move to the background
        grbNoTrackers.Visible = False
    End If
End Property

Private Property Get NoTrackers() As Boolean
    NoTrackers = bNoTrackers
End Property

Private Property Let ShowTabs(newShow As Boolean)
    m_ShowTabs = newShow
    tabTracker.Visible = newShow
    tabTracker_Resize
End Property

Public Property Get ShowTabs() As Boolean
    ShowTabs = m_ShowTabs
End Property

Private Sub NewItem()
    Set ai = New cTrackerItem
    ai.SubmittedBy = devcol.defaultDev
    ai.DateCreated = Date
    ai.Summary = "New"
    AT.AddIndirect ai
    Set ai = AT(AT.count)
    
    frmTrackerItem.bIsNew = True
    frmTrackerItem.Show 1
    SelectTracker AT
End Sub

Private Sub EditItem()
    Set ai = AT(Hex(grdTracker.RowItemData(grdTracker.SelectedRow)))
    frmTrackerItem.Show 1
    SelectTracker AT 'Refresh
End Sub

Private Sub DeleteItem()
    With grdTracker
        If MsgBox("Sure to delete the item '" & AT(Hex(.RowItemData(.SelectedRow))).Summary _
                & "' ? (This action cannot be undone)", vbQuestion + vbYesNo) = vbYes Then
            'Delete the itme
            AT.Remove (Hex(AT(Hex(.RowItemData(.SelectedRow))).id))
            
            SelectTracker AT 'Refresh
        End If
    End With
End Sub

'Cuenta el número de elementos ocultos del tracker activo
Private Function CountHiddenItems() As Integer
    Dim counter As Integer
    Dim ti As cTrackerItem
    counter = 0
    For Each ti In AT
        If ti.Hidden = True Then counter = counter + 1
    Next
    CountHiddenItems = counter
End Function



'Dibuja las filas agrupadas de modo personalizado
Private Sub drawGroupRow(cell As cGridCell, ByVal lHdc As Long, ByVal lLeft As Long, _
      ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)

    Dim hFont As Long
    Dim hFontOld As Long
    Dim tr As RECT
    Dim tBR As RECT
   
    tr.Left = lLeft
    tr.Top = lTop
    tr.Right = lRight
    tr.Bottom = lBottom
   
    LSet tBR = tr
    tBR.Top = tBR.Bottom - 5
    tBR.Bottom = tBR.Bottom - 2
   ' If (cell.Selected) Then
        GradientFillRect lHdc, tBR, vbHighlight, vbWindowBackground, GRADIENT_FILL_RECT_H
  '  Else
  '      GradientFillRect lHDC, tBR, vbButtonShadow, vbWindowBackground, GRADIENT_FILL_RECT_H
  '  End If
    
    Dim fnt As New StdFont
    fnt.Bold = True
    fnt.name = "Tahoma"
    Dim m As IFont
    Set m = fnt
    
    hFont = m.hFont
    hFontOld = SelectObject(lHdc, hFont)
    tr.Bottom = tr.Bottom - 3
    DrawTextA lHdc, " " & ": " & cell.text, -1, tr, cell.TextAlign
    SelectObject lHdc, hFontOld

End Sub

'Dibuja la barra de progreso
Private Sub drawProgressCell(cell As cGridCell, ByVal lHdc As Long, _
      ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, _
      ByVal lBottom As Long)
Dim hBr As Long
Dim tr As RECT
Dim tProgR As RECT

   tr.Left = lLeft + 2
   tr.Top = lTop + 2
   tr.Right = lRight - 2
   tr.Bottom = lTop + grdTracker.DefaultRowHeight - 2

   'Dibujamos la barra de progreso
   LSet tProgR = tr
   tProgR.Right = tProgR.Left + (tProgR.Right - tProgR.Left) * cell.text * 1 / 100
   GradientFillRect lHdc, tProgR, RGB(234, 94, 45), RGB(238, 164, 36), GRADIENT_FILL_RECT_H
   
   'Escribimos el texto encima de la barra
   DrawTextA lHdc, format(CInt(cell.text) / 100, "0%"), -1, tr, cell.TextAlign

   'Creamos el contorno
   hBr = CreateSolidBrush(&H0&)
   FrameRect lHdc, tr, hBr
   DeleteObject hBr
End Sub

Public Sub RefreshTabs()
    Dim tr As cTracker
    Dim i As Integer, selTabID As String
    
    'Store the ID of the selected tab
    If tabTracker.Tabs.count > 0 Then
        selTabID = tabTracker.SelectedTab.ItemData
    Else
        selTabID = 0
    End If
    
    'Add the tabs
    tabTracker.Tabs.Clear
    For Each tr In colTrackers
        tabTracker.Tabs.Add tr.name, , tr.name, tr.IconIndex
        tabTracker.Tabs.item(tabTracker.Tabs.count).ItemData = tr.id
    Next
    
    'Search for the selected tab
    For i = 1 To tabTracker.Tabs.count
        If tabTracker.Tabs.item(i).ItemData = selTabID Then
            tabTracker.Tabs.item(i).Selected = True
            Exit For
        End If
    Next
    
    If colTrackers.count = 0 Then
        NoTrackers = True
    Else
        SelectTracker colTrackers(colTrackers.KeyForName(tabTracker.SelectedTab.Caption))
        NoTrackers = False
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    tabTracker.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    grbNoTrackers.Move 0, 0, Me.ScaleWidth, ScaleHeight
    'Center the labels of the notracker group box
    With grbNoTrackers
        lblNT2.Move (.Width - lblNT2.Width - lblLinkTM.Width) / 2, (.Height - lblNT2.Height) / 2
        lblLinkTM.Move lblNT2.Left + lblNT2.Width, lblNT2.Top
        lblNT1.Move (.Width - lblNT1.Width) / 2, lblNT2.Top - lblNT2.Height - 100
    End With
End Sub

Private Sub grdTracker_GotFocus()
    If grdTracker.SelectedRow > 0 Then grdTracker_SelectionChange grdTracker.SelectedRow, 1
End Sub

Private Sub grdTracker_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If grdTracker.SelectedRow > 0 Then DeleteItem
    End If
End Sub

Private Sub grdTracker_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    bCancel = True
    'Hemos pinchado "marcar como cerrado"
    If lCol = COL_CHECKBOX Then
       AT(Hex(grdTracker.RowItemData(lRow))).Closed = Not AT(Hex(grdTracker.RowItemData(lRow))).Closed
       SelectTracker AT
    Else
       EditItem
    End If
End Sub

Private Sub IGridCellOwnerDraw_Draw( _
      cell As cGridCell, _
      ByVal lHdc As Long, _
      ByVal eDrawStage As ECGDrawStage, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long, _
      bSkipDefault As Boolean _
   )
   If (eDrawStage = ecgBeforeIconAndText) Then
      If (cell.Column = COL_COMPLETED) Then
         drawProgressCell cell, lHdc, lLeft, lTop, lRight, lBottom
         bSkipDefault = True
      End If
      If grdTracker.RowIsGroup(cell.Row) Then
         drawGroupRow cell, lHdc, lLeft, lTop, lRight, lBottom
         bSkipDefault = True
      End If
   End If
End Sub

'Devuelve la altura necesaria para poder contener el texto de una celda completa
Private Function EvaluateTextHeight(lRow As Long, lCol As Long) As Long
    Dim i As Long, lWidth As Long, lHeight As Long
    Dim r As RECT
    
    If lRow <= 0 Or lCol <= 0 Then Exit Function 'Celda no válida
    
    lWidth = 0
    With grdTracker
        'Tomamos la suma de los anchos de las columnas VISIBLES y NO AGRUPADAS
        For i = .RowTextStartColumn To .Columns - 1
            If (AT.ColumnVisible(.ColumnTag(i))) And Not .ColumnIsGrouped(i) Then
                lWidth = lWidth + .ColumnWidth(i)
            End If
        Next
        
        SetRect r, 0, 0, lWidth, 0 'Ancho del rectángulo, el alto se calcula
        InflateRect r, -2, 0 'Necesario para que coincida con la región de la celda
        'Calculamos la altura necesaria para escribir el texto simulando una escritura sobre el rectángulo
        lHeight = DrawTextA(GetDC(.hwnd), .CellText(lRow, lCol) & vbNullChar, -1, r, DT_WORDBREAK Or DT_CALCRECT)
        
    End With
    
    EvaluateTextHeight = lHeight + grdTracker.DefaultRowHeight + 4
End Function

'Busca la fila que se corresponde con el id especificado (en el itemData)
'Si no la encuentra, devuelve -1
Private Function RowIndexByID(id As Long) As Long
    Dim i As Integer
    With grdTracker
        For i = 1 To .Rows
            If .RowItemData(i) = id Then RowIndexByID = i: Exit Function
        Next
    End With
    RowIndexByID = -1
End Function

'Ordena los elementos del tracker activo en el grid según la columna que se le indique
Private Sub Sort(ByVal lCol As Long, Optional order As ECGSortOrderConstants = CCLOrderAscending)
    Dim i As Integer, colIcon As Long
    
    With grdTracker.SortObject
        
        'Borramos el icono de la antigua sortcolumn
        If .count > 0 Then
            colIcon = grdTracker.ColumnImage(.SortColumn(1))
            If colIcon = ICD_SORTASCENDING Or colIcon = ICD_SORTDESCENDING Then grdTracker.ColumnImage(.SortColumn(1)) = -1
        End If
        
        .Clear
        .SortColumn(1) = lCol
        .SortType(1) = grdTracker.ColumnSortType(lCol)
        .SortOrder(1) = order

        'Ponemos el Icono en lCol
        If grdTracker.ColumnImage(lCol) = -1 Then 'La columna no tiene icono propio
            grdTracker.ColumnImage(lCol) = ICD_SORTASCENDING + (order - 1)
            grdTracker.ColumnImageOnRight(lCol) = True
        End If
        
    End With
    grdTracker.Sort
End Sub

'Establece las columnas del grid
Private Sub CreateGrid()
    With grdTracker
        .ImageList = ilstGrid 'Lista de iconos
        
        .Redraw = False 'Para mayor velocidad
        .GridLineColor = CLR_GRIDLINES
        .BackColor = CLR_BACKGRID
        .GridFillLineColor = CLR_BACKGRID
        .StretchLastColumnToFit = False
        .Editable = True
        .HighlightBackColor = &H575283

        'Ponemos todas las columnas y definimos sus características
        'En el ColumnTag guardamos el número que se corresponde con la enumeración Tracker Columns
        .AddColumn "CheckBox", , , ICD_CHECKHEADER, 25, , True, , True, , , CCLSortIcon
        .ColumnTag(COL_CHECKBOX) = tcCheckBox
        
        .AddColumn "Status", , , ICD_STATUS, 35, , True, , True, , , CCLSortExtraIcon
        .ColumnTag(COL_ICONS) = tcIcons
        
        .AddColumn "Summary", "Summary", , , 250
        .ColumnTag(COL_SUMMARY) = tcSummary
        
        .AddColumn "Category", "Category"
        .ColumnTag(COL_CATEGORY) = tcCategory
        
        .AddColumn "Module", "Module"
        .ColumnTag(COL_MODULE) = tcModule
        
        .AddColumn "Assigned To", "Assigned To", , , 100
        .ColumnTag(COL_ASSIGNEDTO) = tcAssignedTo
        
        .AddColumn "Progress", "Progress", ecgHdrTextALignCentre, , 100, eSortType:=CCLSortNumeric
        .ColumnTag(COL_COMPLETED) = tcCompleted
        
        .AddColumn "Priority", , ecgHdrTextALignCentre, ICD_PRIORITY, 25, , True, eSortType:=CCLSortNumeric
        .ColumnTag(COL_PRIORITY) = tcPriority
        
        .AddColumn "Created", "Created", ecgHdrTextALignRight, , 70, sFmtString:="Short Date", eSortType:=CCLSortDate
        .ColumnTag(COL_DATECREATED) = tcCreated
        
        .AddColumn "Modified", "Modified", ecgHdrTextALignRight, , 70, sFmtString:="Short Date", eSortType:=CCLSortDate
        .ColumnTag(COL_DATEMODIFIED) = tcModified
        
        .AddColumn "Closed", "Closed", ecgHdrTextALignRight, , 70, sFmtString:="Short Date", eSortType:=CCLSortDate
        .ColumnTag(COL_DATECLOSING) = tcDateclosing
        
        .AddColumn "Submitted by", "Submitted By", , , 100
        .ColumnTag(COL_SUBMITTEDBY) = tcSubmittedBy
        
        .AddColumn "Detailed Description", "Detailed Description", , , 96 + 256 + 96 + 96, , , , , , True
        .ColumnTag(COL_DETAILEDDESC) = tcDetailedDesc
        
        .RowTextStartColumn = COL_SUMMARY
                
        .SetHeaders
        
        .Redraw = True
    End With
End Sub

'Establece el tracker activo y determina las columnas del grid que serán visibles
Private Sub SelectTracker(tracker As cTracker)
    Dim i As Long
    Dim ScrollPosX As Long, ScrollPosY As Long 'Pos Scrollbars
    Dim SelItemID As Long 'Elemento seleccionado
    Dim selRow As Long
    
    With grdTracker
        .Redraw = False
           
        If AT.name = tracker.name Then 'Si es el mismo tracker guardamos datos
            ScrollPosX = .ScrollOffsetX
            ScrollPosY = .ScrollOffsetY
            If .SelectedRow > 0 Then
                SelItemID = .RowItemData(.SelectedRow)
            Else
                SelItemID = -1
            End If
        Else
            ScrollPosX = 0: ScrollPosY = 0: SelItemID = -1
        End If
    
        'Determinamos las columnas que se verán. La propiedad tag de la columna guarda
        'el valor que corresponde a la columna según la enumeración TrackerColumns
        For i = 1 To .Columns
            .ColumnVisible(i) = (tracker.ColumnVisible(CInt(.ColumnTag(i))))
        Next
        
        Set AT = tracker
        FillGrid AT
        
        Sort AT.SortColumn, AT.SortOrder 'Ordena
        
        'Restauramos la posición del scrollbar y del elemento seleccionado
        .ScrollOffsetX = ScrollPosX
        .ScrollOffsetY = ScrollPosY
        selRow = RowIndexByID(SelItemID)
        If selRow > -1 Then .SelectedRow = selRow
        
        .Redraw = True
    End With
End Sub

'Rellena el grid con los elementos de un tracker (el activo por defecto)
Private Sub FillGrid(tracker As cTracker)
    Dim i As Long
    Dim it As cTrackerItem
      
    grdTracker.Clear
    
    Dim boldfnt As New StdFont 'Fuente en negrita
    boldfnt.Bold = True
    Dim fnt As New StdFont
    
    'Para cada elemento añadimos una fila y escribimos en las celdas
    For Each it In tracker
    With grdTracker
        'Si el elemento está cerradola fuente se pone tachada
        fnt.Strikethrough = it.Closed
        boldfnt.Strikethrough = it.Closed
    
        .AddRow , it.id 'El id del elemento se guarda en el ItemData
        
        .CellDetails .Rows, COL_CHECKBOX, lIconIndex:=IIf(it.Closed, ICD_CHECK, ICD_UNCHECK), lIndent:=3 'Checkbox
        .CellIcon(.Rows, COL_ICONS) = IIf(it.Locked, ICD_LOCKED, -1) 'Bloqueado
        .CellExtraIcon(.Rows, COL_ICONS) = IIf(it.Hidden, ICD_HIDDEN, -1) 'Oculto
        .CellDetails .Rows, COL_SUMMARY, it.Summary, oFont:=boldfnt 'Sumary
        .CellDetails .Rows, COL_CATEGORY, it.Category, oFont:=fnt 'Category
        .CellDetails .Rows, COL_MODULE, it.module, oFont:=fnt 'Module
        .CellDetails .Rows, COL_ASSIGNEDTO, it.AssignedTo, oFont:=fnt 'AssignedTo
        .CellDetails .Rows, COL_DATECREATED, it.DateCreated, DT_RIGHT, oFont:=fnt 'Created
        .CellDetails .Rows, COL_DATEMODIFIED, format(it.DateModified, "Short Date"), DT_RIGHT, oFont:=fnt 'Modified
        .CellDetails .Rows, COL_PRIORITY, it.Priority, DT_CENTER, oFont:=fnt 'Priority
        .CellDetails .Rows, COL_COMPLETED, it.Completed, DT_CENTER, oFont:=fnt
        .CellDetails .Rows, COL_DATECLOSING, IIf(it.Closed, _
                            format(it.DateClosing, "Short Date"), "not closed"), DT_RIGHT, oFont:=fnt 'Date Closening
        .CellDetails .Rows, COL_SUBMITTEDBY, it.SubmittedBy, oFont:=fnt 'Submitted by
        .CellDetails .Rows, COL_DETAILEDDESC, it.DetailedDesc, DT_WORDBREAK, , , RGB(0, 0, 0), oFont:=fnt    'Detailded Desc
        
        'Determinamos si se debe o no mostrar la descripción
        If (tracker.ShowDescription) And Not tracker.AutoExpandSelItems Then
            .RowHeight(.Rows) = EvaluateTextHeight(.Rows, COL_DETAILEDDESC)
        End If
        
        If tracker.ColorItemsByPriority Then 'COLOR BY PRIORITY
            'Establece el color de fondo de la fila en funcion de la prioridad
            For i = 1 To .Columns - 1
                .CellBackColor(.Rows, i) = RGB(222, 227 - it.Priority * 8, 230 - it.Priority * 8)
                .cell(.Rows, i).ForeColor = CLR_GRIDTEXT
            Next
        End If
        
        If tracker.ColorOldItems Then 'COLOR OLD ITEMS
            'Si han pasado más de 30 días la fecha aparece en rojo
            If DateDiff("d", it.DateCreated, Date) > tracker.OldItemsDays Then .CellForeColor(.Rows, COL_DATECREATED) = CLR_OLDDATE
        End If
        
        .RowVisible(.Rows) = (Not it.Hidden) Or m_ShowHiddenItems 'Oculto?
    End With
    Next
    
    Set boldfnt = Nothing
    
End Sub

Private Sub Form_Load()
    tabTracker.ImageList = ilstTabs 'Image list del tab
    CreateGrid 'Columnas y disposición del grid
    grdTracker.OwnerDrawImpl = Me 'La interfaz IOwnerdraw está implementada aquí
    
    Update
    
    'mark the ShowTracker menu checkbox
    frmMain.cMenu.ItemChecked(frmMain.cMenu.IndexForKey("mnuProjectTracker")) = True
End Sub

 Public Sub Update()
    If Not openedProject Is Nothing Then
        'Establece la lista de desarrolladores
        Set devcol = openedProject.devcol
        'Establece la colección de trackers
        Set colTrackers = openedProject.colTrackers
              
        If openedProject.colTrackers.count = 0 Then 'No hay trackers
            NoTrackers = True
        Else
            Set AT = colTrackers(1) 'Tracker activo el primero
            RefreshTabs
            SelectTracker AT 'Selecciona el tracker Bugs
        End If
        grdTracker.Visible = True
    Else
        Set devcol = Nothing
        Set colTrackers = Nothing
        grdTracker.Clear
        grdTracker.Visible = False
        tabTracker.Tabs.Clear
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AT = Nothing
    Set colTrackers = Nothing
    
    'Uncheck the ShowTracker option
    frmMain.cMenu.ItemChecked(frmMain.cMenu.IndexForKey("mnuProjectTracker")) = False
    
End Sub

Private Sub Form_Initialize()
    'Init the properties
    bNoTrackers = False
    m_ShowTabs = True
End Sub

Private Sub grdTracker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then CreatePopupMenu "mnu"
End Sub

Private Sub grdTracker_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim i As Integer
    
    If lRow <= 0 Or lCol <= 0 Then Exit Sub
    'Muestra la descripción detallada de un elemento al seleccionarlo
    If AT.AutoExpandSelItems Then
        With grdTracker
            .Redraw = False
            'Establece la altura por defecto para todos
            For i = 1 To .Rows
                .RowHeight(i) = .DefaultRowHeight
            Next
            'Calcula la altura para la fila seleccionada
            If Not .RowIsGroup(lRow) Then
                .RowHeight(lRow) = EvaluateTextHeight(lRow, COL_DETAILEDDESC)
            End If
            .Redraw = True
        End With
    End If
    
End Sub

Private Sub grdTracker_ColumnClick(ByVal lCol As Long)
    'Ordenamos el grid
    AT.SortOrder = IIf(grdTracker.SortObject.SortOrder(1) = CCLOrderAscending, CCLOrderDescending, CCLOrderAscending)
    AT.SortColumn = lCol
    SelectTracker AT
    If AT.AutoExpandSelItems = True Then Call grdTracker_SelectionChange(grdTracker.SelectedRow, grdTracker.SelectedRow)
End Sub

Private Sub grdTracker_ColumnWidthReallyChanged(ByVal lCol As Long, lWidth As Long)
    If AT.ShowDescription Then
        SelectTracker AT
    ElseIf AT.AutoExpandSelItems = True Then
        SelectTracker AT
        Call grdTracker_SelectionChange(grdTracker.SelectedRow, grdTracker.SelectedRow)
    End If
End Sub

Private Sub lblLinkTM_Click()
    frmTrackerManager.Show 1
End Sub

Private Sub tabTracker_Resize()
    On Error Resume Next
    grdTracker.Move tabTracker.Left, tabTracker.Top, tabTracker.Width, tabTracker.Height - IIf(tabTracker.Visible, 380, 0)
End Sub

Private Sub tabTracker_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    SelectTracker colTrackers(Hex(theTab.ItemData))
    
    If iButton = vbRightButton Then
        CreatePopupMenu "mnu" 'Menu contextual
    End If
End Sub

Private Sub CreatePopupMenu(sKey As String)
    Dim rID As Long, rID2 As Long, lMnu As Long, i As Long
    Dim mnu As cMenus, lIndex As Long
    
    Set mnu = New cMenus
    Set mnu.ImageList = ilstMenus
    
    With mnu
        .DrawStyle = M_Style
        Call .CreateFromNothing(Me.hwnd)
        rID = .AddItem(0, Key:=sKey)

        .AddItem rID, "Add Item...", Key:="AddItem"
        If grdTracker.SelectedRow > 0 Then 'Si hay elementos seleccionados mostramos Edit
            .AddItem rID, "Edit Item...", Key:="EditItem"
            .AddItem rID, "Delete item", Key:="DelItem"
        End If
        .AddItem rID, "-"
        rID2 = .AddItem(rID, "Visible columns")
        For i = 1 To grdTracker.Columns - 1
            .AddItem rID2, grdTracker.ColumnKey(i), Checked:=AT.ColumnVisible(CLng(grdTracker.ColumnTag(i))), Key:="~COL" & i
            .ItemTag("~COL" & i) = grdTracker.ColumnTag(i)
            'ItemData(i - 1) = CLng(f.grdTracker.ColumnTag(i))
            'If AT.ColumnVisible(CLng(grdTracker.ColumnTag(i))) Then .ItemChecked(i + rID2) = True
        Next
        .AddItem rID, "Show hidden items (" & CountHiddenItems & ")", Key:="ShowHidden", Checked:=m_ShowHiddenItems
        
        .AddItem rID, "-"
        rID2 = .AddItem(rID, "Show tracker" & Space(5), , , "SelTracker")
        Dim tr As cTracker
        For Each tr In colTrackers
            'Añadimos los trackers
            .AddItem rID2, tr.name, , , "~TR" & tr.name
        Next
            
        .AddItem rID, "Tracker Manager...", , , "TrackerMan"
        .AddItem rID, "-"
        .AddItem rID, "Developers List", Key:="DevList"
        
        .AddItem rID, "-"
        .AddItem rID, "Show tabs", Key:="ShowTabs", Checked:=ShowTabs
        
        lMnu = .PopupMenu(sKey)
        If lMnu <> 0 Then
            lIndex = .IndexForID(lMnu)
            Select Case .ItemKey(lIndex)
            Case "TrackerMan" 'Tracker Manager
                frmTrackerManager.Show 1

            Case "AddItem" 'Añadir
                NewItem
            Case "EditItem" 'Editar
                EditItem
            Case "DelItem" 'Eliminar
                DeleteItem

            Case "DevList" 'Lista de desarrolladores
                frmDevelopersList.Show 1
                
            Case "ShowTabs" 'Mostar/Ocultar barra de lengüetas
                ShowTabs = Not .ItemChecked(lIndex)
                
            Case "ShowHidden"
                m_ShowHiddenItems = Not .ItemChecked(lIndex)
                SelectTracker AT
            End Select
            
            If .ItemKey(lIndex) Like "~TR*" Then 'Mostrar tracker *
                SelectTracker colTrackers(colTrackers.IndexForName(.ItemCaption(lIndex)))
                tabTracker.Tabs.item(.ItemCaption(lIndex)).Selected = True
            End If
            
            If .ItemKey(lIndex) Like "~COL*" Then 'Visible columns
                AT.ColumnVisible(CLng(.ItemTag(lIndex))) = Not .ItemChecked(lIndex)
                SelectTracker AT
            End If
                
        End If
    End With
    Set mnu = Nothing
End Sub

'Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant
'
'End Function

'Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
''On Error Resume Next
''    tabTracker.Move Left, Top, Right, Bottom
''    grbNoTrackers.Move Left, Top, Right, Bottom
''    'Center the labels of the notracker group box
''    With grbNoTrackers
''        lblNT2.Move (.Width - lblNT2.Width - lblLinkTM.Width) / 2, (.Height - lblNT2.Height) / 2
''        lblLinkTM.Move lblNT2.Left + lblNT2.Width, lblNT2.Top
''        lblNT1.Move (.Width - lblNT1.Width) / 2, lblNT2.Top - lblNT2.Height - 100
''    End With
'End Function
