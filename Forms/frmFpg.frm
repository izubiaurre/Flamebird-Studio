VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Begin VB.Form frmFpg 
   Caption         =   "Fpg editor"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmFpg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin vbalTBar6.cReBar cRebar 
      Left            =   0
      Top             =   600
      _ExtentX        =   5106
      _ExtentY        =   873
   End
   Begin vbalIml6.vbalImageList ilFpg 
      Left            =   1320
      Top             =   1800
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   4592
      Images          =   "frmFpg.frx":2B8A
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ADDMAPÿEDITMAPÿDELMAPÿ"
   End
   Begin vbAcceleratorSGrid6.vbalGrid grd 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      HighlightBackColor=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      DisableIcons    =   -1  'True
   End
   Begin vbalTBar6.cToolbar tbrFpg 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmFpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'Config constants
Private Const THUMB_WIDTH = 70
Private Const THUMB_HEIGHT = 70
Private Const GRID_COLUMNSEP = 22
Private Const GRID_ROWSEP = 3
Private Const GRID_NAME_HEIGHT = 14
'API Declarations
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_END_ELLIPSIS As Long = &H8000
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_CENTER As Long = &H1
'MSG Constants (for future multi-language support)
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "Fpg saved succesfully!"
Private Const MSG_ERROR_LOADING_MAP = "Error loading the map: "
Private Const MSG_ERROR_ADDING_MAP = "Error adding the map: "
Private Const MSG_LOAD_ERRORLOADING = "An error occurred loading the fpg: "

Private m_Fpg As New cFpg
Private m_IsDirty As Boolean
Private m_Title As String
Private m_FilePath As String
Private lastColSelected As Long
Private lastRowSelected As Long
Private m_addToProject As Boolean

Implements IFileForm
Implements IGridCellOwnerDraw
Implements IPropertiesForm

Public Property Get fpg() As cFpg
    If Not m_Fpg Is Nothing Then
        Set fpg = m_Fpg
    End If
End Property

Private Property Get SelectedMap() As cMap
    If grd.SelectionCount > 0 Then
        Set SelectedMap = m_Fpg.map(grd.SelectedCol + grd.Columns * (grd.SelectedRow - 1) - 1)
    End If
End Property

Public Function EditMapDescription(ByVal newVal As String) As Long
    If Not SelectedMap Is Nothing Then
        SelectedMap.description = newVal
        EditMapDescription = -1
        IsDirty = True
        grd.Redraw = True
    End If
End Function

'Configure grid
Private Sub ConfigureGrid()
    With grd
        .Redraw = False
        .Header = False
        .DefaultRowHeight = THUMB_HEIGHT + GRID_ROWSEP + GRID_NAME_HEIGHT
        .SelectionAlphaBlend = False
        .SelectionOutline = False
        .multiSelect = True
        .DrawFocusRectangle = False
        .HighlightForeColor = vbWindowText
        '.HotTrack = True
        .RowTextStartColumn = 1
        .OwnerDrawImpl = Me
        .Redraw = True
    End With
End Sub

'Calculates the optimal number of rows a cols
Public Sub CalculateGrid()
    Dim lRows As Long
    Dim lCols As Long
    Dim i As Integer
    
    If Not m_Fpg Is Nothing Then
        If m_Fpg.Available Then
            lCols = ScaleX(grd.Width, vbTwips, vbPixels) \ (THUMB_WIDTH + GRID_COLUMNSEP)
            lCols = IIf(lCols = 0, 1, lCols)
            lRows = m_Fpg.MapCount / lCols
            lRows = IIf(lRows = 0, 1, lRows)
            While (lRows * lCols < m_Fpg.MapCount)
                lRows = lRows + 1
            Wend
            
            With grd
                .Redraw = False
                While (.Columns > 0)
                    .RemoveColumn (1)
                Wend
                While (.Rows > 0)
                    .RemoveRow (1)
                Wend
                For i = 1 To lCols
                    .AddColumn lcolumnwidth:=THUMB_WIDTH + GRID_COLUMNSEP
                Next
                
                For i = 1 To lRows
                    .AddRow
                Next
                .Redraw = True
            End With
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Long
    i = grd.SelectedCol + grd.Columns * (grd.SelectedRow - 1)
    If i > 0 Then
        frmMain.setStatusMessage "Map nº: " & i & "         (" & fpg.map(i).Width & "x" & fpg.map(i).Height & ")                - Name: " & fpg.map(i).description
        'frmMain.StatusBar.PanelText("MAIN") = "Map nº: " & i & " - " & fpg.map(i).Width & "," & fpg.map(i).Height & " - Name: " & fpg.map(i).Description
    End If
End Sub

Private Sub Form_Load()
    'Configure toolbar
    With tbrFpg
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList ilFpg.hIml, CTBImageListNormal
        .CreateToolbar 16, True, True, True
        .AddButton "Show and edit fpg properties", 3, , , "Fpg properties", CTBAutoSize, "Properties"
        .AddButton , , , , , CTBSeparator
        .AddButton "Adds an existing map", 0, , , "Add", CTBAutoSize, "AddMap"
        .AddButton "Edit selected map", 1, , , "Edit", CTBAutoSize, "EditMap"
        .AddButton "Delete selected map", 2, , , "Delete", CTBAutoSize, "DeleteMap"
    End With
    'Create the rebar
    With cReBar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        End If
        .CreateRebar Me.Hwnd
        .AddBandByHwnd tbrFpg.Hwnd, , True, False
    End With
    'Configure grid
    ConfigureGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msgRes As VbMsgBoxResult
    'Ask for saving if the document is dirty and we are not closing the entire application
    If IFileForm_IsDirty = True And UnloadMode <> vbFormMDIForm Then
        msgRes = MsgBox("The file '" & IFileForm_Title & "' is modified. " _
                    & "Save it?", vbYesNoCancel + vbQuestion, "Save")
        If msgRes = vbYes Then 'Save
            SaveFileOfFileForm Me
        ElseIf msgRes = vbCancel Then 'Cancel
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        Me.grd.Move 0, ScaleY(cReBar.RebarHeight, vbPixels, vbTwips), Me.ScaleWidth, Me.ScaleHeight
        CalculateGrid
        cReBar.RebarSize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cReBar.RemoveAllRebarBands 'Just for safety
End Sub

Private Sub grd_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    'Avoid selecting empty cells
    Dim i As Long
    
    On Error GoTo errhandler
    
    If Not m_Fpg Is Nothing Then
        If (lCol + grd.Columns * (lRow - 1)) > m_Fpg.MapCount Then
            grd.CellSelected(lRow, lCol) = False
            If lastRowSelected > 0 And lastColSelected > 0 Then
                grd.CellSelected(lastRowSelected, lastColSelected) = True
            End If
        Else
            lastColSelected = lCol
            lastRowSelected = lRow
        End If
    End If
    frmProperties.RefreshProperties
    
    i = grd.SelectedCol + grd.Columns * (grd.SelectedRow - 1)
    If i >= 0 Then
        frmMain.setStatusMessage "Map nº: " & i & "         (" & fpg.map(i - 1).Width & "x" & fpg.map(i - 1).Height & ")            - Name: " & fpg.map(i - 1).description
        'frmMain.StatusBar.PanelText("MAIN") = "Map nº: " & i & " - " & fpg.map(i - 1).Width & "," & fpg.map(i - 1).Height & " - Name: " & fpg.map(i - 1).Description
    End If
    
    Exit Sub
    
errhandler:
    Resume Next
End Sub

'Shows the opendialog to adds maps to the fpg
Private Sub addMap()
    Dim m As New cMap
    Dim msgResult As VbMsgBoxResult
    Dim replace As Boolean
    Dim lResult As Long
    Dim i As Integer, fileCount As Integer
    Dim sFiles() As String
    
    On Error GoTo errhandler
    
    fileCount = ShowOpenDialog(sFiles, getFilter("MAP"), False, True)
    If fileCount > 0 Then
        For i = LBound(sFiles) To UBound(sFiles)
            Set m = New cMap
            If m.Load(sFiles(i)) = -1 Then 'Map load succesfull
                'The code already exists
                If m_Fpg.Exists(m.Code) Then
                    msgResult = MsgBox("This FPG contains another map with the code: " & m.Code & _
                                    ". Replace it?", vbQuestion + vbYesNoCancel)
                    If msgResult = vbYes Then
                        replace = True
                    ElseIf msgResult = vbCancel Then
                        Exit Sub
                    Else
                        m.Code = m_Fpg.FreeCode
                    End If
                End If
                'Add the map to the fpg
                If m_Fpg.Add(m, replace) <> -1 Then
                    MsgBox MSG_ERROR_ADDING_MAP & m_Fpg.GetLastError
                End If
                IsDirty = True
            Else
                MsgBox MSG_ERROR_LOADING_MAP & m.GetLastError
            End If
        Next
    End If
    
    Exit Sub
    
errhandler:
    If Err.Number > 0 Then ShowError "frmFpg.AddMap"
End Sub

Private Sub tbrFpg_ButtonClick(ByVal lButton As Long)
    Select Case tbrFpg.ButtonKey(lButton)
    Case "Properties" 'FPG properties
        grd.ClearSelection
        frmProperties.RefreshProperties
    Case "AddMap"
        addMap
        CalculateGrid
    Case "DeleteMap"
        
        m_Fpg.Remove SelectedMap.Code
        CalculateGrid
    Case "EditMap"
        MsgBox "EditMap"
    End Select
End Sub

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

Private Property Get IFileForm_AlreadySaved() As Boolean
    IFileForm_AlreadySaved = IIf(m_FilePath = "", False, True)
End Property

Private Function IFileForm_CloseW() As Long
    MsgBox "TODO CLOSE"
    IFileForm_CloseW = 0
End Function

Private Property Get IFileForm_FileName() As String
    IFileForm_FileName = FSO.GetFileName(m_FilePath)
End Property

Private Property Get IFileForm_FilePath() As String
    IFileForm_FilePath = m_FilePath
End Property

Private Function IFileForm_Identify() As EFileFormConstants
    IFileForm_Identify = FF_FPG
End Function

Private Property Get IFileForm_IsDirty() As Boolean
    IFileForm_IsDirty = m_IsDirty
End Property

Private Function IFileForm_Load(ByVal sFile As String) As Long
    Dim lResult As Long
    
    Screen.MousePointer = vbHourglass
    lResult = m_Fpg.Load(sFile)  'Load fpg
    If (lResult) Then
        m_FilePath = sFile
        CalculateGrid 'Refreshes grid
        IsDirty = False
    Else
        MsgBox MSG_LOAD_ERRORLOADING + m_Fpg.GetLastError, vbCritical
    End If
    Screen.MousePointer = 0
    
    IFileForm_Load = lResult
End Function

Private Function IFileForm_NewW(ByVal iUntitledCount As Integer) As Long
    m_Title = "Untitled fpg " & CStr(iUntitledCount)
    Caption = m_Title

    m_addToProject = modMenuActions.NewAddToProject
    Set m_Fpg = New cFpg
    IFileForm_NewW = m_Fpg.New16
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
    Dim lResult As Long
    
    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    lResult = m_Fpg.Save(sFile) 'Save the map
    If (lResult) Then 'Save succesful
        'Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then
            If Not openedProject Is Nothing Then addFileToProject sFile
        End If
        
        If m_FilePath <> sFile Then 'Show a success message only if the file name is different
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        
        IsDirty = False
        m_FilePath = sFile
    Else
        MsgBox MSG_SAVE_ERRORSAVING + m_Fpg.GetLastError, vbCritical
    End If
End Function

Private Property Get IFileForm_Title() As String
    Dim sTitle As String
    If IFileForm_FilePath = "" Then
        sTitle = m_Title
    Else
        sTitle = IFileForm_FileName
    End If
    IFileForm_Title = sTitle
End Property

Public Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

Private Sub IGridCellOwnerDraw_Draw(cell As cGridCell, ByVal lHdc As Long, ByVal eDrawStage As ECGDrawStage, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, bSkipDefault As Boolean)
    Dim fXSizeIndex As Single, fYSizeIndex As Single
    Dim X As Long, Y As Long
    Dim map As cMap
    If Not m_Fpg Is Nothing Then
        If m_Fpg.Available Then
            If (eDrawStage = ecgBeforeIconAndText) Then
                Set map = m_Fpg.map(cell.Column + grd.Columns * (cell.Row - 1) - 1)
                If Not map Is Nothing Then
                    bSkipDefault = True
                    'Calculate the size and the coords
                    X = (GRID_COLUMNSEP) / 2 + lLeft
                    X = IIf(X < lLeft, lLeft, X)
                    Y = lTop
                    'Draw the thumb image
                    map.Draw lHdc, X, Y, THUMB_WIDTH, THUMB_HEIGHT, False
                    'Draw Name of the map
                    Dim rc As RECT
                    rc.Left = lLeft - 2
                    rc.Right = lRight - 2
                    rc.Top = lBottom - GRID_NAME_HEIGHT
                    rc.Bottom = lBottom
                    DrawTextA lHdc, map.description, -1, rc, DT_VCENTER + DT_MODIFYSTRING + DT_END_ELLIPSIS + DT_CENTER
                Else
                  '  MsgBox "Error en IGRidCellOwnerdraw"
                End If
            End If
        End If
    End If
End Sub

Private Function IPropertiesForm_GetProperties() As cProperties
    Dim cp As String
    Dim i As Integer
    Dim cpoint() As Integer
    Dim props As cProperties
    Dim map As cMap
    Set props = New cProperties
    
    If Not m_Fpg Is Nothing Then
        If grd.SelectionCount > 0 Then
            If grd.SelectionCount = 1 Then 'One item selected
                Set map = SelectedMap
                If Not map Is Nothing Then 'Valid selected map
                    With props
                    .Add "Description", "Description", ptText, Me, "EditMapDescription", map.description, True, 32
                    .Add "Width", "Width", ptNumeric, Me, "WidthP_Changed", map.Width, False
                    .Add "Height", "Height", ptNumeric, Me, "HeightP_changed", map.Height, False
                    .Add "Code", "Code", ptInteger, Me, "EditCode", map.Code, False, 9999, 0, False
                    'CPointsCount()-1
                    For i = 0 To map.CPointsCount() - 1
                        cpoint() = map.ControlPoint(i)
                        cp = cp & "(" & i & ": " & cpoint(0) & ", " & cpoint(1) & ");"
                    Next
                    .Add "C.Points", "CP", ptLink, Me, "EditCP", cp, False
                    End With
                    props("Description").description = "Specifies the name of the map in the FPG"
                    props("Code").description = "Specifies the code of the map in the FPG"
                    props("Width").description = "Specifies the with of the MAP"
                    props("Height").description = "Specifies the height of the MAP"
                    props("CP").description = "Set control points (including center)"
                End If
            Else 'More than one item selected
            End If
        Else 'No maps selected -->Properties of the FPG
            With props
             .Add "Maps", "Maps", ptNumeric, Me, "", m_Fpg.MapCount, False
             .Add "Depth", "Depth", ptCombo, Me, "", (m_Fpg.Depth \ 8) - 1, False
             .Add "Palette", "Palette", ptLink, Me, "", IIf(m_Fpg.Depth = 8, "Palette 256 colors", "None"), False
            End With
            
            props("Depth").AddOption "8 bits"
            props("Depth").AddOption "16 bits"
            props("Maps").description = "Number of maps in the Fpg"
            props("Depth").description = "Bits per pixel of the Fpg"
            props("Palette").description = "Palette to use with the Fpg (only 8 bits)"
        End If
    End If


    Set IPropertiesForm_GetProperties = props
End Function

