VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.Form frmImport 
   Caption         =   "Import"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   4275
   WindowState     =   2  'Maximized
   Begin vbalIml6.vbalImageList ilFnt 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   10332
      Images          =   "frmImport.frx":2B8A
      Version         =   131072
      KeyCount        =   9
      Keys            =   "ÿÿÿÿÿÿÿÿ"
   End
   Begin CodeSenseCtl.CodeSense cs 
      Height          =   3135
      Left            =   0
      OleObjectBlob   =   "frmImport.frx":5406
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin vbalTBar6.cToolbar tbrImport 
      Height          =   375
      Left            =   120
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin vbalTBar6.cReBar rebar 
      Left            =   2280
      Top             =   0
      _ExtentX        =   2143
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmImport"
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

'MSG Constants (for future multi-language support)
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "File saved succesfully!"
'Private Const MSG_PAINTMAP_ERRORPAINTING = "An error occurred when trying to paint the fnt: "
'Private Const MSG_LOAD_ERRORLOADING = "An error occurred loading the fnt: "

Private Const FAST_SCROLL_STEPS As Integer = 12 'desplazamiento con Shift

Private m_IsDirty As Boolean 'This should never be set directly. Use the IsDirty property instead
Private m_Title

Public rangoActual As CodeSense.range
Private WithEvents m_ContextMenu As cMenus
Attribute m_ContextMenu.VB_VarHelpID = -1

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1
Private WithEvents m_ImportMenu As cMenus
Attribute m_ImportMenu.VB_VarHelpID = -1
Private m_FilePath As String
Private m_addToProject As Boolean

Implements IFileForm
Implements IPropertiesForm

Private Sub Form_Activate()
    frmMain.setStatusMessage ("Import file")
End Sub

Private Sub Form_Load()
    
'    'Configure toolbar
'    With tbrFnt
'        .ImageSource = CTBExternalImageList
'        .DrawStyle = T_Style
'        .SetImageList ilFnt.hIml, CTBImageListNormal
'        .CreateToolbar 16, True, True, True, 16
'        .AddButton "Zoom In", 0, , , , CTBAutoSize, "ZoomIn"
'        .AddButton "Restore Zoom", 1, , , , CTBAutoSize, "ZoomRestore"
'        .AddButton "Zoom Out", 2, , , , CTBAutoSize, "ZoomOut"
'        .AddButton eButtonStyle:=CTBSeparator
'        .AddButton "Toogle transparency", 3, , , "", CTBAutoSize, "ToogleTrans"
'        .AddButton eButtonStyle:=CTBSeparator
'        .AddButton "Edit palette", 5, , , , CTBAutoSize, "EditPalette"
'        .AddButton eButtonStyle:=CTBSeparator
'        .AddButton "Write text to map", 6, , , , CTBAutoSize, "WriteTexToMap"
'        .AddButton "Import font", 7, , , , CTBAutoSize, "ImportFont"
'        .AddButton "Export font", 8, , , , CTBAutoSize, "ExportFont"
'        '.AddButton "...", 4, , , "...", CTBDropDownArrow + CTBAutoSize, "AddToFpg"
'    End With
'    'Create the rebar
'    With rebar
'        If A_Bitmaps Then
'            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
'        End If
'        .CreateRebar Me.Hwnd
'        .AddBandByHwnd Me.tbrFnt.Hwnd, , True, False
'    End With
'    rebar.RebarSize

'    'Set up scroll bars:
'    Set m_cScroll = New cScrollBars
'    m_cScroll.create picScrollBox.Hwnd

    ' configure the edition control
    cs.LineNumbering = True
    cs.LineNumberStart = 1
    cs.LineNumberStyle = cmDecimal
    cs.LineToolTips = True
    cs.BorderStyle = cmBorderStatic
    cs.EnableDragDrop = True
    cs.EnableCRLF = True
    cs.TabSize = 2
    cs.ColorSyntax = False
    cs.Language = "Bennu"
    cs.DisplayLeftMargin = True
    cs.AutoIndentMode = cmIndentPrevLine
    LoadCSConf cs
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msgRes As VbMsgBoxResult
    'Ask for saving if the document is dirty
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
        rebar.RebarSize
        cs.Move 0, ScaleY(rebar.RebarHeight, vbPixels, vbTwips)
        cs.Width = Me.ScaleWidth
        cs.Height = Me.ScaleHeight - ScaleY(rebar.RebarHeight, vbPixels, vbTwips)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rebar.RemoveAllRebarBands 'Just for safety
End Sub


'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
Private Property Get IFileForm_AlreadySaved() As Boolean
    IFileForm_AlreadySaved = IIf(m_FilePath = "", False, True)
End Property

Private Function IFileForm_CloseW() As Long
    MsgBox "TO DO (CLOSE)"
    IFileForm_CloseW = 0
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

Private Property Get IFileForm_FileName() As String
    IFileForm_FileName = FSO.GetFileName(m_FilePath)
End Property

Private Property Get IFileForm_FilePath() As String
    IFileForm_FilePath = m_FilePath
End Property

Private Function IFileForm_Identify() As EFileFormConstants
    IFileForm_Identify = FF_MAP
End Function

Private Property Get IFileForm_IsDirty() As Boolean
    IFileForm_IsDirty = m_IsDirty
End Property

Private Function IFileForm_Load(ByVal sFile As String) As Long
    Dim lResult As Long
    Dim Ext As String
    
    Ext = FSO.GetExtensionName(sFile)
    
    cs.OpenFile (sFile)
    'There is no way to determine if the cs.openfile fails so assume everything goes well
    'since we check for the existence of the file in the NewFileForm function this should work
    'well (any file is supposed to be able to be opened in text format)
    lResult = -1
    m_FilePath = sFile
    IsDirty = False
       
    frmMain.setStatusMessage
    
    IFileForm_Load = lResult
End Function

Private Function IFileForm_NewW(ByVal iUntitledCount As Integer) As Long
    Dim sFiles() As String
    Dim fileCount As Integer
    Dim lResult As Long
    
    m_Title = "Untitled fnt " & CStr(iUntitledCount)
            
    frmMain.setStatusMessage ("Converting file to fnt format...")
    Screen.MousePointer = vbHourglass
    
    fileCount = ShowOpenDialog(sFiles, getFilter("IMPORTABLE_GRAPHICS"), False, False)
    
    If fileCount > 0 Then
        'lResult = fnt.Import(sFiles(0))
        'If lResult <> -1 Then
        '    MsgBox "An error ocurred trying to import the file: " & fnt.GetLastError, vbCritical
        'Else
            m_addToProject = modMenuActions.NewAddToProject
            IsDirty = True
        'End If
    End If
    
    Screen.MousePointer = 0
    frmMain.setStatusMessage
    
    IFileForm_NewW = lResult
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
 Dim lResult As Long
    Dim sFileBMK As String
    'Dim fs As FileSystemObject
    Dim A As textStream
    Dim i As Long
    
    
    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    On Error GoTo errhandler
    cs.SaveFile sFile, False 'Save the file
    

    ' HERE THERE SHOULD BE SOME KIND OF COMPROBATION FOR ERRORS AFTER SAVEFILE
    lResult = -1
    If (lResult) Then 'Save succesful
        ' Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then addFileToProject sFile
    
        If m_FilePath <> sFile Then 'Show a success message only if the name changed
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        m_FilePath = sFile
        IsDirty = False
    Else
        MsgBox MSG_SAVE_ERRORSAVING, vbCritical
    End If

errhandler:
    If Err.Number > 0 Then lResult = -1: Resume Next
    
End Function


Private Sub Cs_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    On Error Resume Next
    
    cs.HighlightedLine = -1
    
    Dim rangoTemp As CodeSense.range
    Set rangoTemp = cs.GetSel(True)
    
    If Not rangoActual Is Nothing Then
        'detectamos cambio de linea
        If rangoTemp.StartLineNo <> rangoActual.StartLineNo Then
            ' si se estaba mostrando el tooltip lo eliminamos
'            If Not showingToolTip Is Nothing Then
'                showingToolTip.Destroy
'                Set showingToolTip = Nothing
'            End If
        End If
    End If
    
    Set rangoActual = cs.GetSel(True) 'ubica la posicion actual y la guarda en una var de alcance modular
End Sub

Private Function Cs_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim lParentIndex, iP2 As Long
    Dim s, sl, sw, n, c As Boolean
    ' s for selected
    '   sl for single line selected
    '   sw for single word selected
    ' n for nothing selected
    ' c for converteable selections

On Error Resume Next

    s = False
    n = False
    sl = False
    sw = False
    c = False
    
    If rangoActual.IsEmpty Then
        n = True
    Else
        s = True
        If rangoActual.StartLineNo = rangoActual.EndLineNo Then
            If cs.SelText = cs.CurrentWord Then
                Debug.Print cs.SelText & "..." & cs.CurrentWord
                If isBin(cs.SelText) Or isHex(cs.SelText) Or IsNumeric(cs.SelText) Then
                    c = True
                End If
                sw = True
            Else
                sl = True
            End If
        End If
    End If
    
    If (Button = 2) Then
        
        Set m_ContextMenu = Nothing
        Set m_ContextMenu = New cMenus
        m_ContextMenu.DrawStyle = M_Style
        Set m_ContextMenu.ImageList = frmMain.ImgList1.Object
        m_ContextMenu.CreateFromNothing Me.Hwnd
        
        lParentIndex = m_ContextMenu.AddItem(0, Key:="ContextMenu")
        With m_ContextMenu
            If s Then
                .AddItem lParentIndex, "C&ut", "Ctrl+X", , "mnuEditCut", , , , 5
                .AddItem lParentIndex, "&Copy", "Ctrl+C", , "mnuEditCopy", , , , 4
            End If
            If cs.CanPaste Then
                .AddItem lParentIndex, "&Paste", "Ctrl+V", , "mnuEditPaste", , , , 6
            End If
            .AddItem lParentIndex, "-"
            If n Then
                .AddItem lParentIndex, "&Select all", "Ctrl+A", , "mnuEditSelectAll", , , , 75
                .AddItem lParentIndex, "Select line", "Ctrl+Shift+L", , "mnuEditSelectLine", , , , 76
                .AddItem lParentIndex, "Select word", "Ctrl+Shift+W", , "mnuEditSelectWord", , , , 86
            Else
                .AddItem lParentIndex, "Deselect", , , "mnuEditDeselect"
            End If
            If n Then
                .AddItem lParentIndex, "-"
                .AddItem lParentIndex, "Duplicate line", "Ctrl+D", , "mnuEditDuplicateLine", , , , 83
                .AddItem lParentIndex, "Delete line", "Ctrl+R", , "mnuEditDeleteLine", , , , 84
                .AddItem lParentIndex, "Clear line", , , "mnuEditClearLine"
                .AddItem lParentIndex, "Up line      ^", "Ctrl+Shift+Up", , "mnuEditUpLine", , , , 87
                .AddItem lParentIndex, "Down line  v", "Ctrl+Shift+Down", , "mnuEditDownLine", , , , 88
            End If

            If n Or sw Then
                .AddItem lParentIndex, "-"
            End If
            .AddItem lParentIndex, "&Search...", "Ctrl+F", , "mnuNavigationSearch", , , , 13
            If sw Or sl Then
                .AddItem lParentIndex, "Search next selected", "Ctrl+F3", , "mnuNavigationSearchNextWord", , , , 89
                .AddItem lParentIndex, "Search prev selected", "Ctrl+Shift+F3", , "mnuNavigationSearchPrevWord", , , , 90
            End If
            .AddItem lParentIndex, "-"
            .AddItem lParentIndex, "&Replace...", "Ctrl+H", , "mnuNavigationReplace", Image:=62
        
            .PopupMenu "ContextMenu"
        End With
        

    End If
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'This property is not part of the interface, just helps to reduce code
'by setting the caption of the form properly
Private Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property
Private Sub m_ContextMenu_Click(ByVal Index As Long)
    MsgBox m_ContextMenu.ItemKey(Index)
    Select Case m_ContextMenu.ItemKey(Index)
        Case "mnuEditCut":                      Call mnuEditCut
        Case "mnuEditCopy":                     Call mnuEditCopy
        Case "mnuEditPaste":                    Call mnuEditPaste
        Case "mnuEditSelectAll":                Call mnuEditSelectAll
        Case "mnuEditSelectWord":               Call mnuEditSelectWord
        Case "mnuEditSelectLine":               Call mnuEditSelectLine
        Case "mnuEditDeselect":                 Call mnuEditDeselect
        Case "mnuEditClearLine":                Call mnuEditClearLine
        Case "mnuEditDuplicateLine":            Call mnuEditDuplicateLine
        Case "mnuEditDeleteLine":               Call mnuEditDeleteLine
        Case "mnuEditUpLine":                   Call mnuEditUpLine
        Case "mnuEditDownLine":                 Call mnuEditDownLine
            
        Case "mnuNavigationSearch":             Call mnuNavigationSearch
        Case "mnuNavigationSearchNext":         Call mnuNavigationSearchNext
        Case "mnuNavigationSearchPrev":         Call mnuNavigationSearchPrev
        Case "mnuNavigationSearchNextWord":     Call mnuNavigationSearchNextWord
        Case "mnuNavigationSearchPrevWord":     Call mnuNavigationSearchPrevWord
        Case "mnuNavigationReplace":            Call mnuNavigationReplace
    End Select
    
End Sub

Private Sub tbrImport_ButtonClick(ByVal lButton As Long)
    Dim sKey As String
    
    sKey = tbrImport.ButtonKey(lButton)
    Select Case sKey
    Case "ToogleBookmark"
        modMenuActions.mnuBookmarkToggle
    Case "NextBookmark"
        modMenuActions.mnuBookmarkNext
    Case "PreviousBookmark"
        modMenuActions.mnuBookmarkPrev
    Case "DeleteBookmarks"
        modMenuActions.mnuBookmarkDel
    Case "ShiftRight"
        modMenuActions.mnuEditTab
    Case "ShiftLeft"
        modMenuActions.mnuEditUnTab
    Case "Comment"
        modMenuActions.mnuEditComment
    Case "Uncomment"
        modMenuActions.mnuEditUnComment
    Case "EditBookmarks"
        modMenuActions.mnuBookmarkEdit
    Case "PrevPos"
        modMenuActions.mnuNavigationPrevPosition
    Case "NextPos"
        modMenuActions.mnuNavigationNextPosition
    End Select
        
End Sub

Private Sub tbrImport_DropDownPress(ByVal lButton As Long)
'    Dim X As Long, Y As Long
'    Dim lIndex As Long
'    tbrMap.GetDropDownPosition lButton, X, Y
'
'    Select Case tbrMap.ButtonKey(lButton)
'        Case "AddToFpg":
'            createFpgsMenu
'            Call m_FpgsMenu.PopupMenu("FpgsMenu", _
'                Me.ScaleX(Me.Left + X, vbTwips, vbPixels), Me.ScaleY(Y, vbTwips, vbPixels) + rebar.RebarHeight * 1.5, TPM_VERNEGANIMATION + TPM_LEFTALIGN)
'    End Select
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'INTERFACE IPROPERTIESFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
Private Function IPropertiesForm_GetProperties() As cProperties
 
    Dim props As cProperties
    Set props = New cProperties

    With props
        
    End With

    Set IPropertiesForm_GetProperties = props
    
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IPROPERTIESFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

'-------------------------------------------------------------------------------------
'FUNCTION: Save()
'DESCRIPTION: Saves a Import file
'RETURNS: True if no error, otherwise False.
'-------------------------------------------------------------------------------------
Public Function Save(sFile As String) As Long
    Dim lSucceded As Boolean
    lSucceded = True
    Save = lSucceded
End Function



