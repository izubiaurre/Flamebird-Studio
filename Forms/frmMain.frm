VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Object = "{04609A82-EA10-423E-B61B-CACCC4ABDFCF}#1.0#0"; "tabdock.ocx"
Object = "{4F11FEBA-BBC2-4FB6-A3D3-AA5B5BA087F4}#1.0#0"; "vbalsbar6.ocx"
Object = "{A2FCE68E-ACA2-4A09-9C2D-F53E1195D5FC}#1.0#0"; "vbalhkct6.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "Flamebird"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin vbalHkCt6.vbalHookControl ctlAccelerator 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   0
      Width           =   9075
      Begin vbalTBar6.cToolbar tbrMain 
         Height          =   375
         Left            =   4440
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
      End
   End
   Begin vbalTBar6.cReBar cReBar 
      Left            =   4080
      Top             =   3960
      _ExtentX        =   2355
      _ExtentY        =   873
   End
   Begin vbalIml6.vbalImageList ImgListDialog 
      Left            =   6240
      Top             =   5280
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   8
      Size            =   17648
      Images          =   "frmMain.frx":6852
      Version         =   131072
      KeyCount        =   4
      Keys            =   "PRGÿMAPÿFBPÿFPG"
   End
   Begin vbalSbar6.vbalStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      SimpleStyle     =   0
   End
   Begin TabDock.TTabDock TabDock 
      Left            =   5880
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      AutoShowCollapseCaptions=   0   'False
      AutoExpand      =   0   'False
      CollapseInterval=   0
      Gradient1       =   0
   End
   Begin vbalIml6.vbalImageList ImgList1 
      Left            =   6960
      Top             =   5280
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   109060
      Images          =   "frmMain.frx":AD62
      Version         =   131072
      KeyCount        =   95
      Keys            =   $"frmMain.frx":25786
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuNavigation 
      Caption         =   "&Navigation"
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
   End
   Begin VB.Menu mnuExecute 
      Caption         =   "&Run"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "P&lugins"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
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

'Note that some of the features of this demonstration will
'not work correctly unless you have the correct version of
'COMCTRL32.DLL installed (4.71 or higher required)

Option Explicit
Option Compare Text

Implements ISubclass

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal Hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function InvalidateRectBynum Lib "user32.dll" Alias "InvalidateRect" (ByVal Hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function ValidateRect Lib "user32.dll" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal Hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal Hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Const WM_MDINEXT = &H224
Private Const WM_MDICREATE As Long = &H220
Private Const WM_MDIDESTROY As Long = &H221

Private WithEvents MDITabs As cMDITabs
Attribute MDITabs.VB_VarHelpID = -1
Public WithEvents cMenu As cMenus
Attribute cMenu.VB_VarHelpID = -1
Public CancelUnload As Boolean 'Used to interact with frmSave
Private WithEvents fileMenu As cMenus
Attribute fileMenu.VB_VarHelpID = -1
Private bWasMinimized As Boolean

'-------------------------------------------------------------------------------
'GENERAL
'-------------------------------------------------------------------------------

Public Sub setStatusMessage(Optional Msg As String = "")
    StatusBar.PanelText("MAIN") = Msg
    Call StatusBar.RedrawPanel("MAIN")
End Sub

'Returns a reference to the current ActiveForm as an IFileForm
Public Property Get ActiveFileForm() As IFileForm
    If Not ActiveForm Is Nothing Then
        If TypeOf ActiveForm Is IFileForm Then
            Set ActiveFileForm = ActiveForm
        Else
            Set ActiveFileForm = Nothing
        End If
    End If
End Property

Public Sub RefreshTabs()
    If Not MDITabs Is Nothing Then
        MDITabs.ForceRefresh
    End If
End Sub

Private Function FormForHwnd(ByVal Hwnd As Long) As Form
   Dim frmChild As Form
   For Each frmChild In Forms
      If (frmChild.Hwnd = Hwnd) Then
         Set FormForHwnd = frmChild
         Exit For
      End If
   Next
End Function

Public Sub CreateToolsMenu()
    Dim i As Integer
    Dim parentIndex As Long
    Dim toDelete() As String
    Dim nToDelete As Integer
    
    parentIndex = cMenu.IndexForKey("mnuTools")
'    'Clear existing tools
'    ReDim toDelete(cMenu.ItemCount) As String
'    nToDelete = 0
'    For i = 1 To cMenu.ItemCount
'        If cMenu.ItemKey(i) Like "TOOL*" Then
'            toDelete(nToDelete) = cMenu.ItemKey(i)
'            nToDelete = nToDelete + 1
'        End If
'    Next
'    For i = 0 To nToDelete - 1
'        cMenu.RemoveItem cMenu.IndexForKey(toDelete(i))
'    Next

    'Add tools
    If ExternalToolsCount > 0 Then
        cMenu.AddItem parentIndex, "-"
        For i = 0 To ExternalToolsCount - 1
            cMenu.AddItem parentIndex, ExternalTools(i).Title, , , "TOOL" & CStr(i)
        Next
    End If
End Sub


'-------------------------------------------------------------------------------
'CONTROLS
'-------------------------------------------------------------------------------

Private Sub MDIForm_Load()
    'Subclass WM_GETMINMAXINFO to control the minimun size
    MinHeight = 240
    MinWidth = 320
    AttachMessage Me, Me.Hwnd, WM_GETMINMAXINFO
    
    Caption = App.Title & " " & App.Major & "." & App.Minor & App.Revision

    CreateInterface
    
    InitKey_Accelerators
    
    If Command <> "" Then 'Your program is asked to open some file
        Dim fname As String
        'fname = GetLongFilename(Command)
        fname = FSO.GetAbsolutePathName(Command)
        OpenFileByExt fname
        RefreshTabs
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState <> vbMinimized Then
        cRebar.RebarSize
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim areDirtyForms As Boolean
    Dim ff As IFileForm
    Dim f As Form
    
    'Look for at least one fileform with changes and show the Save Form
    areDirtyForms = False
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.IsDirty = True Then
                areDirtyForms = True
                Exit For
            End If
        End If
    Next
    If areDirtyForms = True Then
        frmSave.Show vbModal, Me
    End If
    
    'To unload or not to unload...
    If CancelUnload = True Then
        CancelUnload = False
        Cancel = 1
    Else
        SaveIDEConf
        
        If Not openedProject Is Nothing Then CloseProject 'Close project (and so its also saved)
        DetachMessage Me, Me.Hwnd, WM_GETMINMAXINFO 'End subclassification
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
    Dim archivo
    
    For Each archivo In Data.Files
        OpenFileByExt archivo
    Next archivo
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    cRebar.RemoveAllRebarBands 'Just for safety
    
    Unload frmProjectBrowser
    Unload frmProperties
    Unload frmProgramInspector
    Unload frmOutput
    Unload frmDebug
    Unload frmErrors
    
    finishFMOD
End Sub

Private Sub MDITabs_CloseWindow(ByVal Hwnd As Long)
    Dim f As Form
    Dim tr As RECT
    
    GetWindowRect frmMain.Hwnd, tr
    LockWindowUpdate Hwnd
    
    Set f = FormForHwnd(Hwnd)
    If Not f Is Nothing Then Unload f
    
    LockWindowUpdate False
    InvalidateRect 0, tr, 0 'TODO: See if we cann refresh only the part bellow the toolbar
End Sub

Private Sub MDITabs_TabBarClick(ByVal iButton As MouseButtonConstants, ByVal screenX As Long, ByVal screenY As Long)
    Dim lParentIndex As Long
    Dim rID As Long
    If iButton = vbLeftButton Then
        mnuFileNewFile
    ElseIf iButton = vbMiddleButton Then
        If MsgBox("Close all the opened files?", vbQuestion + vbYesNo, "Closing all opened files") = vbYes Then
        End If
    ElseIf iButton = vbRightButton Then
        Set fileMenu = New cMenus
            With fileMenu
                .DrawStyle = M_Style
                Set .ImageList = ImgList1.Object
                Call .CreateFromNothing(Me.Hwnd)
                lParentIndex = .AddItem(0, Key:="FileMenu")
                .AddItem lParentIndex, "&File...", "Ctrl+O", , "mnuFileOpenFile", , , , 1
                .AddItem lParentIndex, "-"
                .AddItem lParentIndex, "&Source...", , , "mnuFileOpenSource", , , , 19
                .AddItem lParentIndex, "&Map...", , , "mnuFileOpenMap", , , , 21
                .AddItem lParentIndex, "Fp&g...", , , "mnuFileOpenFpg", , , , 22
                .AddItem lParentIndex, "F&nt...", , , "mnuFileOpenFnt", , , , 91
                .AddItem lParentIndex, "&Import...", , , "mnuFileOpenImp"
                rID = .PopupMenu("FileMenu")
            End With
        
        If (rID <> 0) Then
            fileMenu.CurrentMenuIndex = fileMenu.IndexForID(rID)
        
            Select Case fileMenu.ItemKey(fileMenu.CurrentMenuIndex)
                Case "mnuFileOpenFile"
                    mnuFileOpenFile
                Case "mnuFileOpenSource"
                    mnuFileOpenSource
                Case "mnuFileOpenMap"
                    mnuFileOpenMap
                Case "mnuFileOpenFpg"
                    mnuFileOpenFpg
                Case "mnuFileOpenFnt"
                    mnuFIleOpenFnt
                Case "mnuFileOpenImp"
                    mnuFileOpenImp
            End Select
        End If
        Set fileMenu = Nothing
    End If
End Sub

Private Sub MDITabs_TabClick(ByVal iButton As MouseButtonConstants, ByVal Hwnd As Long, ByVal screenX As Long, ByVal screenY As Long)
    Dim lParentIndex As Long
    Dim rID As Long
    If iButton = vbRightButton Then
        Set fileMenu = New cMenus
            With fileMenu
                .DrawStyle = M_Style
                Call .CreateFromNothing(Me.Hwnd)
                lParentIndex = .AddItem(0, Key:="FileMenu")
                '.AddItem lParentIndex, "Always run this project", , , "markasmain"
                .AddItem lParentIndex, "Save", , , "save"
                .AddItem lParentIndex, "Close", , , "close"
                rID = .PopupMenu("FileMenu")
            End With
        
        If (rID <> 0) Then
            fileMenu.CurrentMenuIndex = fileMenu.IndexForID(rID)
        
            Select Case fileMenu.ItemKey(fileMenu.CurrentMenuIndex)
                Case "save"
                    mnuFileSave
                Case "close"
                    mnuFileClose
            End Select
        End If
        Set fileMenu = Nothing
    ElseIf iButton = vbMiddleButton Then ' close tab
        MsgBox "Close!"
       mnuFileClose
    End If
End Sub

Private Sub MDITabs_WindowChanged(ByVal Hwnd As Long)
On Error GoTo errhandler:
    Dim f As Form
    Dim fFileForm As IFileForm
    Dim file As String

    If openedProject Is Nothing Then Exit Sub

    Set f = FormForHwnd(Hwnd)

    If TypeOf f Is IFileForm Then
        Set fFileForm = f
        file = fFileForm.FilePath
        ' Find file in PB
        frmProjectBrowser.FindAndSelect (file)
        ' enable/disable Program inspector
        If fFileForm.Identify = FF_SOURCE Then
            If frmProgramInspector.Visible Then
                frmProgramInspector.tv_program.Enabled = True
                frmProgramInspector.tv_program.BackColor = RGB(256, 256, 256)
            End If
        Else
            If frmProgramInspector.Visible Then
                frmProgramInspector.tv_program.Enabled = False
                frmProgramInspector.tv_program.BackColor = RGB(128, 128, 128)
            End If
        End If
    End If
    frmProperties.RefreshProperties
    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("TabClick"): Resume Next
End Sub

'-------------------------------------------------------------------------------
' INTERFACE RELATED
'-------------------------------------------------------------------------------

Public Sub SaveIDEConf()
    Dim i As Integer
    
    'Save state
    With Ini
        .Path = App.Path & CONF_FILE
        .Section = Me.name
        .SaveFormPosition Me

        For i = 1 To TabDock.DockedForms.count
            .Section = TabDock.DockedForms(i).Key

            .Key = "dockedState"
            .Value = TabDock.DockedForms(TabDock.DockedForms(i).Key).state

            .Key = "align"
            .Value = TabDock.DockedForms(TabDock.DockedForms(i).Key).Panel.align

            .Key = "style"
            .Value = TabDock.DockedForms(TabDock.DockedForms(i).Key).Style

            .Key = "percent"
            .Value = TabDock.DockedForms(TabDock.DockedForms(i).Key).percent

            .Key = "visible"
            .Value = IIf(TabDock.DockedForms(TabDock.DockedForms(i).Key).Visible = True, "1", "0")
        Next i
    End With
End Sub

'VERY IMPORTANT: Variables are passed byref and its values are used to se the default values
'of the ini. This variables are modified with the values in the ini file
Private Sub LoadDockedFormConf(ByVal formName As String, dockedState As tdDockedState, _
        align As tdAlignProperty, Style As tdDockStyles, percent As Long, _
        expanded As Boolean, isVisible As Boolean)
    With Ini
        .Path = App.Path & CONF_FILE
        .Section = formName
        
        .Key = "dockedState" 'Docked state
        .Default = CStr(dockedState)
        dockedState = IIf(.Value = "1", 1, 2)
        
        .Key = "align" 'Align
        .Default = CStr(align)
        align = CInt(.Value)
        
        .Key = "style" 'Style
        .Default = CStr(Style)
        Style = CInt(.Value)
        
        .Key = "percent" 'Percent (size)
        .Default = CStr(percent)
        percent = CInt(.Value)
        
'        .Key = "expanded"
'        '.Default = expanded
'        expanded = .value
        
        .Key = "visible"
        .Default = IIf(isVisible = True, "1", "0")
        isVisible = IIf(.Value = "1", True, False)
    End With
End Sub

Private Function CalculateToolBarWidth(tBR As cToolbar) As Integer
    Dim i As Integer
    Dim w As Integer
    If Not tBR Is Nothing Then
        For i = 0 To tBR.ButtonCount
            w = w + tBR.ButtonWidth(i)
        Next
    End If
    CalculateToolBarWidth = w
End Function

Private Sub CreateInterface()
    'Set the main form position
    With Ini
        .Path = App.Path & CONF_FILE
        .Section = Me.name
        .LoadFormPosition Me, Me.Width, Me.Height
    End With
    
    CreateMenu
    LoadPlugins
    LoadRecents
    
    CreateToolBars

    'Create the MDI Tabs
    Set MDITabs = New cMDITabs
    MDITabs.Attach Me.Hwnd
    MDITabs.ColorOf(TC_ACTIVETABBG) = MDITabs.ColorOf(TC_TOPBARLINE)
    MDITabs.ColorOf(TC_CHILDRECT) = MDITabs.ColorOf(TC_TOPBARLINE)
    MDITabs.DrawIcons = True

    'Configure status bar
    With StatusBar
        .ImageList = ImgList1.hIml
        .AddPanel estbrStandard, , , , , True, , , "MAIN"
        .AddPanel estbrStandard, "Mode Normal", "Shows current selection mode", , 240, , , , "SEL_MODE"
        .AddPanel estbrOwnerDraw, , , , 46, , , , "FXI_OUTPUT_INFO"
        .AddPanel estbrCaps
        .AddPanel estbrNum
        .AddPanel estbrScrl
        .AddPanel estbrIns
        .AddPanel estbrNoBorders, , , , 3
    End With
    CreateDockableInterface
End Sub

Public Sub CreateMenuFromStrMatrix(oMenu As cMenus, ParentKey As String, _
                baseKey As String, StrMatrix() As String)
    Dim i As Integer, mnuKey As String
    With oMenu
        'Recorremos los elementos del array y los añadimos al menu
        For i = LBound(StrMatrix) To UBound(StrMatrix)
            If StrMatrix(i) <> "" Then
                mnuKey = baseKey & CStr(i)
                If CBool(.IndexForKey(mnuKey)) Then
                    .ItemCaption(.IndexForKey(CStr(mnuKey))) = StrMatrix(i) 'Cambiamos el nombre
                Else
                    .AddItem .IndexForKey(ParentKey), StrMatrix(i), , , CStr(mnuKey)
                End If
            End If
        Next
    End With
End Sub

Private Sub CreateMenu()
    Dim iP, iP2 As Long
    
    Set cMenu = New cMenus
    
    With cMenu
        .DrawStyle = M_Style
        'Set the image list
        Set .ImageList = ImgList1.Object
        Call .CreateFromForm(Me)
    End With
    
    With cMenu
        
        'MENU FILE
        iP = .IndexForKey("mnuFile")
            iP2 = .AddItem(iP, "&New", , , "mnuFileNew")
                .AddItem iP2, "&File...", , , "mnuFileNewFile", , , , 1
                .AddItem iP2, "-"
                .AddItem iP2, "&Project...", , , "mnuFileNewProject", , , , 20
                .AddItem iP2, "&Source", "Ctrl+N", , "mnuFileNewSource", , , , 19
                .AddItem iP2, "&Map", "Ctrl+M", , "mnuFileNewMap", , , , 21
                .AddItem iP2, "Fp&g", , , "mnuFileNewFpg", , , , 22
                '.AddItem iP2, "F&nt", , , "mnuFileNewFnt"
                .AddItem iP2, "&Import", , , "mnuFileNewImp"
            iP2 = .AddItem(iP, "&Open", , , "mnuFileOpen", , , , 2)
                .AddItem iP2, "&File...", "Ctrl+O", , "mnuFileOpenFile", , , , 1
                .AddItem iP2, "-"
                .AddItem iP2, "&Project...", , , "mnuFileOpenProject", , , , 20
                .AddItem iP2, "&Source...", , , "mnuFileOpenSource", , , , 19
                .AddItem iP2, "&Map...", , , "mnuFileOpenMap", , , , 21
                .AddItem iP2, "Fp&g...", , , "mnuFileOpenFpg", , , , 22
                .AddItem iP2, "F&nt...", , , "mnuFileOpenFnt", , , , 91
                .AddItem iP2, "&Song...", , , "mnuFileOpenSong", , , , 53
                .AddItem iP2, "&Import ...", , , "mnuFileOpenImp"
            .AddItem iP, "&Close", "Ctrl+W", , "mnuFileClose", , , , 33
            .AddItem iP, "Close A&ll", , , "mnuFileCloseAll", , , , 34
            .AddItem iP, "-"
            .AddItem iP, "Close pro&ject", , , "mnuFileCloseProject", False, , , 12
            .AddItem iP, "-"
            .AddItem iP, "&Save", "Ctrl+S", , "mnuFileSave", , , , 3
            .AddItem iP, "Save &as...", , , "mnuFileSaveAs", , , , 73
            .AddItem iP, "Save &All", "Ctrl+Shift+S", , "mnuFileSaveAll", , , , 31
            .AddItem iP, "-"
            .AddItem iP, "&Print...", "Ctrl+P", , "mnuFilePrint", Image:=67
            .AddItem iP, "-"
            .AddItem iP, "Recent &Files", , , "mnuFileRecentFiles"
            .AddItem iP, "&Recent Projects", , , "mnuFileRecentProjects"
            .AddItem iP, "-"
            .AddItem iP, "Preferences...", "F10", , "mnuEditPreferences", Image:=11
            .AddItem iP, "-"
            .AddItem iP, "&Exit", "Ctrl+Q", , "mnuFileExit", , , , 32

        'MENU EDIT
        iP = .IndexForKey("mnuEdit")
            .AddItem iP, "Undo", "Ctrl+Z", , "mnuEditUndo", , , , 7
            .AddItem iP, "Redo", "Ctrl+Y", , "mnuEditRedo", , , , 8
            .AddItem iP, "-"
            .AddItem iP, "C&ut", "Ctrl+X", , "mnuEditCut", , , , 5
            .AddItem iP, "&Copy", "Ctrl+C", , "mnuEditCopy", , , , 4
            .AddItem iP, "&Paste", "Ctrl+V", , "mnuEditPaste", , , , 6
            .AddItem iP, "-"
            .AddItem iP, "&Select all", "Ctrl+A", , "mnuEditSelectAll", , , , 75
            .AddItem iP, "Select line", "Ctrl+Shift+L", , "mnuEditSelectLine", , , , 76
            .AddItem iP, "Select word", "Ctrl+Shift+W", , "mnuEditSelectWord", , , , 86
            .AddItem iP, "Deselect", , , "mnuEditDeselect"
            .AddItem iP, "-"
            .AddItem iP, "Duplicate line", "Ctrl+D", , "mnuEditDuplicateLine", , , , 83
            .AddItem iP, "Delete line", "Ctrl+R", , "mnuEditDeleteLine", , , , 84
            .AddItem iP, "Clear line", , , "mnuEditClearLine"
            .AddItem iP, "Up line      ^", "Ctrl+Shift+Up", , "mnuEditUpLine", , , , 87
            .AddItem iP, "Down line  v", "Ctrl+Shift+Down", , "mnuEditDownLine", , , , 88
            .AddItem iP, "-"
'            .AddItem iP, "Delete word from cursor pos", "Ctrl+Del", , "mnuEditDeleteWordFromCursor", , , , 79
'            .AddItem iP, "Delete word", "Ctrl+Alt+Del", , "mnuEditDeleteWord"
'            .AddItem iP, "-"
            .AddItem iP, "Shift line &left", "Tab", , "mnuEditTab", Image:=40
            .AddItem iP, "Shift line &right", "Shift+Tab", , "mnuEditUnTab", Image:=41
            .AddItem iP, "-"
            .AddItem iP, "&Comment", "Ctrl+J", , "mnuEditComment", Image:=42
            .AddItem iP, "U&nComment", "Ctrl+Shift+J", , "mnuEditUnComment", Image:=43
            .AddItem iP, "-"
            .AddItem iP, "&UPPER CASE", "Ctrl+U", , "mnuEditUpperCase", Image:=60
            .AddItem iP, "lo&wer case", "Ctrl+L", , "mnuEditLowerCase", Image:=61
            .AddItem iP, "&Proper Case", , , "mnuEditFirstCase", Image:=94
            .AddItem iP, "Sentence case.", , , "mnuEditSentenceCase", Image:=93
            .AddItem iP, "iNVERSE cASE", , , "mnuEditChangeCase", Image:=92
            .AddItem iP, "-"
            
            iP2 = .AddItem(iP, "&Convert") 'Conversions
                .AddItem iP2, "Bin -> Hex", , , "mnuConvertBinHex"
                .AddItem iP2, "Bin -> Dec", , , "mnuConvertBinDec"
                .AddItem iP2, "-"
                .AddItem iP2, "Hex -> Bin", , , "mnuConvertHexBin"
                .AddItem iP2, "Hex -> Dec", , , "mnuConvertHexDec"
                .AddItem iP2, "-"
                .AddItem iP2, "Dec -> Bin", , , "mnuConvertDecBin"
                .AddItem iP2, "Dec -> Hex", , , "mnuConvertDecHex"
                
            .AddItem iP, "-"
            iP2 = .AddItem(iP, "&Insert")   ' Insert
                .AddItem iP2, "Insert &Date/Time", , , "mnuEditDateTime", Image:=69
                .AddItem iP2, "Insert &ASCII", "Ctrl+Insert", , "mnuEditInsertASCII", Image:=65
                
            .AddItem iP, "-"
            iP2 = .AddItem(iP, "&Advanced")
                .AddItem iP2, "Auto &indent", "Ctrl+I", , "mnuThotIndent", Image:=68
                .AddItem iP2, "Unitify folder &backslashes", , , "mnuThotUnitifyBackslashes", , , , 72
                
            .AddItem iP, "-"
            .AddItem iP, "Column mode", "Ctrl+K", , "mnuEditColumnMode", , , , 78
            .AddItem iP, "-"
            .AddItem iP, "Code completion help", "Ctrl+Space", , "mnuEditCodeCompletionHelp"
'            .AddItem iP, "-"
'            .AddItem iP, "Preferences...", "F10", , "mnuEditPreferences", Image:=11
        
        'MENU NAVIGATION
        iP = .IndexForKey("mnuNavigation")
            .AddItem iP, "&Search...", "Ctrl+F", , "mnuNavigationSearch", , , , 13
            .AddItem iP, "Search &next", "F3", , "mnuNavigationSearchNext", , , , 14
            .AddItem iP, "Search &prev", "Shift+F3", , "mnuNavigationSearchPrev"
            .AddItem iP, "Search in &files...", , , "mnuNavigationSearchInFiles", , , , 80
            .AddItem iP, "-"
            .AddItem iP, "Search next selected", "Ctrl+F3", , "mnuNavigationSearchNextWord", , , , 89
            .AddItem iP, "Search prev selected", "Ctrl+Shift+F3", , "mnuNavigationSearchPrevWord", , , , 90
            .AddItem iP, "-"
            .AddItem iP, "&Replace...", "Ctrl+H", , "mnuNavigationReplace", Image:=62
            .AddItem iP, "Repla&ce in files...", , , "mnuNavigationReplaceInFiles"
            .AddItem iP, "-"
            .AddItem iP, "Go to line...", "Ctrl+G", , "mnuNavigationGotoLine", , , , 77
            .AddItem iP, "Go to identation", , , "mnuNavigationGotoIdent", , , , 81
            .AddItem iP, "Go to matching &brace", "Ctrl+Shift+B", , "mnuNavigationGotoMatchBrace", , , , 85
            .AddItem iP, "-"
            .AddItem iP, "Go to definition", , , "mnuNavigationGotoDefinition"
            .AddItem iP, "Previous position", , , "mnuNavigationPrevPosition"
            .AddItem iP, "Next position", , , "mnuNavigationNextPosition"
            .AddItem iP, "-"
            .AddItem iP, "Prev function/process", "Alt+Up", , "mnuNavigationPrevFunc", Image:=70
            .AddItem iP, "Next function/process", "Alt+Down", , "mnuNavigationNextFunc", Image:=71
            .AddItem iP, "-"
            iP2 = .AddItem(iP, "&Bookmarks") 'Bookmarks
                .AddItem iP2, "Bookmark &toggle", "Ctrl+F2", , "mnuBookmarkToggle", Image:=24
                .AddItem iP2, "Bookmark &Next", "F2", , "mnuBookmarkNext", Image:=25
                .AddItem iP2, "Bookmark &Prev", "Shift+F2", , "mnuBookmarkPrev", Image:=26
                .AddItem iP2, "&Delete all", , , "mnuBookmarkDel", Image:=27
                .AddItem iP2, "-"
                .AddItem iP2, "&Edit Bookmarks", , , "mnuBookmarkEdit", Image:=95
                .AddItem iP2, "-"
                .AddItem iP2, "&Add this To Do", , , "mnuBookmarkToDo", Image:=29
            
        'MENU PROJECT
        iP = .IndexForKey("mnuProject")
            .AddItem iP, "&Set the active source as main", , , "mnuProjectSetAsMainSource", False, , , 1
            .AddItem iP, "C&lose", , , "mnuProjectClose", False, , , 12
            .AddItem iP, "-"
            .AddItem iP, "&Add Files...", , , "mnuProjectAddFile", False, , , 23
            .AddItem iP, "&Remove current file from project", , , "mnuProjectRemoveFrom", False
            .AddItem iP, "-"
            .AddItem iP, "&Properties", , , "mnuProjectProperties", False, , , 18
            .AddItem iP, "-"
            .AddItem iP, "Show/Hide &Tracker", , , "mnuProjectTracker", False, False, mcs_Icon, 29
            .AddItem iP, "&Developer List", , , "mnuProjectDevList", False, , , 28
      
        'MENU EXECUTE
        iP = .IndexForKey("mnuExecute")
            .AddItem iP, "&Compile this file", "F5", , "mnuExecuteCompileFile", , , , 35
            .AddItem iP, "R&un this file", "Shift+F5", , "mnuExecuteRunFile", , , , 10
            .AddItem iP, "Compile and run this &file", "F6", , "mnuExecuteCompileAndRunFile", , , , 38
            .AddItem iP, "-"
            .AddItem iP, "Compile pro&ject", "F7", , "mnuExecuteBuild", , , , 36
            .AddItem iP, "&Run project", "Shift+F7", , "mnuExecuteRun", , , , 39
            .AddItem iP, "Compile and run &project", "F8", , "mnuExecuteBuildAndRun", , , , 37
            
        'MENU TOOLS
        iP = .IndexForKey("mnuTools")
            .AddItem iP, "&Calculator", "F12", , "mnuToolsCalculator", , , , 58
            .AddItem iP, "-"
            '.AddItem iP, "&Icon changer", "Ctrl+F12", , "mnuToolsIconChanger"
            '.AddItem iP, "-"
            .AddItem iP, "&MS-DOS command", "F9", , "mnuToolsCommand", Image:=64
            .AddItem iP, "C&all last MS-DOS command", "Ctrl+F9", , "mnuToolsLastCommand"
            .AddItem iP, "-"
            .AddItem iP, "Configure &Tools", , , "mnuToolsConfigureTools", , , , 59
            .AddItem iP, "-"
            
'        'MENU THOT
'        iP = .IndexForKey("mnuThot")
'            .AddItem iP, "Auto &indent", "Ctrl+I", , "mnuThotIndent", Image:=68
'            .AddItem iP, "Auto &declare functions", , , "mnuThotDeclareFunctions"
'            .AddItem iP, "-"
'            .AddItem iP, "Unitify folder &backslashes", , , "mnuThotUnitifyBackslashes", , , , 72
'            .AddItem iP, "&Unitify files", , , "mnuThotUnitifyFiles"
'            .AddItem iP, "-"
'            .AddItem iP, "File &truster", , , "mnuThotFileTruster"
'            .AddItem iP, "-"
'            .AddItem iP, "&Migrate from old sources", , , "mnuThotMigrate"
'            .AddItem iP, "-"
'            .AddItem iP, "Add &process", , , "mnuThotAddProcess"
'            .AddItem iP, "Add &function", , , "mnuThotAddFunction"
'            .AddItem iP, "Add &struct", , , "mnuThotAddStruct"
            
                
        'MENU VIEW
        iP = .IndexForKey("mnuView")
            iP2 = .AddItem(iP, "Show / Hide toolbars")
                .AddItem iP2, "Standard", , , "mnuViewToolBarStandard"
                '.AddItem iP2, "Edit", , , "mnuViewToolBarEdit"
            .AddItem iP, "-"
            .AddItem iP, "Project Browser", "Ctrl+1", , "mnuViewProjectBrowser", Image:=44
            .AddItem iP, "Program Inspector", "Ctrl+2", , "mnuViewProgramInspector", Image:=45
            .AddItem iP, "Properties", "Ctrl+3", , "mnuViewProperties", Image:=46
            .AddItem iP, "Compiler output", "Ctrl+4", , "mnuViewCompilerOutput", Image:=47
            .AddItem iP, "Debuger output", "Ctrl+5", , "mnuViewDebugOutput", Image:=48
            .AddItem iP, "Error output", "Ctrl+6", , "mnuViewErrorOutput", Image:=49
            .AddItem iP, "-"
            .AddItem iP, "Full screen", "F11", , "mnuViewFullScreen", True, Image:=50
            
        'MENU HELP
        iP = .IndexForKey("mnuHelp")
            .AddItem iP, "Language Help (Spanish only)", , , "mnuHelpIndex", Image:=66
            .AddItem iP, "-"
            .AddItem iP, "Fenix Wiki (Spanish only)", , , "mnuHelpWiki"
            .AddItem iP, "Fenix Forum (Spanish only)", , , "mnuHelpFenixForum"
            .AddItem iP, "-"
            .AddItem iP, "Bennu Wiki", , , "mnuHelpBennuWiki"
            .AddItem iP, "Bennu Forum", , , "mnuHelpBennuForum"
            .AddItem iP, "-"
            .AddItem iP, "FBMX project page", , , "mnuHelpProjectPage"
            .AddItem iP, "-"
            .AddItem iP, "&About FBMX...", , , "mnuHelpAbout", Image:=82
    End With
    
    CreateToolsMenu
End Sub

Private Function CreateToolBars()
    With tbrMain
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList ImgList1.hIml, CTBImageListNormal
        
        .CreateToolbar 16, False, False, False
        
        .AddButton "New", 0, , , , CTBDropDown, "New"
        .AddButton "Open", 1, , , , CTBDropDown, "Open"
        .AddButton "Save", 2, , , , CTBNormal, "Save"
        .AddButton "Save All", 30, , , , CTBNormal, "SaveAll"
        .AddButton "Close", 32, , , , CTBNormal, "Close"
        .AddButton "Close All", 33, , , , CTBNormal, "CloseAll"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Cut", 4, , , , CTBNormal, "Cut"
        .AddButton "Copy", 3, , , , CTBNormal, "Copy"
        .AddButton "Paste", 5, , , , CTBNormal, "Paste"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Undo", 6, , , , CTBNormal, "Undo"
        .AddButton "Redo", 7, , , , CTBNormal, "Redo"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Compile and run this file", 37, , , , CTBNormal, "CompileAndRunFile"
        .AddButton "Compile and run project", 36, , , , CTBNormal, "BuildAndRun"
        .AddButton "", -1, , , , CTBSeparator
        .AddButton "Preferences", 10, , , , CTBNormal, "Preferences"
    End With

    With cRebar
        ' Background bitmap
        If A_Bitmaps Then .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        
        .CreateRebar picHolder.Hwnd
        

        'Add
        .AddBandByHwnd tbrMain.Hwnd, , False, False, "MainBar"

        'The optimal height of the buttons should be 22px but for an strange
        'reason, if we set this minimum height to 22, the toolbar appears
        'overlaped by tabdock panels when we restore the window (after having being
        'minimized) with windowState = 0 (normal).
        'So we choose 23 :)
        .BandChildMinHeight(0) = 23
        .BandChildMinHeight(1) = 23

        'Adjust band widths
'        .BandChildMinWidth(.BandIndexForData("MainBar")) = CalculateToolBarWidth(tbrMain)
'        .BandChildIdealWidth(.BandIndexForData("MainBar")) = CalculateToolBarWidth(tbrMain)
    End With
End Function

'Private Sub LoadDockedFormConfiguration()
'    Dim state As tdDockedState, align As tdAlignProperty
'    Dim style As tdDockStyles, percent As Integer, expanded As Boolean
'    Dim dockedForm As TDockForm
'    'Properties form
'    state = tdDocked
'    align = tdAlignRight
'    style = tdDockRight Or tdDockLeft Or tdDockFloat
'    percent = 30
'    expanded = True
'    Call LoadDockedFormConf(frmProperties.name, state, align, style, percent, expanded)
'
'    'Project browser
'    state = tdDocked
'    align = tdAlignLeft
'    style = tdDockRight Or tdDockLeft Or tdDockFloat
'    percent = 30
'    expanded = True
'    Call LoadDockedFormConf(frmProjectBrowser.name, state, align, style, percent, expanded)
'
'    'Program inspector
'    state = tdDocked
'    align = tdAlignRight
'    style = tdDockRight Or tdDockFloat Or tdDockLeft
'    percent = 30
'    expanded = True
'    Call LoadDockedFormConf(frmProgramInspector.name, state, align, style, percent, expanded)
'
'    'Compiler output
'    state = tdDocked
'    align = tdAlignBottom
'    style = tdDockBottom Or tdDockFloat Or tdDockTop
'    percent = 70
'    expanded = True
'    Call LoadDockedFormConf(frmOutput.name, state, align, style, percent, expanded)
'
'    'Debug output
'    state = tdDocked
'    align = tdAlignBottom
'    style = tdDockBottom Or tdDockFloat Or tdDockTop
'    percent = 70
'    expanded = True
'    Call LoadDockedFormConf(frmDebug.name, state, align, style, percent, expanded)
'
'    'Error output
'    state = tdDocked
'    align = tdAlignBottom
'    style = tdDockBottom Or tdDockFloat Or tdDockTop
'    percent = 70
'    expanded = True
'    Call LoadDockedFormConf(frmErrors.name, state, align, style, percent, expanded)
'End Sub
Private Sub CreateDockableInterface()
    'Create dockable panels
    With TabDock
        .GrabMain Me.Hwnd
        .AutoExpand = False
        .AutoShowCaptionOnCollapse = True
        .CaptionStyle = tdCaptionVSNet
        
        .CollapseInterval = 3000
        
        .PanelResizeRight = True
        .PanelResizeLeft = True
        .PanelResizeBottom = True
        .PanelResizeTop = True
        
        .PanelBottomDockFormResize = True
        
        .SmartSizing = True
        .UseITDockMoveEvents = True
    End With
    
    TabDock.AddForm frmProjectBrowser, tdDocked, tdAlignLeft, frmProjectBrowser.name, _
                tdDockRight Or tdDockFloat Or tdDockLeft _
                , 30, False, True
                
    TabDock.AddForm frmProperties, tdDocked, tdAlignLeft, frmProperties.name, _
                tdDockRight Or tdDockFloat Or tdDockLeft _
                , 30, False, True
                
    TabDock.AddForm frmProgramInspector, tdDocked, tdAlignRight, frmProgramInspector.name, _
                tdDockRight Or tdDockFloat Or tdDockLeft _
                , 30, False, True
                
    TabDock.AddForm frmOutput, tdDocked, tdAlignBottom, frmOutput.name, _
                tdDockBottom Or tdDockTop Or tdDockFloat _
                , 70, False, True
                
    TabDock.AddForm frmDebug, tdDocked, tdAlignBottom, frmDebug.name, _
                tdDockBottom Or tdDockTop Or tdDockFloat _
                , 70, False, True
                
    TabDock.AddForm frmErrors, tdDocked, tdAlignBottom, frmErrors.name, _
                tdDockBottom Or tdDockTop Or tdDockFloat _
                , 70, False, True
    
    TabDock.Show
    
    Dim i As Integer
    Dim DockedForm As TDockForm
    Dim state As tdDockedState, align As tdAlignProperty, Style As tdDockStyles
    Dim percent As Long, expanded As Boolean, isVisible As Boolean
    'TRICK. Hide all forms
    For i = 1 To TabDock.DockedForms.count
        TabDock.FormHide TabDock.DockedForms(i).Key
    Next
    With Ini
        For i = 1 To TabDock.DockedForms.count
            Set DockedForm = TabDock.DockedForms(i).Object
            state = DockedForm.state
            expanded = True
            isVisible = True
            percent = DockedForm.percent
            align = DockedForm.Panel.align
            LoadDockedFormConf DockedForm.Key, state, align, Style, percent, expanded, isVisible

            If isVisible = True Then
                TabDock.FormShow DockedForm.Key
            End If
            'If state = tdUndocked Then
'                TabDock.FormUndock dockedForm.Key
'            End If
        Next
    End With
'
'    If expanded = False Then
'        DockedForm.Panel.Panel_Collapse
'    End If
End Sub

'-------------------------------------------------------------------------------
'MENUS
'-------------------------------------------------------------------------------

Private Sub cMenu_Click(ByVal Index As Long)

    Select Case cMenu.ItemKey(Index)
    Case "mnuFileExit":                     Unload Me
    Case "mnuFileNewFile":                  Call mnuFileNewFile
    Case "mnuFileNewProject":               Call mnuFileNewProject
    Case "mnuFileNewSource":                Call mnuFileNewSource
    Case "mnuFileNewMap":                   Call mnuFileNewMap
    Case "mnuFileNewFpg":                   Call mnuFileNewFpg
    Case "mnuFileNewImp":                   Call mnuFileNewImp
    Case "mnuFileOpenFile":                 Call mnuFileOpenFile
    Case "mnuFileOpenProject":              Call mnuFileOpenProject
    Case "mnuFileOpenSource":               Call mnuFileOpenSource
    Case "mnuFileOpenMap":                  Call mnuFileOpenMap
    Case "mnuFileOpenFpg":                  Call mnuFileOpenFpg
    Case "mnuFileOpenFnt":                  Call mnuFIleOpenFnt
    Case "mnuFileOpenImp":                  Call mnuFileOpenImp
    Case "mnuFileOpenSong":                 Call mnuFileOpenSong
    Case "mnuFileClose":                    Call mnuFileClose
    Case "mnuFileCloseAll":                 Call mnuFileCloseAll
    Case "mnuFileCloseProject":             Call CloseProject
    Case "mnuFileSave":                     Call mnuFileSave
    Case "mnuFileSaveAll":                  Call mnuFileSaveAll
    Case "mnuFileSaveAs":                   Call mnuFileSaveAs
    Case "mnuFilePrint":                    Call mnuFilePrint
    
    Case "mnuEditUndo":                     Call mnuEditUndo
    Case "mnuEditRedo":                     Call mnuEditRedo
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
    Case "mnuEditDeleteWord":               Call mnuEditDeleteWord
    Case "mnuEditDeleteWordFromCursor":     Call mnuEditDeleteWordFromCursor
    Case "mnuEditDownLine":                 Call mnuEditDownLine
    Case "mnuEditTab":                      Call mnuEditTab
    Case "mnuEditUnTab":                    Call mnuEditUnTab
    Case "mnuEditComment":                  Call mnuEditComment
    Case "mnuEditUnComment":                Call mnuEditUnComment
    Case "mnuEditUpperCase":                Call mnuEditUpperCase
    Case "mnuEditLowerCase":                Call mnuEditLowerCase
    Case "mnuEditChangeCase":               Call mnuEditChangeCase
    Case "mnuEditFirstCase":                Call mnuEditFirstCase
    Case "mnuEditSentenceCase":             Call mnuEditSentenceCase
    Case "mnuEditColumnMode":               Call mnuEditColumnMode
    Case "mnuEditCodeCompletionHelp":       Call mnuEditCodeCompletionHelp
    
    Case "mnuConvertBinHex":                Call mnuConvertBinHex
    Case "mnuConvertBinDec":                Call mnuConvertBinDec
    Case "mnuConvertHexBin":                Call mnuConvertHexBin
    Case "mnuConvertHexDec":                Call mnuConvertHexDec
    Case "mnuConvertDecBin":                Call mnuConvertDecBin
    Case "mnuConvertDecHex":                Call mnuConvertDecHex
    
    Case "mnuEditDateTime":                 Call mnuEditDateTime
    Case "mnuEditInsertASCII":              Call mnuEditInsertASCII
    Case "mnuEditPreferences":              Call mnuEditPreferences
            
    Case "mnuNavigationSearch":             Call mnuNavigationSearch
    Case "mnuNavigationSearchNext":         Call mnuNavigationSearchNext
    Case "mnuNavigationSearchPrev":         Call mnuNavigationSearchPrev
    Case "mnuNavigationSearchNextWord":     Call mnuNavigationSearchNextWord
    Case "mnuNavigationSearchPrevWord":     Call mnuNavigationSearchPrevWord
    Case "mnuNavigationReplace":            Call mnuNavigationReplace
    Case "mnuNavigationGotoLine":           Call mnuNavigationGoToLine
    Case "mnuNavigationGotoIdent":          Call mnuNavigationGoToIdent
    Case "mnuNavigationNextFunc":           Call mnuNavigationNextFunc
    Case "mnuNavigationPrevFunc":           Call mnuNavigationPrevFunc
    Case "mnuNavigationPrevPosition":       Call mnuNavigationPrevPosition
    Case "mnuNavigationNextPosition":       Call mnuNavigationNextPosition
    Case "mnuNavigationGotoDefinition":     Call mnuNavigationGoToDefinition
    Case "mnuNavigationGotoMatchBrace":     Call mnuNavigationGotoMatchBrace
    
    Case "mnuBookmarkToggle":               Call mnuBookmarkToggle
    Case "mnuBookmarkNext":                 Call mnuBookmarkNext
    Case "mnuBookmarkPrev":                 Call mnuBookmarkPrev
    Case "mnuBookmarkDel":                  Call mnuBookmarkDel
    Case "mnuBookmarkToDo":                 Call mnuBookmarkToDo
    Case "mnuBookmarkEdit":                 Call mnuBookmarkEdit

    Case "mnuExecuteCompileFile":           Call mnuExecuteCompileFile
    Case "mnuExecuteRunFile":               Call mnuExecuteRunFile
    Case "mnuExecuteCompileAndRunFile":     Call mnuExecuteCompileAndRunFile
    Case "mnuExecuteBuild":                 Call mnuExecuteBuild
    Case "mnuExecuteRun":                   Call mnuExecuteRun
    Case "mnuExecuteBuildAndRun":           Call mnuExecuteBuildAndRun
    
    Case "mnuProjectProperties":            Call mnuProjectProperties
    Case "mnuProjectTracker":               Call mnuProjectTracker
    Case "mnuProjectDevList":               Call mnuProjectDevList
    Case "mnuProjectSetAsMainSource":       Call mnuProjectSetAsMainSource
    Case "mnuProjectClose":                 Call CloseProject
    Case "mnuProjectAddFile":               Call mnuProjectAddFile
    Case "mnuProjectRemoveFrom":            Call mnuProjectRemoveFile
    
    Case "mnuToolsCalculator":              Call mnuToolsCalculator
    Case "mnuToolsIconChanger":             Call mnuToolsIconChanger
    Case "mnuToolsCommand":                 Call mnuToolsCommand
    Case "mnuToolsLastCommand":             Call mnuToolsLastCommand
    Case "mnuToolsConfigureTools":          Call mnuToolsConfigureTools
    
    Case "mnuThotIndent":                   Call mnuThotIndent
    Case "mnuThotDeclareFunctions":         Call mnuThotDeclareFunctions
    Case "mnuThotUnitifyBackslashes":       Call mnuThotUnitifyBackslashes
    Case "mnuThotUnitifyFiles":             Call mnuThotUnitifyFiles
    Case "mnuThotFileTruster":              Call mnuThotFileTruster
    Case "mnuThotMigrate":                  Call mnuThotMigrate
    Case "mnuThotAddProcess":               Call mnuThotAddProcess
    Case "mnuThotAddFunction":              Call mnuThotAddFunction
    Case "mnuThotAddStruct":                Call mnuThotAddStruct
    
    Case "mnuViewToolBarStandard":          Call mnuViewToolBarStandard
    Case "mnuViewProjectBrowser":           Call mnuViewProjectBrowser
    Case "mnuViewProgramInspector":         Call mnuViewProgramInspector
    Case "mnuViewProperties":               Call mnuViewProperties
    Case "mnuViewCompilerOutput":           Call mnuViewCompilerOutput
    Case "mnuViewDebugOutput":              Call mnuViewDebugOutput
    Case "mnuViewErrorOutput":              Call mnuViewErrorOutput
    Case "mnuViewFullScreen":               Call mnuViewFullScreen
    
    Case "mnuHelpIndex":                    Call mnuHelpIndex
    Case "mnuHelpWiki":                     Call mnuHelpWiki
    Case "mnuHelpFenixForum":               Call mnuHelpFenixForum
    Case "mnuHelpBennuWiki":                Call mnuHelpBennuWiki
    Case "mnuHelpBennuForum":               Call mnuHelpBennuForum
    Case "mnuHelpProjectPage":              Call mnuHelpProjectPage
    Case "mnuHelpAbout":                    Call mnuHelpAbout
    
    Case Else
        'Recents (se pone al final porque puede darse el caso de que algun menu empiece por mnuRec)
        If cMenu.ItemKey(Index) Like "mnuRec*#" Then
            mnuFileRecentOpen cMenu.ItemCaption(Index)
        ElseIf cMenu.ItemKey(Index) Like "TOOL*" Then
            mnuToolsRunTool CInt(Right(cMenu.ItemKey(Index), Len(cMenu.ItemKey(Index)) - 4))
        Else
            ' si es un plugin
            If cMenu.ItemParentID(Index) = cMenu.IndexForKey("mnuPlugins") Then
                RunPlugin (cMenu.ItemKey(Index))
            End If
        End If
   End Select
End Sub

Private Sub InitKey_Accelerators()
    With ctlAccelerator
      .AddAccelerator vbKeyUp, vbCtrlMask + vbShiftMask, "Line up"
      .AddAccelerator vbKeyDown, vbCtrlMask + vbShiftMask, "Line down"
      .AddAccelerator vbKeyUp, vbAltMask, "Prev function"
      .AddAccelerator vbKeyDown, vbAltMask, "Next function"
      .AddAccelerator vbKeyInsert, vbCtrlMask, "Insert ASCII"
      .AddAccelerator vbKeyL, vbCtrlMask + vbShiftMask, "Select line"
      .AddAccelerator vbKeyW, vbCtrlMask + vbShiftMask, "Select word"
      .AddAccelerator vbKeyF1, vbShiftMask, "Wiki Help"
      .Enabled = True
   End With
End Sub

Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)
Dim sActive As String
   
   ' Show the accelerator we got and also whether the form is currently
   ' the system active form:
   If Not ctlAccelerator.IsActive Then
      sActive = "[InActive]"
   End If
    ' run the menu action
    Select Case (ctlAccelerator.Key(nIndex))
        Case "Line up"
            mnuEditUpLine
        Case "Line down"
            mnuEditDownLine
        Case "Prev function"
            mnuNavigationPrevFunc
        Case "Next function"
            mnuNavigationNextFunc
        Case "Insert ASCII"
            mnuEditInsertASCII
        Case "Select line"
            mnuEditSelectLine
        Case "Select word"
            mnuEditSelectWord
        Case "Wiki Help"
            mnuHelpWikiToken
   End Select
End Sub

Private Sub StatusBar_DrawItem(ByVal lHdc As Long, ByVal iPanel As Long, ByVal lLeftPixels As Long, ByVal lTopPixels As Long, ByVal lRightPixels As Long, ByVal lBottomPixels As Long)
    Dim tr As RECT
    Dim debugAvailable As Boolean
    Dim errorsAvailable As Boolean
    
    tr.Left = lLeftPixels + 2
    tr.Right = lRightPixels - 1
    tr.Top = lTopPixels
    tr.Bottom = lBottomPixels - 1
    
    debugAvailable = IIf(frmDebug.txtOutput <> "", True, False)
    errorsAvailable = IIf(frmErrors.txtOutput <> "", True, False)
    ImgList1.DrawImage IIf(debugAvailable, 48, 51), lHdc, tr.Left, tr.Top
    ImgList1.DrawImage IIf(errorsAvailable, 49, 52), lHdc, tr.Left + 20, tr.Top
End Sub

'-------------------------------------------------------------------------------
'TOOLBARS
'-------------------------------------------------------------------------------

Private Sub TbrMain_ButtonClick(ByVal lButton As Long)
    Select Case tbrMain.ButtonKey(lButton)
        Case "New":                 mnuFileNewFile
        Case "Open":                mnuFileOpenFile
        Case "Close":               mnuFileClose
        Case "CloseAll":            mnuFileCloseAll
        Case "Save":                mnuFileSave
        Case "SaveAll":             mnuFileSaveAll
        Case "Undo":                mnuEditUndo
        Case "Redo":                mnuEditRedo
        Case "Cut":                 mnuEditCut
        Case "Copy":                mnuEditCopy
        Case "Paste":               mnuEditPaste
        Case "CompileAndRunFile":   mnuExecuteCompileAndRunFile
        Case "BuildAndRun":         mnuExecuteBuildAndRun
        Case "Preferences":         frmPreferences.Show vbModeless, Me
    End Select
End Sub

Private Sub TbrMain_DropDownPress(ByVal lButton As Long)
    Dim X As Long, Y As Long
    Dim lIndex As Long
    tbrMain.GetDropDownPosition lButton, X, Y
    
    Select Case tbrMain.ButtonKey(lButton)
        Case "New":
            Call cMenu.PopupMenu("mnuFileNew", (Me.Left + X + 50) / Screen.TwipsPerPixelX, (Me.Top + Y + tbrMain.Height + 310) / Screen.TwipsPerPixelY, TPM_VERNEGANIMATION)
            
        Case "Open":
            Call cMenu.PopupMenu("mnuFileOpen", (Me.Left + X + 50) / Screen.TwipsPerPixelX, (Me.Top + Y + tbrMain.Height + 310) / Screen.TwipsPerPixelY, TPM_VERNEGANIMATION)
    End Select
End Sub

Private Sub cReBar_ChevronPushed(ByVal wID As Long, ByVal lLeft As Long, _
                        ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
    Dim v As Variant
   v = cRebar.BandData(wID)
   If Not IsMissing(v) Then
      Select Case v
        Case "MainBar"
            tbrMain.ChevronPress lRight \ Screen.TwipsPerPixelX + 1, lTop \ Screen.TwipsPerPixelY
      End Select
   End If
End Sub

Private Sub cReBar_HeightChanged(lNewHeight As Long)
    cRebar.RebarSize
   If picHolder.align = 1 Or picHolder.align = 2 Then
      picHolder.Height = lNewHeight * Screen.TwipsPerPixelY
   Else
      picHolder.Width = lNewHeight * Screen.TwipsPerPixelY
   End If
End Sub

Private Sub picHolder_Resize()
    cRebar.RebarSize
    If picHolder.align = 1 Or picHolder.align = 2 Then
        picHolder.Height = cRebar.RebarHeight * Screen.TwipsPerPixelY
    Else
        picHolder.Width = cRebar.RebarHeight * Screen.TwipsPerPixelY
    End If
End Sub


'-------------------------------------------------------------------------------
'SUBCLASSING
'-------------------------------------------------------------------------------

Private Property Get ISubclass_MsgResponse() As TabDock.EMsgResponse

End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As TabDock.EMsgResponse)
    'Tell the subclasser what to do for this message
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal Hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim mmiT As MINMAXINFO

    Select Case iMsg
        Case WM_GETMINMAXINFO
            'Copy parameter to local variable for processing
            CopyMemory mmiT, ByVal lParam, LenB(mmiT)
            'Minimun width and height for sizing
            mmiT.ptMinTrackSize.X = MinWidth
            mmiT.ptMinTrackSize.Y = MinHeight
            'Copy modified results back to parameter
            CopyMemory ByVal lParam, mmiT, LenB(mmiT)
    End Select
End Function
