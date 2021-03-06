Attribute VB_Name = "modMenuActions"
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Dar�o Cutillas)
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
Option Base 0

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal Hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function WinExec Lib "kernel32.dll" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Const SW_SHOWDEFAULT As Long = 10
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_BORDER As Long = &H800000
Private Const GWL_STYLE As Long = -16

Private Const MSG_MNUACTIONS_MAINSOURCENOTDEFINED As String = "Main source of the project has not been defined yet"

Private fDoc As frmDoc, fMap As frmMap, fImp As frmImport

Private Enum NewTypeConstants
    NT_NONE
    NT_PROJECT
    NT_SOURCE
    NT_MAP
    NT_FPG
    NT_PAL
    NT_IMP
    NT_FNT
    NT_LIST
End Enum

' For createImp
Private Type modules
    m(40) As String
    dat(40) As Byte
    n_Fun(1000) As Long
End Type
Private fileModules(100) As String
Private batCreate As Long
Private pa_scrb As Long
Private num As Long
'    string module[] = "mod_blendop", "mod_cd", "mod_dir", "mod_draw", "mod_effects", "mod_file", "mod_flic",
'    "mod_grproc", "mod_joy", "mod_key", "mod_m7", "mod_map", "mod_math", "mod_mem", "mod_mouse", "mod_path",
'    "mod_proc", "mod_rand", "mod_regex",
'    "mod_say", "mod_screen", "mod_scroll", "mod_sort", "mod_sound", "mod_string", "mod_sys", "mod_text",
'    "mod_time", "mod_timers", "mod_video", "mod_wm", "mod_ffi";
'    _dat[] = 10, 10, 7, 17, 4, 16, 16, 6, 32, 1, 4, 107, 14, 22, 1, 3, 8, 2, 4, 2, 12, 5, 5, 25, 22, 2, 11, 3, 1, 7, 8, 2;
'    n_fun[390];
'    string fileModules[50];
'    batCreate;
'    pa_scrb;
'    num;
'    goIncludes;

Private m_NewType As Integer
Private m_NewAddToProject As Boolean

Private Property Get ActiveFileForm() As IFileForm
    Set ActiveFileForm = frmMain.ActiveFileForm
End Property

Private Property Get ActiveForm() As Form
    Set ActiveForm = frmMain.ActiveForm
End Property

'Once readed, it m_newaddtoproject is set to false
Public Property Get NewAddToProject() As Boolean
    NewAddToProject = m_NewAddToProject
    m_NewAddToProject = False
End Property

Public Property Let NewAddToProject(newVal As Boolean)
    m_NewAddToProject = newVal
End Property

Public Property Let newType(sType As String)
    Select Case sType
        Case "PROJECT"
            m_NewType = NT_PROJECT
        Case "SOURCE"
            m_NewType = NT_SOURCE
        Case "MAP"
            m_NewType = NT_MAP
        Case "FPG"
            m_NewType = NT_FPG
        Case "PAL"
            m_NewType = NT_PAL
        Case "FNT"
            m_NewType = NT_FNT
        Case "IMP"
            m_NewType = NT_IMP
        Case "LIST"
            m_NewType = NT_LIST
        Case Else
            MsgBox "Error en modMenuActions.newType"
    End Select
End Property

'-------------------------------------------------------------------------------
'START HELP FUNCTIONS
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'END HELP FUNCTIONS
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'START FILE MENU
'-------------------------------------------------------------------------------
Public Sub mnuFileNewFile()
    m_NewAddToProject = False
    frmNewFile.Show 1, frmMain
    
    Select Case m_NewType
        Case NT_NONE
            Exit Sub
        Case NT_PROJECT
            NewProject
        Case NT_SOURCE
            Call NewFileForm(FF_SOURCE)
        Case NT_MAP
            Call NewFileForm(FF_MAP)
        Case NT_FPG
            Call NewFileForm(FF_FPG)
        Case NT_IMP
            Call NewFileForm(FF_IMP)
'        Case NT_PAL
'            Call NewFileForm(FF_PAL)
'        Case NT_FNT
'            Call NewFileForm(FF_FNT)
        Case Else
            MsgBox "Sorry. Option not available yet", vbInformation
    End Select
End Sub

Public Sub mnuFileNewProject()
    NewProject
End Sub

Public Sub mnuFileNewSource()
    NewFileForm FF_SOURCE
End Sub

Public Sub mnuFileNewMap()
    NewFileForm FF_MAP
End Sub

Public Sub mnuFileNewFpg()
    NewFileForm FF_FPG
End Sub

Public Sub mnuFileNewImp()
    NewFileForm FF_IMP
End Sub

Public Sub mnuFileOpenFile()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles(), getFilter("READABLE_FILES"), True) > 0 Then
        OpenMultipleFiles sFiles()
    End If
End Sub

Public Sub mnuFileOpenSource()
    OpenFileOfFileForm FF_SOURCE
End Sub

Public Sub mnuFileOpenMap()
    OpenFileOfFileForm FF_MAP
End Sub

Public Sub mnuFileOpenFpg()
    OpenFileOfFileForm FF_FPG
End Sub

Public Sub mnuFIleOpenFnt()
    OpenFileOfFileForm FF_FNT
End Sub

Public Sub mnuFileOpenSong()
    OpenSong
End Sub

Public Sub mnuFileOpenImp()
    OpenFileOfFileForm FF_IMP
End Sub

Public Sub mnuFileOpenProject()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles, getFilter("FBP"), False, False) > 0 Then
        OpenProject sFiles(0)
    End If
End Sub

Public Sub mnuFileSave()
    If Not ActiveFileForm Is Nothing Then
        SaveFileOfFileForm ActiveFileForm, False
    End If
End Sub

Public Sub mnuFileSaveAll()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            SaveFileOfFileForm ff
        End If
    Next
End Sub

Public Sub mnuFileCloseSources()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_SOURCE Then
                Unload f
            End If
        End If
    Next
    frmMain.RefreshTabs
End Sub

Public Sub mnuFileCloseFnts()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_FNT Then
                Unload f
            End If
        End If
    Next
    frmMain.RefreshTabs
End Sub

Public Sub mnuFileCloseImps()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_IMP Then
                Unload f
            End If
        End If
    Next
    frmMain.RefreshTabs
End Sub

Public Sub mnuFileCloseMaps()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_MAP Then
                Unload f
            End If
        End If
    Next
    frmMain.RefreshTabs
End Sub
Public Sub mnuFileCloseFpgs()
    Dim f As Form, ff As IFileForm
    
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_FPG Then
                Unload f
            End If
        End If
    Next
    frmMain.RefreshTabs
End Sub

Public Sub mnuFileClose()
    Dim f As Form
    Dim tr As RECT
    If Not ActiveFileForm Is Nothing Then
        Set f = ActiveForm
        GetWindowRect f.Hwnd, tr
        tr.Top = 0
        LockWindowUpdate f.Hwnd
        If Not f Is Nothing Then Unload f
        LockWindowUpdate False
        InvalidateRect frmMain.Hwnd, tr, 0 'Refreshes the tab
    End If
End Sub

Public Sub mnuFileCloseAll()
    Dim lastHwnd As Long
    lastHwnd = -1
    Do Until (ActiveFileForm Is Nothing)
        If lastHwnd = ActiveForm.Hwnd Then Exit Do 'If the form still is visible, the user selected cancel
        lastHwnd = ActiveForm.Hwnd
        mnuFileClose
    Loop
End Sub

Public Sub mnuFileSortByName()


End Sub

Public Sub mnuFileSortByType()
    frmMain.sortTabs
End Sub

Public Sub mnuFileSaveAs()
    If Not ActiveFileForm Is Nothing Then
        SaveFileOfFileForm ActiveFileForm, True
    End If
End Sub

Public Sub mnuFileRecentOpen(sFile As String)
    If Not Dir(sFile) = "" Then
        OpenFileByExt sFile
    Else ' Doesn't exists the file
        MsgBox "File not found!", vbCritical
    End If
End Sub

Public Sub mnuFilePrint()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.PrintContents cmPrnHDC, cmPrnPromptDlg
        End If
    End If
End Sub

Public Sub mnuFileConfPrint()
    MsgBox "Function not working!"
End Sub
'-------------------------------------------------------------------------------
'END FILE MENU
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'START EDIT MENU
'-------------------------------------------------------------------------------
Public Sub mnuEditUndo()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.CanUndo Then
                fDoc.cs.Undo
            End If
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            If fImp.cs.CanUndo Then
                fImp.cs.Undo
            End If
        ElseIf ActiveFileForm.Identify = FF_LIST Then
            'Set fList = ActiveForm
        End If
    End If
End Sub

Public Sub mnuEditRedo()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.CanRedo Then
                fDoc.cs.Redo
            End If
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            If fImp.cs.CanRedo Then
                fImp.cs.Redo
            End If
        ElseIf ActiveFileForm.Identify = FF_LIST Then
            'Set fList = ActiveForm
        End If
    End If
End Sub

Public Sub mnuEditCut()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.CanCut Then
                fDoc.cs.Cut
            End If
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            If fImp.cs.CanCut Then
                fImp.cs.Cut
            End If
        ElseIf ActiveFileForm.Identify = FF_LIST Then
            'Set fList = ActiveForm
        End If
    End If
End Sub

Public Sub mnuEditCopy()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.CanCopy Then
                fDoc.cs.Copy
            End If
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            If fImp.cs.CanCopy Then
                fImp.cs.Copy
            End If
        ElseIf ActiveFileForm.Identify = FF_LIST Then
            'Set fList = ActiveForm
        End If
    End If
End Sub

Public Sub mnuEditPaste()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.CanPaste Then
                fDoc.cs.Paste
            End If
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            If fImp.cs.CanPaste Then
                fImp.cs.Paste
            End If
        ElseIf ActiveFileForm.Identify = FF_LIST Then
            'Set fList = ActiveForm
        End If
    End If
End Sub

Public Sub mnuEditDateTime()
    On Error Resume Next
    Dim timedate As String
    Dim Pos As New CodeSense.position
    timedate = Date & "/" & time
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Pos.ColNo = fDoc.rangoActual.StartColNo
            Pos.LineNo = fDoc.rangoActual.StartLineNo
            fDoc.cs.DeleteSel
            fDoc.cs.InsertText timedate, Pos
        End If
    End If
End Sub

Public Sub mnuEditInsertASCII()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            frmInsertASCII.Show 1, frmMain
        End If
    End If
End Sub

Public Sub mnuEditSelectAll()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdSelectAll
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdSelectAll
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdSelectAll
        End If
    End If
End Sub

Public Sub mnuEditSelectWord()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdWordLeft
            fDoc.cs.ExecuteCmd cmCmdWordEndRightExtend
        End If
    End If
End Sub

Public Sub mnuEditSelectLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdSelectLine
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdSelectLine
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdSelectLine
        End If
    End If
End Sub
Public Sub mnuEditDuplicateLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBeginUndo
                fDoc.cs.ExecuteCmd cmCmdSelectLine
                fDoc.cs.ExecuteCmd cmCmdCopy
                fDoc.cs.ExecuteCmd cmCmdNewLine
                fDoc.cs.ExecuteCmd cmCmdLineUp
                fDoc.cs.ExecuteCmd cmCmdPaste
                fDoc.cs.ExecuteCmd cmCmdLineDown
                fDoc.cs.ExecuteCmd cmCmdPaste
            fDoc.cs.ExecuteCmd cmCmdEndUndo
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdBeginUndo
                fImp.cs.ExecuteCmd cmCmdSelectLine
                fImp.cs.ExecuteCmd cmCmdCopy
                fImp.cs.ExecuteCmd cmCmdNewLine
                fImp.cs.ExecuteCmd cmCmdLineUp
                fImp.cs.ExecuteCmd cmCmdPaste
                fImp.cs.ExecuteCmd cmCmdLineDown
                fImp.cs.ExecuteCmd cmCmdPaste
            fImp.cs.ExecuteCmd cmCmdEndUndo
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdBeginUndo
'                fDoc.cs.ExecuteCmd cmCmdSelectLine
'                fDoc.cs.ExecuteCmd cmCmdCopy
'                fDoc.cs.ExecuteCmd cmCmdNewLine
'                fDoc.cs.ExecuteCmd cmCmdLineUp
'                fDoc.cs.ExecuteCmd cmCmdPaste
'                fDoc.cs.ExecuteCmd cmCmdLineDown
'                fDoc.cs.ExecuteCmd cmCmdPaste
'            fDoc.cs.ExecuteCmd cmCmdEndUndo
        End If
    End If
End Sub

Public Sub mnuEditDeselect()
    
    Dim lineText As Integer
    Dim curLine As Integer
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdClearSelection
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdClearSelection
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdClearSelection
        End If
    End If
End Sub

Public Sub mnuEditDeleteLine()
    
    Dim lineText As Integer
    Dim curLine As Integer
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdLineDelete
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdLineDelete
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdLineDelete
        End If
    End If
End Sub

Public Sub mnuEditClearLine()
    
    Dim lineText As Integer
    Dim curLine As Integer
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBeginUndo
                fDoc.cs.ExecuteCmd cmCmdSelectLine
                fDoc.cs.ExecuteCmd cmCmdLineDeleteToEnd
                fDoc.cs.ExecuteCmd cmCmdLineDeleteToStart
            fDoc.cs.ExecuteCmd cmCmdEndUndo
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdBeginUndo
                fImp.cs.ExecuteCmd cmCmdSelectLine
                fImp.cs.ExecuteCmd cmCmdLineDeleteToEnd
                fImp.cs.ExecuteCmd cmCmdLineDeleteToStart
            fImp.cs.ExecuteCmd cmCmdEndUndo
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdBeginUndo
'                fDoc.cs.ExecuteCmd cmCmdSelectLine
'                fDoc.cs.ExecuteCmd cmCmdLineDeleteToEnd
'                fDoc.cs.ExecuteCmd cmCmdLineDeleteToStart
'            fDoc.cs.ExecuteCmd cmCmdEndUndo
        End If
    End If
End Sub

Public Sub mnuEditUpLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdLineTranspose
            fDoc.cs.ExecuteCmd cmCmdLineUp
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdLineTranspose
            fImp.cs.ExecuteCmd cmCmdLineUp
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdLineTranspose
'            fDoc.cs.ExecuteCmd cmCmdLineUp
        End If
    End If
End Sub

Public Sub mnuEditDownLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdLineDown
            fDoc.cs.ExecuteCmd cmCmdLineTranspose
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdLineDown
            fImp.cs.ExecuteCmd cmCmdLineTranspose
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdLineDown
'            fDoc.cs.ExecuteCmd cmCmdLineTranspose
        End If
    End If
End Sub
Public Sub mnuEditDeleteWordFromCursor()
'deletes the word starting from cursor position
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdWordDeleteToEnd
        End If
    End If
End Sub

Public Sub mnuEditDeleteWord()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdWordDeleteToEnd
            fDoc.cs.ExecuteCmd cmCmdWordDeleteToStart
        End If
    End If
End Sub

Public Sub mnuNavigationSearch()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFind
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFind
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdFind
        End If
    End If
End Sub

Public Sub mnuNavigationSearchNext()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindNext
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFindNext
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdFindNext
        End If
    End If
End Sub

Public Sub mnuNavigationSearchPrev()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindPrev
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFindPrev
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdFindPrev
        End If
    End If
End Sub

Public Sub mnuNavigationSearchNextWord()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindNextWord
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFindNextWord
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdFindNextWord
        End If
    End If
End Sub

Public Sub mnuNavigationSearchPrevWord()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindPrevWord
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFindPrevWord
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
'            fDoc.cs.ExecuteCmd cmCmdFindPrevWord
        End If
    End If
End Sub
    
Public Sub mnuNavigationReplace()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindReplace
        ElseIf ActiveFileForm.Identify = FF_IMP Then
            Set fImp = ActiveForm
            fImp.cs.ExecuteCmd cmCmdFindReplace
        ElseIf ActiveFileForm.Identify = FF_LIST Then
'            Set fDoc = ActiveForm
        End If
    End If
End Sub

Public Sub mnuNavigationGoToLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdGoToLine, -1
        End If
    End If
End Sub

Public Sub mnuNavigationGotoMatchBrace()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdGoToMatchBrace
        End If
    End If
End Sub

Public Sub mnuNavigationGoToIdent()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdGoToIndentation
        End If
    End If
End Sub

Public Sub mnuNavigationNextFunc()
    Dim sLine As String
    Dim i As Long, j As Long, g As Long
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            j = fDoc.rangoActual.StartLineNo + 1
            For i = j To fDoc.cs.LineCount
                sLine = LCase(fDoc.cs.getLine(i))
                If InStr(1, sLine, "process") Or InStr(1, sLine, "function") Then
                    g = i - j
'                    For j = 0 To g
'                        fDoc.cs.ExecuteCmd cmCmdLineDown
'                    Next j
                    j = j + g
                    fDoc.cs.ExecuteCmd cmCmdGoToLine, j
                    Exit Sub
                End If
            Next i
        End If
    End If
End Sub

Public Sub mnuNavigationPrevFunc()
    Dim sLine As String
    Dim i As Long, j As Long, g As Long
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            j = fDoc.rangoActual.StartLineNo - 1
            'MsgBox j
            For i = j To 1 Step -1
                sLine = LCase(fDoc.cs.getLine(i))
                'MsgBox i & " line: " & sLine
                If InStr(1, sLine, "process") Or InStr(1, sLine, "function") Then
                    'MsgBox "found at " & i
                    g = j - i
                    'MsgBox g
'                    For j = 0 To g
'                        fDoc.cs.ExecuteCmd cmCmdLineUp
'                    Next j
                    j = j - g
                    fDoc.cs.ExecuteCmd cmCmdGoToLine, j
                    Exit Sub
                End If
            Next i
        End If
    End If
End Sub

Public Sub mnuEditColumnMode()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.EnableColumnSel Then
                fDoc.cs.EnableColumnSel = False
                frmMain.StatusBar.PanelText("SEL_MODE") = "Mode Normal"
            Else
                fDoc.cs.EnableColumnSel = True
                frmMain.StatusBar.PanelText("SEL_MODE") = "Mode Column - Use CTRL + Mouse to select"
            End If
        End If
    End If
End Sub

Public Sub mnuEditCodeCompletionHelp()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdCodeList
        End If
    End If
End Sub

Public Sub mnuNavigationPrevPosition()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.codePosIndex > 1 Then
                fDoc.codePosIndex = fDoc.codePosIndex - 1
                fDoc.cs.ExecuteCmd cmCmdGoToLine, fDoc.getPos(fDoc.codePosIndex)
                frmDoc.refreshPosList
            End If
        End If
    End If
End Sub

Public Sub mnuNavigationNextPosition()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.codePosIndex < fDoc.uPos Then
                fDoc.codePosIndex = fDoc.codePosIndex + 1
                fDoc.cs.ExecuteCmd cmCmdGoToLine, fDoc.getPos(fDoc.codePosIndex)
                frmDoc.refreshPosList
            End If
        End If
    End If
End Sub

Public Sub mnuNavigationGoToDefinition()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            'MsgBox fDoc.cs.CurrentToken
            If isReservedWord(fDoc.cs.CurrentWord) Then
                MsgBox fDoc.cs.CurrentWord & " is language reserved word"
            ElseIf isUserDefinedType(fDoc.cs.CurrentWord) Then
                MsgBox fDoc.cs.CurrentWord & " is user defined type"
            ElseIf isOperator(fDoc.cs.CurrentWord) Then
                 MsgBox fDoc.cs.CurrentWord & " is an operator"
            ElseIf isDefinedType(fDoc.cs.CurrentWord) Then
                MsgBox fDoc.cs.CurrentWord & " is defined type"
'            ElseIf isReservedFunction(fDoc.cs.CurrentWord) Then
'                MsgBox fDoc.cs.CurrentWord & " is language reserved function"
            ElseIf frmProgramInspector.findNode(fDoc.cs.CurrentWord) Then
                
            Else
                MsgBox fDoc.cs.CurrentWord & " is no token"
            End If
        End If
    End If
End Sub

Public Sub mnuEditTab()
    Dim text    As String
    Dim Pos     As New CodeSense.position
    Dim line    As Integer
    Dim tabline As String
    Dim tabLen  As Integer
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdIndentSelection
        End If
    End If
End Sub

Public Sub mnuEditUnTab()
    Dim text        As String
    Dim textTemp    As String
    Dim Pos         As New CodeSense.position
    Dim line        As Integer
    Dim tabLen      As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdUnindentSelection
        End If
    End If
End Sub

Public Sub mnuEditComment()
    Dim text    As String
    Dim Pos     As New CodeSense.position
    Dim line    As Integer
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Pos.ColNo = fDoc.rangoActual.StartColNo
            Pos.LineNo = fDoc.rangoActual.StartLineNo
            
            If fDoc.cs.SelText = "" Then
                Pos.ColNo = 0
                fDoc.cs.InsertText "//", Pos
            ElseIf fDoc.rangoActual.StartColNo = 0 And _
                    fDoc.rangoActual.EndColNo >= fDoc.cs.GetLineLength(fDoc.rangoActual.EndLineNo) Then
                text = ""
                For line = fDoc.rangoActual.StartLineNo To fDoc.rangoActual.EndLineNo
                    If Not line = fDoc.rangoActual.EndLineNo Then
                        text = text & "//" & fDoc.cs.getLine(line) & Chr(vbKeyReturn)
                    Else
                        text = text & "//" & fDoc.cs.getLine(line)
                    End If
                Next line
                fDoc.cs.ReplaceSel text
            Else
                text = fDoc.cs.SelText
                text = "/*" & text & "*/"
                fDoc.cs.ReplaceSel text
            End If
            
        End If
    End If
End Sub

Public Sub mnuEditUnComment()
    Dim text        As String
    Dim Pos         As New CodeSense.position
    Dim line        As Integer
    Dim lineTest    As String
    Dim lineLen     As Integer
    Dim tabLen      As Integer
    Dim spacedLine  As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.rangoActual Is Nothing Then
                Exit Sub
            End If
            Pos.ColNo = fDoc.rangoActual.StartColNo
            Pos.LineNo = fDoc.rangoActual.StartLineNo
            
            If fDoc.cs.SelText = "" Then
                line = fDoc.rangoActual.StartLineNo
                lineTest = fDoc.cs.getLine(line)
                lineLen = fDoc.cs.GetLineLength(line)
                tabLen = fDoc.cs.TabSize
                spacedLine = replace(lineTest, Chr(vbKeyTab), Space(tabLen))
                
                If Left(LTrim(spacedLine), 2) = "//" Then '_____//comment types
                   
                    lineTest = replace(lineTest, "//", "", , 1)
                    fDoc.cs.SelectLine line, True
                    fDoc.cs.ReplaceSel lineTest
                End If
                
            ElseIf Right(fDoc.cs.SelText, 2) = "*/" And Left(fDoc.cs.SelText, 2) = "/*" Then
                text = fDoc.cs.SelText
                text = Left(Right(text, Len(text) - 2), Len(text) - 4)
                fDoc.cs.ReplaceSel text
            Else
                If Left(fDoc.cs.getLine(fDoc.rangoActual.StartLineNo), 2) = "//" Then
                    If fDoc.rangoActual.EndColNo >= fDoc.cs.GetLineLength(fDoc.rangoActual.EndLineNo) Then
                        text = ""
                        For line = fDoc.rangoActual.StartLineNo To fDoc.rangoActual.EndLineNo
                            
                            lineTest = fDoc.cs.getLine(line)
                            lineLen = fDoc.cs.GetLineLength(line)
                            tabLen = fDoc.cs.TabSize
                            spacedLine = replace(lineTest, Chr(vbKeyTab), Space(tabLen))
                            
                            If Left(lineTest, 2) = "//" Then    'if line starts with comments, delete comments
                                If Not line = fDoc.rangoActual.EndLineNo Then
                                    text = text & Right(lineTest, lineLen - 2) & Chr(vbKeyReturn)
                                Else
                                    lineTest = replace(lineTest, "//", "", , 1)
                                    text = text & lineTest
                                End If
                            ElseIf Left(LTrim(spacedLine), 2) = "//" Then '_____//comment types
                                
                                lineTest = replace(lineTest, "//", "", , 1)
                                
                                If Not line = fDoc.rangoActual.EndLineNo Then
                                    text = text & lineTest & " " & Chr(vbKeyReturn)
                                Else
                                    text = text & lineTest & " "
                                End If
                                
                            Else    'enter line as it is
                                If Not line = fDoc.rangoActual.EndLineNo Then
                                    text = text & lineTest & " " & vbCrLf 'Chr(vbKeyReturn)
                                Else
                                    text = text & lineTest & " "
                                End If
                            End If
                            
                        Next line
                        fDoc.cs.ReplaceSel text
                    End If
                End If
            End If
            
        End If
    End If
End Sub

Public Sub mnuEditUpperCase()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            text = fDoc.cs.SelText
            text = UCase(text)
            fDoc.cs.ReplaceSel text
        End If
    End If
End Sub

Public Sub mnuEditLowerCase()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            text = fDoc.cs.SelText
            text = LCase(text)
            fDoc.cs.ReplaceSel text
        End If
    End If
End Sub

Public Sub mnuEditFirstCase()       ' Proper Case
    Dim text        As String
    Dim textTemp    As String
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    Dim nextUCase   As Boolean
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            text = fDoc.cs.SelText
            textLen = Len(text)
            nextUCase = True
            For Pos = 1 To textLen
                Char = Right(Left(text, Pos), 1)  'get pos'th char in string
                If Char = " " Then                 'if char is space
                    nextUCase = True
                ElseIf nextUCase Then
                    Char = UCase(Char)
                    nextUCase = False
                Else
                    Char = LCase(Char)
                    nextUCase = False
                End If
                textTemp = textTemp & Char
            Next Pos

            fDoc.cs.ReplaceSel textTemp
        End If
    End If
End Sub

Public Sub mnuEditSentenceCase()       ' Sentence case.
    Dim text        As String
    Dim textTemp    As String
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    Dim nextUCase   As Boolean
    Dim afterPoint  As Boolean
    
    nextUCase = True
    afterPoint = True

    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            text = fDoc.cs.SelText
            textLen = Len(text)
            nextUCase = True
            For Pos = 1 To textLen
                Char = Right(Left(text, Pos), 1)  ' get pos'th char in string
                
                If Char = "." Or Char = "?" Or Char = "!" Then
                    afterPoint = True
                ElseIf Char <> " " And Not afterPoint Then
                    afterPoint = False
                End If
                    
                If Char = " " And afterPoint Then                  ' if char is space
                    nextUCase = True
                ElseIf nextUCase And afterPoint Then
                    Char = UCase(Char)
                    nextUCase = False
                    afterPoint = False
                Else
                    Char = LCase(Char)
                    nextUCase = False
                End If
                
                textTemp = textTemp & Char
            Next Pos

            fDoc.cs.ReplaceSel textTemp
        End If
    End If
End Sub

Public Sub mnuEditChangeCase()      ' iNVERT cASE
    Dim text        As String
    Dim textTemp    As String
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            text = fDoc.cs.SelText
            textLen = Len(text)
            For Pos = 1 To textLen
                Char = Right(Left(text, Pos), 1)  'get pos'th char in string
                charTemp = UCase(Char)
                If charTemp = Char Then                 'if char is UCase
                    Char = LCase(Char)
                Else
                    Char = charTemp
                End If
                textTemp = textTemp & Char
            Next Pos

            fDoc.cs.ReplaceSel textTemp
        End If
    End If
End Sub

Public Sub mnuThotIndent()
    Dim text        As String
    Dim textTemp    As String       ' string where we save all the changes we made at identation
    Dim textTemp100 As String       ' temporal string to save changes before export to textTemp
    Dim cont100     As Integer
'    Dim failures    As Integer, no_failures As Integer
    Dim lineText    As String
    Dim curLine     As Long
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    Dim curInd      As Long
    Dim nextInd     As Long
    Dim startLine   As Long
    Dim startTime
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            
            textTemp = ""
            curInd = 0
            nextInd = 0
            cont100 = 0
'            failures = 0
'            no_failures = 0
    
            
            If fDoc.cs.SelText = "" Then    'indent all the source code
                fDoc.cs.ExecuteCmd cmCmdSelectAll
                text = fDoc.cs.SelText
                startLine = 0
            Else                            'indent only the selection
                startLine = fDoc.rangoActual.StartLineNo
                text = fDoc.cs.SelText
            End If
            
            If fDoc.rangoActual.EndLineNo - startLine > 1000 Then
                MsgBox "The autoidentation process can take a long time to perform." & vbNewLine & _
                        "If the program do not refresh or it seems to be fallen don't close, wait a bit more." & vbNewLine & _
                        "Please be patient and wait a while", vbExclamation
            End If
            
            startTime = time
            
            For curLine = startLine To fDoc.rangoActual.EndLineNo
            
                curInd = nextInd
                lineText = fDoc.cs.getLine(curLine)
                If cont100 = 500 Then
                    cont100 = 0
                    textTemp = textTemp & textTemp100
                    textTemp100 = ""
                    frmMain.StatusBar.PanelText("MAIN") = "Autoidenting: Finished percent: " & CLng((curLine - fDoc.rangoActual.StartLineNo) * 100 / (fDoc.rangoActual.EndLineNo - fDoc.rangoActual.StartLineNo)) & "% - Please wait...  "
                End If

                If indChange(lineText) = -2 Then
                    curInd = 0
                    nextInd = 1
                ElseIf indChange(lineText) = -1 Then
                    curInd = curInd - 1
                    nextInd = curInd
                ElseIf indChange(lineText) = 1 Then
                    nextInd = curInd + 1
                ElseIf indChange(lineText) = 2 Then
                    nextInd = curInd
                    curInd = curInd - 1
                End If
                
                If getLineInd(lineText) = curInd Then
                    textTemp100 = textTemp100 & lineText & vbNewLine
                Else
                    textTemp100 = textTemp100 & putInInd(lineText, curInd) & vbNewLine
                End If
                'frmMain.StatusBar.PanelText("MAIN") = "Finished percent: " & CLng((curLine - fDoc.rangoActual.StartLineNo) * 100 / (fDoc.rangoActual.EndLineNo - fDoc.rangoActual.StartLineNo)) & "% - Please wait...  "
                cont100 = cont100 + 1
            Next
            
            textTemp = textTemp & textTemp100
            fDoc.cs.ReplaceSel textTemp
            'MsgBox "Time taken: " & CDate(startTime - time)
            'MsgBox "Failures: " & no_failures & " of " & fDoc.rangoActual.EndLineNo - fDoc.rangoActual.StartLineNo & " (" & CLng((no_failures / (fDoc.rangoActual.EndLineNo - fDoc.rangoActual.StartLineNo)) * 100) & "%)"
            
        End If
    End If
    textTemp = ""
End Sub

Private Function getLineInd(text As String) As Long
    Dim tabLen As Long
      
    tabLen = fDoc.cs.TabSize
    
    getLineInd = Int((Len(text) - Len(replace(text, Chr(vbKeyTab), ""))) / tabLen)
End Function


Private Function putInInd(text As String, ind As Long) As String
      Dim tabLen As Long
      
      tabLen = fDoc.cs.TabSize
      
      If ind < 0 Then ind = 0
      
      text = replace(text, Chr(vbKeyTab), "")
      text = Space(ind * tabLen) & text
      text = replace(text, Space(tabLen), Chr(vbKeyTab))
      
      putInInd = text
      
End Function

Private Function indChange(l As String) As Long

    Select Case Word(l)
        Case "program", "const", "global", "local", "private", "begin", "process", "function"
            indChange = -2
        Case "end", "until"
            indChange = -1
        Case "if", "for", "while", "loop", "switch", "case", "struct", "repeat", "from", "default"
            If LCase(Right(l, 3)) <> "end" Then
                indChange = 1
            Else
                indChange = 0
            End If
        Case "else", "elif", "elseif", "elsif"
            indChange = 2
        Case Else
            indChange = 0
    End Select

End Function

Private Function Word(line As String) As String
    Dim curWord As String
    Dim i As Long
    
    line = replace(line, Chr(vbKeyTab), "")
    i = 1
    While Mid(line, i, 1) = " "
        i = i + 1
    Wend
    
    If i > 1 And (Len(line) - i) > 0 Then
        line = Right(line, Len(line) - (i - 1))
    End If
    
    i = 1
    While isChar(Mid(line, i, 1))
        i = i + 1
    Wend

    curWord = Left(line, i - 1)

    Word = LCase(curWord)
End Function

Private Function isChar(Char As String) As Boolean
    Dim state As Boolean
    
    state = False
    Select Case Char
        Case "a" To "z", "A" To "Z"
            state = True
        Case " ", "("
            state = False
    End Select
    isChar = state
End Function

Public Sub mnuThotUnitifyBackslashes()
    Dim sel As String
    Dim curLine As Long
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            curLine = fDoc.rangoActual.StartLineNo - 1
            fDoc.cs.ExecuteCmd cmCmdSelectAll
            sel = fDoc.cs.SelText
            sel = replace(sel, "\", "/")
            fDoc.cs.ReplaceSel sel
            fDoc.rangoActual.StartLineNo = curLine
            fDoc.rangoActual.EndLineNo = curLine
            fDoc.cs.ExecuteCmd cmCmdGoToLine, curLine
        End If
    End If
End Sub

Public Sub mnuThotUnitifyFiles()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            MsgBox "Tests filenames to avoid Windows/Unix case types"
        End If
    End If
End Sub

Public Sub mnuThotDeclareFunctions()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            MsgBox "Declares automatically all functions in the source"
        End If
    End If
End Sub

Public Sub mnuThotFileTruster()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            MsgBox "Tests if all files used in the source exists"
        End If
    End If
End Sub

Public Sub mnuThotMigrate()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            MsgBox "Migrate old DIV/DIV2 code to Bennu/Fenix compatible"
        End If
    End If
End Sub

Public Sub mnuThotAddProcess()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            frmThotAdd.Show 1, frmMain
            frmThotAdd.optProcess = True
            'frmThotAdd.tabCategories.SelectedTab = 0
        End If
    End If
End Sub

Public Sub mnuThotAddFunction()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            frmThotAdd.Show 1, frmMain
            frmThotAdd.optFunction = True
            'frmThotAdd.tabCategories.SelectedTab = 0
        End If
    End If
End Sub

Public Sub mnuThotAddStruct()
    Dim text As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            frmThotAdd.Show 1, frmMain
            'frmThotAdd.tabCategories.SelectedTab = 1
        End If
    End If
End Sub


Public Sub mnuBookmarkToggle()
    Dim lineNum As Long
    Dim Index As Long
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkToggle
            lineNum = fDoc.rangoActual.StartLineNo + 1
            Index = fDoc.existsBookmark(lineNum)
            If Index <> -1 Then
                fDoc.delBookmark (Index)
            Else
                fDoc.addBookmark (lineNum)
            End If
            fDoc.refreshBookmarkList

        End If
    End If
End Sub

Public Sub mnuBookmarkNext()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkNext
            fDoc.cs.HighlightedLine = fDoc.rangoActual.StartLineNo
        End If
    End If
End Sub

Public Sub mnuBookmarkPrev()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkPrev
            fDoc.cs.HighlightedLine = fDoc.rangoActual.StartLineNo
        End If
    End If
End Sub

Public Sub mnuBookmarkDel()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkClearAll
            fDoc.delAllBookmark
            fDoc.refreshBookmarkList
        End If
    End If
End Sub

Public Sub mnuBookmarkToDo()
    Dim sel As String
    Dim FN As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            FN = ActiveFileForm.Filename
            frmTodoList.Show
            If fDoc.rangoActual.IsEmpty Then
                sel = fDoc.cs.getLine(fDoc.rangoActual.StartLineNo)
            Else
                sel = fDoc.cs.SelText
            End If
            frmTodoList.newItemToDo filterText(sel), FN
        End If
    End If
End Sub

Public Sub mnuBookmarkEdit()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.getLastBookmarkIndex > 1 Then
                frmBookmarkEditor.Show 1, frmMain
            End If
        End If
    End If
End Sub

Public Function filterText(text As String) As String
' filters not printable chars as vbtab, crlf...
    text = replace(text, Chr(vbKeyTab), "")
    text = replace(text, vbCrLf, "")
    filterText = text
End Function

Public Sub mnuConvertBinHex()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And isBin(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = BinToHex(sText)
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

Public Sub mnuConvertBinDec()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And isBin(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = CStr(BinToDec(sText))
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

Public Sub mnuConvertHexBin()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And isHex(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = HexToBin(sText)
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

Public Sub mnuConvertHexDec()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And isHex(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = CStr(HexToDec(sText))
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

Public Sub mnuConvertDecHex()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And IsNumeric(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = DecToHex(CDbl(sText))
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

Public Sub mnuConvertDecBin()
    Dim sText As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            If fDoc.cs.SelLength > 0 And IsNumeric(fDoc.cs.SelText) Then
                sText = fDoc.cs.SelText
                sText = DecToBin(CDbl(sText))
                fDoc.cs.ReplaceSel (sText)
            End If
        End If
    End If
End Sub

' === Converters ============================================
Public Function DecToBin(DecVal As Double) As String
    Dim A As Double
    Dim b As Integer
    
    A = DecVal
    For b = 1 To Int(Log(DecVal) / Log(2)) + 1
        DecToBin = CDbl(A Mod 2) & DecToBin
        A = CDbl(Int(A / 2))
    Next b
End Function

Public Function DecToHex(DecVal As Double) As String
Dim A As Double, b As Double, c As String, d As Double
    A = DecVal
    For b = 1 To Int(Log(DecVal) / Log(16)) + 1
        d = CDbl(A Mod 16)
        Select Case d
            Case 0 To 9
                c = d
            Case Else
                c = Chr(55 + d)
        End Select
        DecToHex = c & DecToHex
        A = CDbl(Int(A / 16))
    Next b
End Function

Public Function BinToDec(bIn As String) As Double
Dim TotalDec As Double, A As Double
    For A = 1 To Len(bIn)
        TotalDec = (TotalDec * 2) + Mid(bIn, A, 1)
    Next A
    BinToDec = TotalDec
End Function

Public Function BinToHex(bIn As String) As String
    BinToHex = DecToHex(BinToDec(bIn))
End Function

Public Function HexToDec(HexVal As String) As Double
Dim TotalDec As Double, A As Double, c As Double
    For A = 1 To Len(HexVal)
        Select Case (Mid(HexVal, A, 1))
            Case 0 To 9
                c = (Mid(HexVal, A, 1))
            Case Else
                c = (Asc(Mid(HexVal, A, 1)) - 55)
        End Select
        TotalDec = (TotalDec * 16) + c
    Next A
    HexToDec = TotalDec
End Function

Public Function HexToBin(HexVal As String) As String
    HexToBin = DecToBin(HexToDec(HexVal))
End Function

Public Function isHex(hex As String) As Boolean
    Dim i As Double
    Dim h As String
    For i = 1 To Len(hex)
        h = LCase(Mid(hex, i, 1))
        If h <> "a" And h <> "b" And h <> "c" And h <> "d" And h <> "e" And h <> "f" _
            And h <> "1" And h <> "2" And h <> "3" And h <> "4" And h <> "5" And h <> "6" _
            And h <> "7" And h <> "8" And h <> "9" _
        Then
            
            isHex = False
            Exit Function
        End If
    Next i
    isHex = True
End Function

Public Function isBin(bIn As String) As Boolean
    Dim i As Double
    For i = 1 To Len(bIn)
        If Mid(bIn, i, 1) <> "1" And Mid(bIn, i, 1) <> "0" Then
            isBin = False
            Exit Function
        End If
    Next i
    isBin = True
End Function
'=== end converters ================================================

Public Sub mnuEditPreferences()
    frmPreferences.Show vbModal, frmMain
End Sub
'-------------------------------------------------------------------------------
'END EDIT MENU
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'START FILE MENU
'-------------------------------------------------------------------------------
Public Sub mnuViewToolBarStandard()
    Dim Id As Long
    
    Id = frmMain.cRebar.BandIndexForData("MainBar")
    frmMain.cRebar.BandVisible(Id) = Not frmMain.cRebar.BandVisible(Id)
End Sub

Public Sub mnuViewToolBarEdit()
    Dim Id As Long
    
    Id = frmMain.cRebar.BandIndexForData("EditBar")
    frmMain.cRebar.BandVisible(Id) = Not frmMain.cRebar.BandVisible(Id)
End Sub

Public Sub mnuViewProjectBrowser()
    frmMain.TabDock.FormShow frmProjectBrowser.name
End Sub

Public Sub mnuViewProgramInspector()
    frmMain.TabDock.FormShow frmProgramInspector.name
End Sub

Public Sub mnuViewProperties()
    frmMain.TabDock.FormShow frmProperties.name
End Sub

Public Sub mnuViewCompilerOutput()
    frmMain.TabDock.FormShow frmOutput.name
End Sub

Public Sub mnuViewDebugOutput()
    frmMain.TabDock.FormShow frmDebug.name
End Sub

Public Sub mnuViewErrorOutput()
    frmMain.TabDock.FormShow frmErrors.name
End Sub

Public Sub mnuViewFullScreen()
    Dim Hwnd As Long
    Dim newStyle As Long
    Dim DockedForm As TDockForm
    Dim i As Integer
    
    Static oldStyle As Long
    Static inFullScreen As Boolean
    Static oldWindowState As Integer
    
    '#TODO:
    ' - Restore focus to the window who had focus

    Hwnd = frmMain.Hwnd
    LockWindowUpdate Hwnd
    
    If inFullScreen = False Then
        
        oldWindowState = frmMain.WindowState
    
        'This is a trick to achieve the captionbar to be repainted correctly
        frmMain.WindowState = 1
        
        oldStyle = GetWindowLong(Hwnd, GWL_STYLE)
        
        'Hide caption bar
        newStyle = oldStyle And Not (WS_CAPTION Or WS_BORDER Or WS_MINIMIZEBOX Or _
                        WS_MAXIMIZEBOX Or WS_SYSMENU)
        SetWindowLong Hwnd, GWL_STYLE, newStyle
    
        'Maximize window
        frmMain.WindowState = 2
        
        'Hide all panels
        frmMain.TabDock.Visible = False
        
        frmMain.SetFocus
    Else
        'Restore caption bar
        SetWindowLong Hwnd, GWL_STYLE, oldStyle
        
        'Restore window to its previous state
        frmMain.WindowState = oldWindowState
        
        'Show all panels
        frmMain.TabDock.Visible = True
        
        frmMain.SetFocus
    End If
    
    LockWindowUpdate False
    
    inFullScreen = Not inFullScreen
End Sub
'-------------------------------------------------------------------------------
'END FILE MENU
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'START EXECUTE MENU
'-------------------------------------------------------------------------------
Public Sub mnuExecuteCompileFile()
    Dim ff As IFileForm
    
    If Not ActiveFileForm Is Nothing Then
        Set ff = ActiveFileForm
        If ff.Identify = FF_SOURCE Then
            SaveBeforeCompiling
            CompileFFSource ff
        End If
    End If
End Sub

Public Sub mnuExecuteRunFile()
    Dim ff As IFileForm
    
    If Not ActiveFileForm Is Nothing Then
        Set ff = ActiveFileForm
        If ff.Identify = FF_SOURCE Then
            RunFFSource ff
        End If
    End If
End Sub

Public Sub mnuExecuteCompileAndRunFile()
    Dim ff As IFileForm
    
    If Not ActiveFileForm Is Nothing Then
        Set ff = ActiveFileForm
        If ff.Identify = FF_SOURCE Then
            SaveBeforeCompiling
            If CompileFFSource(ff) = True Then
                RunFFSource ff
            End If
        End If
    End If
End Sub

Public Sub mnuExecuteBuild()
    If Not openedProject Is Nothing Then
        If openedProject.mainSource <> "" Then
            If Compile(makePathForProject(openedProject.mainSource)) Then
                MsgBox "Compilation succesfull"
            End If
        Else
            MsgBox MSG_MNUACTIONS_MAINSOURCENOTDEFINED, vbExclamation
        End If
    End If
End Sub

Public Sub mnuExecuteRun()
    If Not openedProject Is Nothing Then
        If openedProject.mainSource <> "" Then
            Run makePathForProject(openedProject.mainSource)
        Else
            MsgBox MSG_MNUACTIONS_MAINSOURCENOTDEFINED, vbExclamation
        End If
    End If
End Sub

Public Sub mnuExecuteBuildAndRun()
    Dim sFile As String
    
    If Not openedProject Is Nothing Then
        If openedProject.mainSource <> "" Then
            sFile = makePathForProject(openedProject.mainSource)
            If Compile(sFile) Then
                Run sFile
            End If
        Else
            MsgBox MSG_MNUACTIONS_MAINSOURCENOTDEFINED, vbExclamation
        End If
    End If
End Sub
'-------------------------------------------------------------------------------
'END EXECUTE MENU
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'START PROJECT MENU
'-------------------------------------------------------------------------------
Public Sub mnuProjectProperties()
        If Not openedProject Is Nothing Then
            Dim fProject As New frmProjectProperties
            fProject.LoadConf
            fProject.Show vbModeless, frmMain
        End If
End Sub
Public Sub mnuProjectDevList()
    frmDevelopersList.Show 1, frmMain
End Sub
Public Sub mnuProjectTracker()
    If frmMain.cMenu.ItemChecked("mnuProjectTracker") = True Then
        Unload frmTodoList
    Else
        frmTodoList.Show
    End If
    frmMain.RefreshTabs
End Sub
Public Sub mnuProjectSetAsMainSource()
    Dim ff As IFileForm
    Set ff = ActiveFileForm
    If Not ff Is Nothing Then
        If ff.Identify = FF_SOURCE And Not ff.FilePath = "" Then
            If openedProject.FileExist(ff.FilePath) Then
                openedProject.mainSource = ff.FilePath
                frmProjectBrowser.RefreshTree
            End If
        End If
    End If
End Sub

Public Sub mnuProjectRemoveFile()
    Dim ff As IFileForm
    Set ff = ActiveFileForm
    If Not ff Is Nothing Then
        If openedProject.FileExist(ff.FilePath) Then
            openedProject.RemoveFile (ff.FilePath)
            frmProjectBrowser.RefreshTree
        End If
    End If
End Sub

Public Sub mnuProjectAddFile()
    Dim sFiles() As String
    Dim fileCount As Integer
    Dim i As Integer
    
    fileCount = ShowOpenDialog(sFiles, getFilter("COMMON_FILES"), True, True)
    If fileCount > 0 Then
        For i = LBound(sFiles) To UBound(sFiles)
            openedProject.AddFile sFiles(i)
        Next
        frmProjectBrowser.RefreshTree
    End If
End Sub

Public Sub mnuProjectCreateImp()
    Dim text        As String
    Dim textTemp    As String       ' string where we save all the changes we made at identation
    Dim lineText    As String
    Dim curLine     As Long
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    Dim startLine   As Long
    
    
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            
            frmMain.StatusBar.PanelText("MAIN") = "Creating import file - Please wait...  "
            
            textTemp = ""
   
            fDoc.cs.ExecuteCmd cmCmdSelectAll
            text = fDoc.cs.SelText
            startLine = 0
            
            If fDoc.rangoActual.EndLineNo - startLine > 1000 Then
                MsgBox "The autoImp process can take a long time to perform." & vbNewLine & _
                        "If the program do not refresh or it seems to be fallen don't close, wait a bit more." & vbNewLine & _
                        "Please be patient and wait a while", vbExclamation
            End If
            
            
            For curLine = startLine To fDoc.rangoActual.EndLineNo
            
                lineText = fDoc.cs.getLine(curLine)
                
                
            Next
            
            fDoc.cs.ExecuteCmd cmCmdClearSelection
        End If
    End If
    textTemp = ""
    text = ""
End Sub

Public Sub mnuProjectBuild()
    frmMain.StatusBar.PanelText("MAIN") = "Building project"
End Sub

'Private Function getModule(line As String) As String
'    Dim strFunc As String
'    Dim i As Long
'
'    strFunc = getFunction(line)
'    For i = LBound(modules) To UBound(modules)
'        If module(i) = strFunc Then
'            getModule = moduleName(i)
'            Exit Function
'        End If
'    Next
'    getModule = ""
'End Function


'Private Function getFunction(line As String) As String
'    Dim curWord As String
'    Dim i As Long
'
'    line = replace(line, Chr(vbKeyTab), "")
'    i = 1
'    While Mid(line, i, 1) = " "
'        i = i + 1
'    Wend
'
'    If i > 1 And (Len(line) - i) > 0 Then
'        line = Right(line, Len(line) - (i - 1))
'    End If
'
'    i = 1
'    While isChar(Mid(line, i, 1))
'        i = i + 1
'    Wend
'
'    curWord = Left(line, i - 1)
'
'    Word = LCase(curWord)
'End Function
'-------------------------------------------------------------------------------
'END PROJECT MENU
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'START TOOLS MENU
'-------------------------------------------------------------------------------
Public Sub mnuToolsCalculator()
    Shell "calc.exe"
End Sub

Public Sub mnuToolsIconChanger()
    MsgBox "Not working yet"
End Sub

Public Sub mnuToolsCommand()
    frmMSDOSCommand.Show 1
    'MsgBox "Show MS-DOS command form"
End Sub

Public Sub mnuToolsLastCommand()
    frmMSDOSCommand.callLastCommand
    'MsgBox "Exectures the last MS-DOS command"
End Sub

Public Sub mnuToolsConfigureTools()
    frmExtensions.Show 1
End Sub

Public Sub mnuToolsRunTool(Index As Integer)
    Dim sCommand As String
    Dim sParams As String
    Dim svar As String
    
    sCommand = ExternalTools(Index).Command
    sParams = ExternalTools(Index).Params
    
    'Replace variables
    svar = ""
    If Not ActiveFileForm Is Nothing Then svar = ActiveFileForm.FilePath
    sParams = replace(sParams, "$(FILE_PATH)", svar)
    
    WinExec ExternalTools(Index).Command & " " & sParams, SW_SHOWDEFAULT
End Sub
'-------------------------------------------------------------------------------
'END TOOLS MENU
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'START HELP MENU
'-------------------------------------------------------------------------------
Public Sub mnuHelpIndex()
    NewWindowWeb App.Path & "/help/fenix/func.php-frame=top.htm"
End Sub
Public Sub mnuHelpWiki()
    NewWindowWeb "http://fenixdocs.com/index.php/", "WIKI: "
End Sub
Public Sub mnuHelpWikiToken()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Dim sword As String
            sword = fDoc.cs.CurrentWord
            If sword <> "" Then
                ' wiki help
                NewWindowWeb "http://wiki.bennugd.org/index.php?title=" & UCase(sword), "WIKI: " & UCase(sword), "http://wiki.bennugd.org/index.php?title=Bennu_Wiki"
            Else
                mnuHelpWiki
            End If
        End If
    End If
End Sub
Public Sub mnuHelpBennuWiki()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Dim sword As String
            sword = fDoc.cs.CurrentWord
            If sword <> "" Then
                ' wiki help
                NewWindowWeb "http://wiki.bennugd.org/index.php?title=" & UCase(sword), "WIKI: " & UCase(sword), "http://wiki.bennugd.org/index.php?title=Bennu_Wiki"
            Else
                NewWindowWeb "http://wiki.bennugd.org/index.php?title=Bennu_Wiki", "Bennu Wiki"
            End If
        End If
    End If
End Sub
Public Sub mnuHelpFenixForum()
    NewWindowWeb "http://www.divsite.net/forum/index.php", "Fenix Forum"
End Sub
Public Sub mnuHelpBennuForum()
    NewWindowWeb "http://forum.bennugd.org/", "Bennu Forum"
End Sub
Public Sub mnuHelpProjectPage()
    NewWindowWeb "https://sourceforge.net/projects/fbtwo/", "Flamebird Homepage"
End Sub
Public Sub mnuHelpAbout()
    frmAbout.Show 1
End Sub
'-------------------------------------------------------------------------------
'END HELP MENU
'-------------------------------------------------------------------------------
'Public Sub RefreshStatusBar()
'    If Not ActiveFileForm Is Nothing Then
'        If ActiveFileForm.Identify = FF_SOURCE Then
'            Set fDoc = ActiveForm
'            If fDoc.rangoActual Is Nothing Then
'                Exit Sub
'            End If
'            If fDoc.rangoActual.StartLineNo = fDoc.rangoActual.EndLineNo Then
'                frmMain.StatusBar.PanelText("MAIN") = "Line: " & fDoc.rangoActual.StartLineNo + 1 _
'                    & " of " & fDoc.cs.LineCount & Chr(vbKeyTab) & "Sel: None"
'            Else
'                frmMain.StatusBar.PanelText("MAIN") = "Line: " & fDoc.rangoActual.StartLineNo + 1 _
'                    & " of " & fDoc.cs.LineCount & Chr(vbKeyTab) & "Sel: " _
'                    & fDoc.rangoActual.StartLineNo + 1 & " to " & fDoc.rangoActual.EndLineNo + 1 _
'                    & "   Len: " & fDoc.cs.SelLengthLogical
'            End If
'        End If
'    End If
'End Sub

Public Sub insertChar(Char As String)
    On Error Resume Next
    Dim Pos As New CodeSense.position
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Pos.ColNo = fDoc.rangoActual.StartColNo
            Pos.LineNo = fDoc.rangoActual.StartLineNo
            fDoc.cs.DeleteSel
            fDoc.cs.InsertText Char, Pos
        End If
        fDoc.rangoActual.StartColNo = fDoc.rangoActual.StartColNo + 1
    End If
End Sub
