Attribute VB_Name = "modMenuActions"
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
Option Base 0

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function WinExec Lib "kernel32.dll" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Const SW_SHOWDEFAULT As Long = 10
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_BORDER As Long = &H800000
Private Const GWL_STYLE As Long = -16

Private Const MSG_MNUACTIONS_MAINSOURCENOTDEFINED As String = "Main source of the project has not been defined yet"

Private fDoc As frmDoc, fMap As frmMap

Private Enum NewTypeConstants
    NT_NONE
    NT_PROJECT
    NT_SOURCE
    NT_MAP
    NT_FPG
    NT_PAL
End Enum
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

Public Sub mnuFileOpenSong()
    OpenSong
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

Public Sub mnuFileClose()
    Dim f As Form
    Dim tr As RECT
    If Not ActiveFileForm Is Nothing Then
        Set f = ActiveForm
        GetWindowRect f.hwnd, tr
        tr.Top = 0
        LockWindowUpdate f.hwnd
        If Not f Is Nothing Then Unload f
        LockWindowUpdate False
        InvalidateRect frmMain.hwnd, tr, 0 'Refreshes the tab
    End If
End Sub

Public Sub mnuFileCloseAll()
    Dim lastHwnd As Long
    lastHwnd = -1
    Do Until (ActiveFileForm Is Nothing)
        If lastHwnd = ActiveForm.hwnd Then Exit Do 'If the form still is visible, the user selected cancel
        lastHwnd = ActiveForm.hwnd
        mnuFileClose
    Loop
End Sub

Public Sub mnuFileSaveAs()
    If Not ActiveFileForm Is Nothing Then
        SaveFileOfFileForm ActiveFileForm, True
    End If
End Sub

Public Sub mnuFileRecentOpen(sFile As String)
    If Not Dir(sFile) = "" Then
        OpenFileByExt sFile
    Else 'No existe el fichero
        MsgBox "File not found!", vbCritical
    End If
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
        End If
    End If
End Sub

Public Sub mnuEditDateTime()
    On Error Resume Next
    Dim timedate As String
    Dim Pos As New CodeSense.Position
    timedate = Date & "/" & Time
    
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

Public Sub mnuEditSelectAll()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdSelectAll
        End If
    End If
End Sub

Public Sub mnuEditSearch()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFind
        End If
    End If
End Sub

Public Sub mnuEditSearchNext()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindNext
        End If
    End If
End Sub
    
Public Sub mnuEditReplace()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdFindReplace
        End If
    End If
End Sub

Public Sub mnuEditGoToLine()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdGoToLine, -1
        End If
    End If
End Sub

Public Sub mnuEditGoToIdent()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdGoToIndentation
        End If
    End If
End Sub

Public Sub mnuAdvancedTab()
    Dim text    As String
    Dim Pos     As New CodeSense.Position
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

Public Sub mnuAdvancedUnTab()
    Dim text        As String
    Dim textTemp    As String
    Dim Pos         As New CodeSense.Position
    Dim line        As Integer
    Dim tabLen      As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdUnindentSelection
        End If
    End If
End Sub

Public Sub mnuAdvancedComment()
    Dim text    As String
    Dim Pos     As New CodeSense.Position
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
                        text = text & "//" & fDoc.cs.getLine(line) & Chr(13)
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

Public Sub mnuAdvancedUnComment()
    Dim text        As String
    Dim Pos         As New CodeSense.Position
    Dim line        As Integer
    Dim lineTest    As String
    Dim lineLen     As Integer
    Dim tabLen      As Integer
    Dim spacedLine  As String
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            Pos.ColNo = fDoc.rangoActual.StartColNo
            Pos.LineNo = fDoc.rangoActual.StartLineNo
            
            If fDoc.cs.SelText = "" Then
                If Left(fDoc.cs.getLine(fDoc.rangoActual.StartLineNo), 2) = "//" Then
                    line = fDoc.rangoActual.StartLineNo
                    text = fDoc.cs.getLine(line)
                    text = Right(text, Len(text) - 2)
                    fDoc.cs.SelectLine line, False
                    fDoc.cs.ReplaceSel text
                End If
            ElseIf Right(fDoc.cs.SelText, 2) = "*/" And Left(fDoc.cs.SelText, 2) = "/*" Then
                text = fDoc.cs.SelText
                text = Left(Right(text, Len(text) - 2), Len(text) - 4)
                fDoc.cs.ReplaceSel text
            Else
                If Left(fDoc.cs.getLine(fDoc.rangoActual.StartLineNo), 2) = "//" Then
                    text = ""
                    For line = fDoc.rangoActual.StartLineNo To fDoc.rangoActual.EndLineNo
                        lineTest = fDoc.cs.getLine(line)
                        lineLen = fDoc.cs.GetLineLength(line)
                        tabLen = fDoc.cs.TabSize
                        spacedLine = replace(lineTest, vbCrLf, Space(tabLen))
                        
                        If Left(lineTest, 2) = "//" Then    'if line starts with comments, delete comments
                            If Not line = fDoc.rangoActual.EndLineNo Then
                                text = text & Right(lineTest, lineLen - 2) & vbNewLine 'Chr(13)
                            Else
                                text = text & Right(lineTest, lineLen - 2)
                            End If
                        ElseIf Left(LTrim(spacedLine), 2) = "//" Then '_____//comment types
                            
                            lineTest = replace(lineTest, "//", "", , 1)
                            
                            If Not line = fDoc.rangoActual.EndLineNo Then
                                text = text & lineTest & " " & vbNewLine 'Chr(13)
                            Else
                                text = text & lineTest & " "
                            End If
                            
                        Else    'enter line as it is
                            If Not line = fDoc.rangoActual.EndLineNo Then
                                text = text & lineTest & " " & vbNewLine ' Chr(13)
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
End Sub

Public Sub mnuAdvancedUpperCase()
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

Public Sub mnuAdvancedLowerCase()
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

Public Sub mnuAdvancedFirstCase()
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

Public Sub mnuAdvancedChangeCase()
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

Public Sub mnuAdvancedIndent()
    Dim text        As String
    Dim textTemp    As String
    Dim lineText    As String
    Dim curLine     As Integer
    Dim Pos         As Integer
    Dim textLen     As Integer
    Dim Char        As String
    Dim charTemp    As String
    Dim curInd      As Long
    Dim nextInd     As Long
    Dim startLine   As Long
    
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            
            textTemp = ""
            
            If fDoc.cs.SelText = "" Then    'indent all the source code
                
                fDoc.cs.ExecuteCmd cmCmdSelectAll
                text = fDoc.cs.SelText
                
                startLine = 0

            Else                            'indent only the selection
                startLine = fDoc.rangoActual.StartLineNo
                text = fDoc.cs.SelText
            End If
            
            curInd = 0
            nextInd = 0
            For curLine = startLine To fDoc.rangoActual.EndLineNo
            
                curInd = nextInd
                lineText = fDoc.cs.getLine(curLine)
                                    
                Select Case indChange(lineText)
                    Case -2
                        curInd = 0
                        nextInd = 1
                    Case -1
                        curInd = curInd - 1
                        nextInd = curInd
                    Case 1
                        nextInd = curInd + 1
                    Case 2
                        nextInd = curInd
                        curInd = curInd - 1
                End Select
                
                If getLineInd(lineText) = curInd Then
                    textTemp = textTemp & lineText & vbNewLine
                Else
                    textTemp = textTemp & putInInd(lineText, curInd) & vbNewLine
                End If
                
            Next
            
            fDoc.cs.ReplaceSel textTemp
        End If
    End If
End Sub

Private Function getLineInd(text As String) As Long
    Dim tabLen As Long
      
    tabLen = fDoc.cs.TabSize
    
    getLineInd = Int((Len(text) - Len(replace(text, vbCrLf, ""))) / tabLen)
End Function


Private Function putInInd(text As String, ind As Long) As String
      Dim tabLen As Long
      
      tabLen = fDoc.cs.TabSize
      
      If ind < 0 Then ind = 0
      
      text = replace(text, vbCrLf, "")
      text = Space(ind * tabLen) & text
      text = replace(text, Space(tabLen), vbCrLf)
      
      putInInd = text
      
End Function

Private Function indChange(l As String) As Long

    Select Case Word(l)
        Case "program", "const", "global", "local", "private", "begin", "process", "function"
            indChange = -2
        Case "end", "until"
            indChange = -1
        Case "if", "for", "while", "loop", "switch", "case", "struct", "repeat"
            If LCase(Right(l, 3)) <> "end" Then
                indChange = 1
            Else
                indChange = 0
            End If
        Case "else", "default"
            indChange = 2
        Case Else
            indChange = 0
    End Select

End Function

Private Function Word(line As String) As String
    Dim curWord As String
    Dim i As Long
    
    line = replace(line, vbCrLf, "")
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


Public Sub mnuBookmarkToggle()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkToggle
        End If
    End If
End Sub

Public Sub mnuBookmarkNext()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkNext
        End If
    End If
End Sub

Public Sub mnuBookmarkPrev()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkPrev
        End If
    End If
End Sub

Public Sub mnuBookmarkDel()
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            fDoc.cs.ExecuteCmd cmCmdBookmarkClearAll
        End If
    End If
End Sub

Public Sub mnuBookmarkToDo()
    Dim sel As String
    If Not ActiveFileForm Is Nothing Then
        If ActiveFileForm.Identify = FF_SOURCE Then
            Set fDoc = ActiveForm
            'fDoc.cs.ExecuteCmd cmCmdBookmarkClearAll
            
        End If
    End If
End Sub

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
    Dim id As Long
    
    id = frmMain.cRebar.BandIndexForData("MainBar")
    frmMain.cRebar.BandVisible(id) = Not frmMain.cRebar.BandVisible(id)
End Sub

Public Sub mnuViewToolBarEdit()
    Dim id As Long
    
    id = frmMain.cRebar.BandIndexForData("EditBar")
    frmMain.cRebar.BandVisible(id) = Not frmMain.cRebar.BandVisible(id)
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
    Dim hwnd As Long
    Dim newStyle As Long
    Dim dockedForm As TDockForm
    Dim i As Integer
    
    Static oldStyle As Long
    Static inFullScreen As Boolean
    Static oldWindowState As Integer
    
    'TODO:
    ' - Restore focus to the window who had focus

    hwnd = frmMain.hwnd
    LockWindowUpdate hwnd
    
    If inFullScreen = False Then
        
        oldWindowState = frmMain.WindowState
    
        'This is a trick to achieve the captionbar to be repainted correctly
        frmMain.WindowState = 1
        
        oldStyle = GetWindowLong(hwnd, GWL_STYLE)
        
        'Hide caption bar
        newStyle = oldStyle And Not (WS_CAPTION Or WS_BORDER Or WS_MINIMIZEBOX Or _
                        WS_MAXIMIZEBOX Or WS_SYSMENU)
        SetWindowLong hwnd, GWL_STYLE, newStyle
    
        'Maximize window
        frmMain.WindowState = 2
        
        'Hide all panels
        frmMain.TabDock.Visible = False
        
        frmMain.SetFocus
    Else
        'Restore caption bar
        SetWindowLong hwnd, GWL_STYLE, oldStyle
        
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
'-------------------------------------------------------------------------------
'END PROJECT MENU
'-------------------------------------------------------------------------------

Public Sub mnuToolsCalculator()
    Shell "calc.exe"
End Sub


Public Sub mnuToolsConfigureTools()
    frmExtensions.Show 1
End Sub

Public Sub mnuToolsRunTool(index As Integer)
    Dim sCommand As String
    Dim sParams As String
    Dim svar As String
    
    sCommand = ExternalTools(index).Command
    sParams = ExternalTools(index).Params
    
    'Replace variables
    svar = ""
    If Not ActiveFileForm Is Nothing Then svar = ActiveFileForm.FilePath
    sParams = replace(sParams, "$(FILE_PATH)", svar)
    
    WinExec ExternalTools(index).Command & " " & sParams, SW_SHOWDEFAULT
End Sub

'-------------------------------------------------------------------------------
'START HELP MENU
'-------------------------------------------------------------------------------
Public Sub mnuHelpIndex()
    NewWindowWeb App.Path & "\help\fenix\func.php-frame=top.htm"
End Sub
Public Sub mnuHelpAbout()
    frmAbout.Show 1
End Sub
'-------------------------------------------------------------------------------
'END HELP MENU
'-------------------------------------------------------------------------------

