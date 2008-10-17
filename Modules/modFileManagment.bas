Attribute VB_Name = "modFileManagment"
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

' **************************************************************
' Funciones de acciones desencadenadas por el menu file
' y aquellas relacionadas con el manejo de archivos
' **************************************************************
Option Explicit

'Recents
Public Enum ERecentConstants
    rtProject
    rtFile
End Enum
Public Type T_RECENT
  RecentFiles(1 To 5) As String
  RecentProjects(1 To 5) As String
End Type
Public Recents As T_RECENT

Private Const MSG_NewFileForm_FILENOTFOUND As String = "Can not open the file. The file does not exist!"
Private Const MSG_SAVEFILEWINDOW_ASKREPLACE As String = "The file already exists. Do you want to replace it?"
Private Const MSG_ADDFILETOPROJECT_NOTADDED As String = "The file was not added to the project."

Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Filters for the Save As and Open dialogs
Public Property Get FileFormOpenFilter(ff As EFileFormConstants) As String
    Dim sFilter As String
    Select Case ff
    Case FF_SOURCE
        sFilter = getFilter("SOURCE")
    Case FF_MAP
        sFilter = getFilter("MAP")
    Case FF_FPG
        sFilter = getFilter("FPG")
    Case FF_FNT
        sFilter = getFilter("FNT")
    Case Else
        MsgBox "Unknow FileFilter type (this should have never happened). Contact FBMX team"
    End Select
    FileFormOpenFilter = sFilter
End Property

Public Property Get FileFormSaveFilter(ff As EFileFormConstants) As String
    Dim sFilter As String
    Select Case ff
    Case FF_SOURCE
        sFilter = getFilter("SOURCE")
    Case FF_MAP
        sFilter = getFilter("MAP")
    Case FF_FPG
        sFilter = getFilter("FPG")
    Case FF_FNT
        sFilter = getFilter("FNT")
    Case Else
        MsgBox "Unknow FileFilter type (this should have never happened). Contact FBMX team"
    End Select
    FileFormSaveFilter = sFilter
End Property

Public Property Get FileFormDefaultExt(ff As EFileFormConstants) As String
    Dim sFilter As String
    Select Case ff
    Case FF_SOURCE
        sFilter = "prg"
    Case FF_MAP
        sFilter = "map"
    Case FF_FPG
        sFilter = "fpg"
    Case FF_FNT
        sFilter = "fnt"
    Case Else
        MsgBox "Unknow FileFilter type (this should have never happened). Contact FBMX team"
    End Select
    FileFormDefaultExt = sFilter
End Property

'Returns the Long Filename associated with sShortFilename
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    'First attempt using 1024 character buffer.

    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    
    'If buffer is too small lRet contains buffer size needed.
    If lRet > Len(sLongFilename) Then
        'Increase buffer size...
        sLongFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    
    'lRet contains the number of characters returned.
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function

' si el path es relativo, le agrega el dir del project
Public Function makePathForProject(ByVal sFile As String) As String
    If sFile = "" Then Exit Function
    ' si no contiene el nombre de la unidad es que el path es relativo
    If FSO.GetDriveName(sFile) = "" Then
        Dim directorioBase As String
        directorioBase = FSO.GetParentFolderName(openedProject.Filename)
        sFile = directorioBase & "\" & IIf(Left(sFile, 1) = "\", Right(sFile, Len(sFile) - 1), sFile)
    End If
    
    makePathForProject = sFile
End Function

' si el path es relativo, le agrega el dir del project, sin el filename
Public Function getPath(ByVal sFile As String) As String
    If sFile = "" Then Exit Function
    ' si no contiene el nombre de la unidad es que el path es relativo
    If FSO.GetDriveName(sFile) = "" Then
        Dim directorioBase As String
        directorioBase = FSO.GetParentFolderName(openedProject.Filename)
        sFile = directorioBase '& "\" & IIf(Left(sFile, 1) = "\", Right(sFile, Len(sFile) - 1), sFile)
    End If
    
    getPath = sFile
End Function

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START PROJECT MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Show NewProject dialog, create the project and show config
Public Sub NewProject()
    Dim sFile As String
    Dim fProject As frmProjectProperties
    
    sFile = ShowSaveDialog("fbp", getFilter("FBP"))
    
    If Len(sFile) > 0 Then
        Call OpenProject(sFile, True)  'Create the project file in disk
        Set fProject = New frmProjectProperties
        fProject.LoadConf
        fProject.SaveConf
        fProject.Show vbModal, frmMain
        frmProjectBrowser.RefreshTree 'Actualiza el árbol
    End If
End Sub

Public Sub OpenProject(ByVal sFile As String, Optional bNew As Boolean)
    Dim oldProject As cProject
    
    'Si hay un proyecto abierto, mantenemos un obj cProject de seguridad y lo cerramos
    If Not openedProject Is Nothing Then
        Set oldProject = openedProject
        CloseProject
    End If
    
    Set openedProject = New cProject
    openedProject.Filename = sFile
    
    If bNew Then 'Si no existe creamos un proyecto con configuración por defecto
        openedProject.projectName = "My project"
        'Categorías por defecto
        openedProject.AddCategory "Source files", "*.prg|*.inc|*.h"
        openedProject.AddCategory "Image files", "*.map|*.fbm|*.png"
        openedProject.AddCategory "Graphics Libraries", "*.fpg|*.fgc"
        openedProject.AddCategory "Sound files", "*.wav"
        openedProject.AddCategory "Music files", "*.mid|*.xm|*.it|*.mod|*.ogg"
        openedProject.AddCategory "Movie files", "*.mpeg|*.fli|*.mpg"
        openedProject.AddCategory "Palette files", "*.pal"
        openedProject.AddCategory "Font files", "*.fnt|*.ttf"

        'Trackers por defecto
        openedProject.colTrackers.Add "Bugs"
        openedProject.colTrackers(openedProject.colTrackers.KeyForName("Bugs")).AddCategory "Interface(Example)"
        
        openedProject.Save 'Crea los archivos del proyecto
    End If
    
    If openedProject.LoadProject Then 'carga satisfactoria
        setProjectMenu (True)
        openedProject.loadCache ' cargamos el cache
        openedProject.LoadIdeStatus 'Carga el estado del proyecto
        openedProject.LoadProjectConf 'Configuración local
        If openedProject.mainSource <> "" Then
             Dim nombre As String
             nombre = openedProject.mainSource
             If FSO.GetDriveName(nombre) = "" Then
                nombre = makePathForProject(nombre)
             End If
        End If
        AddRecent openedProject.Filename, rtProject
        frmProjectBrowser.RefreshTree
    Else
        'Reestablece el antiguo proyecto
        If Not oldProject Is Nothing Then
            Set openedProject = oldProject
            setProjectMenu (True)
            frmProjectBrowser.RefreshTree
            MsgBox "Error In OpenProject. ModFileManagement", vbCritical
        End If
    End If
End Sub

'Closes a project
Public Sub CloseProject()
    If Not openedProject Is Nothing Then
        openedProject.Save 'Guarda el projecto
        openedProject.dumpCache
        Set openedProject = Nothing
        setProjectMenu (False)
        frmProjectBrowser.tvProject.Nodes.Clear
        
        Unload frmTodoList 'cerrar el tracker
    End If
End Sub

Public Sub addFileToProject(ByVal sFile As String)
    Dim added As Boolean
    
    If Not openedProject Is Nothing Then
        added = openedProject.AddFile(sFile, False)
        If added = False Then
            MsgBox MSG_ADDFILETOPROJECT_NOTADDED, vbExclamation
        End If
        frmProjectBrowser.RefreshTree
    End If
End Sub

'Enable/disable project menu
Public Sub setProjectMenu(menuEnabled As Boolean)
    With frmMain.cMenu
        .ItemEnabled(.IndexForKey("mnuProjectSetAsMainSource")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectClose")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectProperties")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectAddFile")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectRemoveFrom")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectDevList")) = menuEnabled
        .ItemEnabled(.IndexForKey("mnuProjectTracker")) = menuEnabled
    End With
End Sub

'Determine if there is an Source Form opened
Public Function DocExist() As Boolean
    Dim Docs As Form
    For Each Docs In Forms
        If TypeOf Docs Is frmDoc Then
            DocExist = True
            Exit Function
        End If
    Next Docs
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END PROJECT MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START FILE FORM MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Returns the number of IFileForms opened
Public Function CountFileForms() As Integer
    Dim frm As Form
    Dim cnt As Integer
    
    For Each frm In Forms
        If TypeOf frm Is IFileForm Then
            cnt = cnt + 1
        End If
    Next frm
    
    CountFileForms = cnt
End Function

'Search a file form whoose FilePath property is the same than the sFile
'parameter and returns the form object
Public Function FindFileForm(ByVal sFile As String) As Form
    Dim frm As Form, fFileForm As IFileForm
    
    sFile = replace(sFile, "/", "\")
    
    Set FindFileForm = Nothing
    
    For Each frm In Forms
        If TypeOf frm Is IFileForm Then
            Set fFileForm = frm
            If StrComp(fFileForm.FilePath, sFile, vbTextCompare) = 0 Then
                Set FindFileForm = frm
                Exit For
            End If
        End If
    Next
    Set fFileForm = Nothing
End Function

'Creates a new File Form optionaly loading a file into it. The return value is a TEMPORAL solution
Public Function NewFileForm(ByVal ff As EFileFormConstants, Optional ByVal sFile As String) As Form

    On Error GoTo errhandler
    
    'This is a temporal solution to handle the Untitled documents count
    Static UntitledCount(0 To 10) As Integer
    
    Dim fFileForm As IFileForm
    Dim fForm As Form 'To acess standard form methods (i.e. show) NOT FOR ANY OTHER PURPOSE
    
    LockWindowUpdate frmMain.Hwnd
    'Determina el formulario que debe crear en función del tipo de archivo
    Select Case ff
    Case FF_SOURCE
        Set fFileForm = New frmDoc
    Case FF_MAP
        Set fFileForm = New frmMap
    Case FF_FPG
        Set fFileForm = New frmFpg
    Case FF_FNT
        Set fFileForm = New frmFnt
    End Select
    Set fForm = fFileForm
      
    'Load fForm
                 
    If sFile = "" Then 'New file
        If fFileForm.NewW(UntitledCount(ff)) Then  'New FileForm succesful
            fForm.Show
            Set NewFileForm = fFileForm 'TEMPORAL: No NewFileForm no debería tener un valor de retorno
            UntitledCount(ff) = UntitledCount(ff) + 1
            If ff = FF_SOURCE Then
                ' insert here necromancer code
                fForm.cs.AddText frmCodeWizard.sCode
                frmCodeWizard.sCode = ""
            End If
        Else
            Unload fFileForm
        End If
    Else 'Open an existing file
        If FSO.FileExists(sFile) Then
            If Not FindFileForm(sFile) Is Nothing Then 'The file is already opened
                Set fForm = FindFileForm(sFile)
                fForm.SetFocus
                Set NewFileForm = fForm 'TEMPORAL
            Else ' The file is not opened
                If fFileForm.Load(sFile) Then 'Loading succesful
                    AddRecent sFile 'Add file to recent list
                    fForm.Show
                    Set NewFileForm = fFileForm 'TEMPORAL
                Else 'Could not load
                    Unload fFileForm
                End If
            End If
        Else 'File does not exist-->Show error (This should never happen)
            MsgBox MSG_NewFileForm_FILENOTFOUND, vbCritical
        End If
    End If
    
    LockWindowUpdate False
    Set fFileForm = Nothing
    
    Exit Function
    
errhandler:
    If Err.Number > 0 Then MsgBox "NewFileForm":    Resume Next
End Function

'Shows open dialog to open a file associated with a FileForm
Public Sub OpenFileOfFileForm(ByVal ff As EFileFormConstants)
    Dim sFiles() As String
    Dim i As Integer
   
    If ShowOpenDialog(sFiles, FileFormOpenFilter(ff), True, True) > 0 Then
        For i = LBound(sFiles) To UBound(sFiles)
            NewFileForm ff, sFiles(i)
        Next
    End If
End Sub

'Opens a file according to its extension. PROVISIONAL
'(File extensions have to be reconsidered to be more general)
Public Sub OpenFileByExt(ByVal sFile As String)
    Dim sExt As String
    Dim ff As EFileFormConstants
    
    sExt = LCase(FSO.GetExtensionName(sFile))
    Select Case sExt
    Case "prg", "txt", "inc", "h"
        NewFileForm FF_SOURCE, sFile
    Case "map"
        NewFileForm FF_MAP, sFile
    Case "fbp"
        OpenProject sFile
    Case "fpg"
        NewFileForm FF_FPG, sFile
    Case "mod", "xm", "it", "s3m", "mid", "ogg", "wav"
        LoadPlayer sFile
    Case "fnt"
        NewFileForm FF_FNT, sFile
    Case Else
        MsgBox "File type not recognized", vbCritical
    End Select
End Sub

'Ask the fileform to save the window (and show the Save dialog if necessary)
Public Sub SaveFileOfFileForm(fFileForm As IFileForm, Optional ByVal bSaveAs As Boolean = False)
    Dim cdlg As cCommonDialog
    Dim sFile As String
    Dim lResult As Long
    Dim msgResult As VbMsgBoxResult
    
    On Error GoTo errhandler
    
    If Not fFileForm Is Nothing Then
        bSaveAs = (Not (fFileForm.AlreadySaved)) Or bSaveAs
        If bSaveAs Then 'Show the Save As dlg
            sFile = ShowSaveDialog(FileFormDefaultExt(fFileForm.Identify), FileFormSaveFilter(fFileForm.Identify))
            If sFile = "" Then Exit Sub
        Else
            sFile = fFileForm.FilePath
        End If
        'Check if the file arleady exists (and different from the existing one) ask
        'for replazing
'        If FSO.FileExists(sFile) And (fFileForm.AlreadySaved = False) Then
'            msgResult = MsgBox(MSG_SAVEFILEWINDOW_ASKREPLACE, vbQuestion + vbYesNo)
'            If msgResult = vbNo Then 'Abort
'                Exit Sub
'            End If
'        End If
        'Call the Save function of the IFileForm object
        lResult = fFileForm.Save(sFile)
        frmMain.RefreshTabs
        If Not lResult Then 'Error ocurred.
            'We don't need to do anything here cause is the Save function
            'of the IFileForm object who takes the control over success o error messages
            'This allows that some of the IFileForms show a message or not (for example
            'it is usefull to inform the user every time a map is saved but not for
            'a source file.
            'But let this be here for future purposes
        End If
    Else 'Null object
        'This should never happen... but can help for finding some bugs
        MsgBox "Empty FileForm in SaveFileWindow sub", vbCritical
    End If
    
    Exit Sub
errhandler:
    If Err.Number = &H4E21 Then  'Cancel error
        Set cdlg = Nothing
        Exit Sub
    ElseIf Err.Number > 0 Then 'Raise any different error
        Err.Raise Err.Number, Err.Source, Err.description
    End If
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END FILE FORM MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'Shows the Open song dialog
Public Sub OpenSong()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles, _
        "All known song modules (*.mod, *.s3m, *.xm, *.it, *.mid)|*.mod;*.s3m;*.xm;*.it;*.mid|" & _
        "All known audio stream files (*.ogg, *.wav)|*.ogg;*.wav|" & _
        "Mod module file (*.mod)|*.mod|" & _
        "S3m module file (*.s3m)|*.s3m|" & _
        "Xm module file (*.xm)|*.xm|" & _
        "It module file (*.it)|*.it|" & _
        "Midi file (*.mid)|*mid|" & _
        "Wave audio file (*.wav)|*.wav|" & _
        "Ogg Vorbis stream file (*.ogg)|*.ogg|" & _
        "All files (*.*)|*.*|" _
        , True, False) > 0 Then
        LoadPlayer sFiles(0)
'    ElseIf ShowOpenDialog(sFiles, getFilter("STREAM"), True, False) > 0 Then
'        LoadPlayer sFiles(0)
        AddRecent sFiles(0)
    End If
End Sub

Private Sub LoadPlayer(ByVal sFile As String)
    Load frmPlayer
    If frmPlayer.Load(sFile) = -1 Then
        frmPlayer.Show 0, frmMain
    Else
        Unload frmPlayer
    End If
End Sub



Public Sub NewWindowWeb(sURL As String, Optional Title As String, Optional Default As String)
    Dim NewBrowser As New frmWebBrowser
    
    If (Title = "") Then
        Title = "FENIX HELP"
    End If
    
    
    If FSO.GetDriveName(sURL) <> "" Then
        If FSO.FileExists(sURL) = False Then
            NewBrowser.Caption = "FENIX HELP"
            NewBrowser.URL = Default
            Exit Sub
        End If
    End If
    
    NewBrowser.Caption = Title
    
    NewBrowser.URL = sURL
    NewBrowser.Show
    frmMain.RefreshTabs
End Sub


'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
' START RECENT LIST MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
Public Sub AddRecent(str As String, Optional RecType As ERecentConstants = rtFile)
    Dim sRecents() As String
    Dim FreeFileNum As Integer, i As Integer
    Dim Pos As Long
    
    Erase sRecents()
    
    If RecType = rtFile Then
        ReDim sRecents(LBound(Recents.RecentFiles) To UBound(Recents.RecentFiles))
        sRecents() = Recents.RecentFiles()
    ElseIf RecType = rtProject Then
        ReDim sRecents(LBound(Recents.RecentProjects) To UBound(Recents.RecentProjects))
        sRecents() = Recents.RecentProjects()
    Else
        MsgBox "The rectype parameter is not a valid RECENTYPE", vbCritical, "Error": Exit Sub
    End If
    
    'Cancel if the recent is already in the list
    Pos = InStr(1, " " & Join(sRecents(), "@"), str, vbTextCompare)
    If Pos > 0 Then Exit Sub      'Ya está en la lista
    
    'Move up the elements of the recent list
    For i = UBound(sRecents) To (LBound(sRecents) + 1) Step -1
        sRecents(i) = sRecents(i - 1)
    Next
    sRecents(1) = str
    
    'Copy the temporary matrix to the recent matrix
    If RecType = rtFile Then
        For i = LBound(Recents.RecentFiles) To UBound(Recents.RecentFiles)
            Recents.RecentFiles(i) = sRecents(i)
        Next
    ElseIf RecType = rtProject Then
        For i = LBound(Recents.RecentProjects) To UBound(Recents.RecentProjects)
            Recents.RecentProjects(i) = sRecents(i)
        Next
    End If
    
    'Save the recent list to the recent.ini file
    FreeFileNum = FreeFile()
    Open App.Path & "\conf\recent.ini" For Binary Access Write As #FreeFileNum
        Put #FreeFileNum, , Recents
    Close #FreeFileNum
    
    'Create the recent menu
    frmMain.CreateMenuFromStrMatrix frmMain.cMenu, "mnuFileRecentFiles", "mnuRecFile", Recents.RecentFiles
    frmMain.CreateMenuFromStrMatrix frmMain.cMenu, "mnuFileRecentProjects", "mnuRecProj", Recents.RecentProjects

End Sub

Public Sub LoadRecents()
    Dim FreeFileNum As Integer, i As Integer
    Dim mnuKey As String

    Erase Recents.RecentFiles
    Erase Recents.RecentProjects
    
    'Load recent list from the recent.ini file
    FreeFileNum = FreeFile()
    Open App.Path & "\conf\recent.ini" For Binary Access Read As #FreeFileNum
        Get #FreeFileNum, , Recents
    Close #FreeFileNum
    
    TrimRecents
    
    frmMain.CreateMenuFromStrMatrix frmMain.cMenu, "mnuFileRecentFiles", "mnuRecFile", Recents.RecentFiles
    frmMain.CreateMenuFromStrMatrix frmMain.cMenu, "mnuFileRecentProjects", "mnuRecProj", Recents.RecentProjects
End Sub

'Limpia los recents
Private Sub TrimRecents()
    TrimRecentArray Recents.RecentFiles
    TrimRecentArray Recents.RecentProjects
End Sub

'Limpia los archivos inexistentes de un array de strings y los pone al final del mismo
Private Sub TrimRecentArray(RecentArray() As String)
    Dim i As Integer
    Dim sFiles As String, sArray() As String
    Dim bError As Boolean 'Controla si se produce un error en el acceso al dispositivo
    
    On Error GoTo errhandler
    
    For i = LBound(RecentArray) To UBound(RecentArray)
        bError = False
        RecentArray(i) = IIf(Dir(RecentArray(i)) = "", "", RecentArray(i)) 'Si no existe el archivo asigna ""
        sFiles = sFiles & IIf((RecentArray(i) = "") Or (bError = True), "", RecentArray(i) & "|")
    Next
    If Not sFiles = "" Then sFiles = Left(sFiles, Len(sFiles) - 1) 'Quita el último |
    
    sArray = Split(sFiles, "|")  'Separa en un array los ficheros
    Erase RecentArray 'Pone todos los elementos de los recents a ""
    'Copia el array a los recents
    For i = 0 To UBound(sArray)
        RecentArray(i + LBound(RecentArray)) = sArray(i)
    Next
    
errhandler:
    If Err.Number = 52 Then 'No se puede tener acceso al dispositivo
        bError = True
        Resume Next
    End If
End Sub
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END RECENT LIST MANAGEMENT
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

