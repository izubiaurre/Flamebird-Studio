VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProject"
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
Private Const CONFFILEEXT  As String = "#conf"
Private Const FBP_MAGIC As String = "fb2project"
Private Const FBP_VERSION As Integer = 2

Public devcol As cDeveloperCollection   'Collection of developers
Public colTrackers As cTrackerCollection

Private m_FileName As String            'Full path of the file of the project
Private m_projectName As String         'Project name

'Configuration variables of the project (non local, to share)
Private m_compilationDir As String      'Compilation directory path (output)
Private m_mainSource As String          'Main file of the project
Private m_colFiles As Collection        'File collection that contains the project
Private m_colCategories As Collection   'Category collection for the catview

'Local configuration variables
Private m_useOtherFenix As Boolean      'Determines if must be used another compiler
Private m_fenixDir As String            'Dir of the alternative compiler
Private m_compilerArguments As String   'Parameters of the compiler

Private iniManager As cInifile

Public Property Let projectName(ByVal vData As String)
    m_projectName = vData
End Property

Public Property Get projectName() As String
    projectName = m_projectName
End Property

Public Sub dumpCache()
    Dim fso1, txtfile, name
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    name = makePathForProject(fso1.GetBaseName(Me.Filename) & ".cache")
    
    If FSO.FileExists(name) Then
        FSO.DeleteFile name, True
    End If
    
    Set txtfile = FSO.CreateTextFile(name, True)
    
    Dim nodito As staticNode

    For Each nodito In includesNodes
        txtfile.WriteLine (nodito.father)
        txtfile.WriteLine (makePathRelative(nodito.Filename))
        txtfile.WriteLine (nodito.Icon)
        txtfile.WriteLine (nodito.Key)
        txtfile.WriteLine (nodito.lineNum)
        txtfile.WriteLine (nodito.name)
        txtfile.WriteLine (nodito.parameters)
        txtfile.WriteLine (nodito.varType)
        txtfile.WriteLine (nodito.varAmbient)
    Next nodito
    
    txtfile.Close
    
    Call SetFileAttributes(name, Hidden)

End Sub
Public Sub loadCache()
    
    Dim fso1 As New FileSystemObject, fil1 As file, ts As textStream
    
    If fso1.FileExists(makePathForProject(fso1.GetBaseName(Me.Filename) & ".cache")) = False Then
        Exit Sub
    End If
    
    Dim nodito As staticNode
    
    For Each nodito In includesNodes
        includesNodes.Remove nodito.Key
    Next nodito
    
    Set fil1 = fso1.GetFile(makePathForProject(fso1.GetBaseName(Me.Filename) & ".cache"))
    Set ts = fil1.OpenAsTextStream(ForReading)
    

    
    'Reads the content of the file
    While ts.AtEndOfStream = False
        Set nodito = New staticNode
        nodito.father = ts.ReadLine
        nodito.Filename = makePathForProject(ts.ReadLine)
        nodito.Icon = ts.ReadLine
        nodito.Key = ts.ReadLine
        nodito.lineNum = ts.ReadLine
        nodito.name = ts.ReadLine
        nodito.parameters = ts.ReadLine
        nodito.varType = ts.ReadLine
        nodito.varAmbient = ts.ReadLine
        includesNodes.Add nodito, nodito.Key
    Wend
    
    ts.Close
    
End Sub

'Public Function MoveFileUp(ByVal sFile As String)
'    Dim temp As String
'    Dim tempKey As String
'    Dim actualPos As Long
'
'    sFile = makePathRelative(sFile)
'    For actualPos = 1 To m_colFiles.count
'        If m_colFiles.item(actualPos) = sFile Then
'            Exit For
'        End If
'    Next
'
'    If actualPos > 1 And actualPos < m_colFiles.count Then
'        temp = m_colFiles.item(actualPos + 1)
'        tempKey = makePathForProject(temp)
'        m_colFiles.Remove (actualPos + 1)
'        m_colFiles.Add temp, tempKey, makePathForProject(sFile)
'    End If
'End Function

' Important: In th m_colFiles collection, Key has the FULL file path
Public Function AddFile(ByVal file As String, Optional AskForNonLocalFiles As Boolean = True) As Boolean
    Dim msgResult As VbMsgBoxResult
    Dim sRootFolder As String, sFile As String
    
    On Error GoTo ErrHandle
    
    sFile = makePathRelative(file)
    sRootFolder = FSO.GetParentFolderName(m_FileName) & "\"
    If Not FileExist(sFile) Then  'File isn't in the project
        'Look if the file isn¡t local (contains drive letter)
        If FSO.GetDriveName(sFile) <> "" And AskForNonLocalFiles Then
            msgResult = MsgBox("The file is out of the project root folder and it could be " _
                                & "inaccessible when moving the project to a different location." _
                                & vbCrLf & vbCrLf & "The project root folder is:" & vbCrLf _
                                & sRootFolder & vbCrLf & vbCrLf _
                                & "Would you like to copy the file there?", vbQuestion + vbYesNoCancel)
            If msgResult = vbYes Then
                FSO.CopyFile sFile, sRootFolder, False  'Copy the file
                sFile = makePathRelative(sRootFolder & FSO.GetFileName(sFile))
            ElseIf msgResult = vbCancel Then 'Cancel
                AddFile = False
                Exit Function
            End If
        End If
        'Add file to the collection
        m_colFiles.Add sFile, makePathForProject(sFile)
        AddFile = True
    Else
        MsgBox "File '" & file & "' is already in the project!", vbExclamation
        AddFile = False
    End If
    
    Exit Function
ErrHandle:
    If Err.Number = 58 Then 'File exists
        MsgBox "There is another file called '" & FSO.GetFileName(sFile) & "' in " _
                & "the project root folder. Please, delete this file before adding the new one" _
                , vbCritical
    Else
        MsgBox "Error trying to add a file" & vbCrLf & "Number: " & Err.Number & vbCrLf & "Description: " & Err.description, vbCritical
    End If
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      RemoveFile
'DESCRIPTION:   Delete a File from the project.
'PARAMS:        sFile: An string containing the path of the file (relative or absolute)
'RETURNS:       True if the file is found in the project files (and so, its removed)
'               Otherwise, false.
'-------------------------------------------------------------------------------------
Public Function RemoveFile(ByVal sFilename As String) As Boolean
    If FileExist(sFilename) = True Then
        m_colFiles.Remove makePathForProject(sFilename)
        RemoveFile = True
    Else
        MsgBox "This file cannot be deleted from the project." 'THIS SHOULD NEVER HAPPEN
        RemoveFile = False
    End If
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      FileExist
'DESCRIPTION:   Check if a file belongs to the project.
'PARAMS:        sFile: An string containing the path of the file (relative or absolute)
'RETURNS:       True if the file is found in the project files. Otherwise, false
'-------------------------------------------------------------------------------------
Public Function FileExist(ByVal sFile As String) As Boolean
    Dim s As Variant
    
    If sFile = "" Then Exit Function
    
    'If the path isn¡t relative (contains drive letter), makes relative
    If FSO.GetDriveName(sFile) <> "" Then sFile = makePathRelative(sFile)
    
    FileExist = False
    For Each s In m_colFiles
        If varType(s) = vbString Then
            If s = sFile Then
                FileExist = True
                Exit For 'File exists
            End If
        End If
    Next
End Function
'Public Function IndexForFile(sFile As String) As Integer
'    Dim s As Variant, i As Integer
'
'    i = 0
'    IndexForFile = 0
'    For Each s In m_colFiles
'        i = i + 1
'        If s = sFile Then
'            IndexForFile = i
'            Exit For
'        End If
'    Next
'End Function

Public Property Get Files() As Collection
    Set Files = m_colFiles
End Property

'Deletes all the files
Public Sub ClearFiles()
    Dim p As Variant
    For Each p In m_colFiles
        m_colFiles.Remove makePathForProject(CStr(p))
    Next
End Sub

'Save ide status (which elements are opened)
Private Sub SaveIdeStatus()
If confFileName <> "" Then
    With iniManager
        .Path = confFileName
        .Section = "IDESTATUS"
        'Files opened
        .Key = "quantity"
        .Value = CountFileForms()
        .Key = "filearray"
        .Value = FileArray()
        .Key = "activeform"
        .Value = ""
        If Not frmMain.ActiveFileForm Is Nothing Then
            .Value = frmMain.ActiveFileForm.Filename
        End If
        .Key = "Flametracker"
        .Value = IIf(frmMain.cMenu.ItemChecked(frmMain.cMenu.IndexForKey("mnuProjectTracker")), "1", "0")
    End With
End If
End Sub

'Restores the status that was the IDE (opened files, ...)
Public Sub LoadIdeStatus()
    Dim i As Integer
    Dim cant As Integer, lista As String, activo As String
    Dim FileArray() As String
    Dim frmActivo As Form
    
    If confFileName <> "" Then
        With iniManager
            .Path = confFileName
            .Section = "IDESTATUS"
            .Key = "quantity"
            cant = CInt(IIf(.Value = "", 0, .Value))
            .Key = "filearray"
            lista = .Value
            .Key = "activeform"
            activo = .Value
            .Key = "Flametracker"
            .Default = "0"
            If .Value = "1" Then
                mnuProjectTracker
            End If
            'frmMain.cMenu.ItemChecked(frmMain.cMenu.IndexForKey("mnuProjectTracker")) = IIf(.Value = "1", True, False)
        End With
    
        If cant > 0 Then 'Were opened files
            FileArray = Split(lista, "|", , vbTextCompare) 'We get an array of splitted elements
    
            For i = LBound(FileArray) To UBound(FileArray)
                OpenFileByExt (FileArray(i))
            Next i
    
            If activo <> "" Then 'Was at least one form opened
                If FSO.FileExists(activo) Then
                    Set frmActivo = FindFileForm(activo)
                    If Not frmActivo Is Nothing Then
                        frmActivo.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Property Let compilerArguments(ByVal vData As String)
    m_compilerArguments = vData
End Property
Public Property Get compilerArguments() As String
    compilerArguments = m_compilerArguments
End Property


Public Property Let useOtherFenix(ByVal vData As Boolean)
    m_useOtherFenix = vData
End Property
Public Property Get useOtherFenix() As Boolean
    useOtherFenix = m_useOtherFenix
End Property

Public Property Let mainSource(ByVal vData As String)
    m_mainSource = vData
End Property
Public Property Get mainSource() As String
    mainSource = m_mainSource
End Property

Public Function makePathRelative(ByVal sFilename As String) As String
    'Searches the dir where it's the project in the file that is going to add
    Dim directorioBase As String
    directorioBase = LCase(FSO.GetParentFolderName(m_FileName))
    Dim num As Long
    num = InStr(LCase(sFilename), directorioBase)
    
    'If found, deletes from path
    If num = 1 Then
        sFilename = Mid(sFilename, Len(directorioBase) + 2)
    End If
    
    makePathRelative = sFilename
End Function

Public Property Let fenixDir(ByVal vData As String)
    m_fenixDir = vData
End Property
Public Property Get fenixDir() As String
    fenixDir = m_fenixDir
End Property

Public Property Let compilationDir(ByVal vData As String)
    m_compilationDir = vData
End Property
Public Property Get compilationDir() As String
    compilationDir = m_compilationDir
End Property

Public Property Let Filename(ByVal vData As String)
    'Gives a new value
    m_FileName = vData
End Property
Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
    Filename = m_FileName
End Property

Public Property Get confFileName() As String
    confFileName = Left(m_FileName, Len(m_FileName) - 3) & CONFFILEEXT
End Property

'Adds a category
Public Sub AddCategory(ByVal sCatName, ByVal sExtensions)
    Dim cat As New cCatViewFolder
    cat.name = sCatName
    cat.Extensions = sExtensions
    m_colCategories.Add cat, cat.name
    Set cat = Nothing
End Sub
Public Sub ClearCategories()
    Dim p As Variant
    For Each p In m_colCategories
        m_colCategories.Remove p.name
    Next
End Sub

Public Property Get Categories() As Collection
    Set Categories = m_colCategories
End Property

'saves the project
Public Sub Save()
    SaveProjectConf
    SaveProject
    SaveIdeStatus
End Sub

'Saves local conf of the project
Private Sub SaveProjectConf()
    If Not confFileName = "" Then
        With iniManager
            .Path = confFileName
            'COMPILATION OPTIONS
            .Section = "COMPILATION"
            'Use alternative compiler
            .Key = "usespecificfenix"
            .Default = "NO"
            .Value = IIf(m_useOtherFenix, "YES", "NO")
            'Alternative compiler dir
            .Key = "fenixpath"
            .Default = ""
            .Value = m_fenixDir
            'Compilation parameters
            .Key = "parameters"
            .Default = ""
            .Value = m_compilerArguments
            'OTHERS
            .Section = "OTHER"
            'Default Developer
            .Key = "defaultdev"
            .Default = ""
            .Value = devcol.defaultDev
        End With
    End If
End Sub

'Loads the configuration of the project from local file
Public Sub LoadProjectConf()
    If Not confFileName = "" Then
        With iniManager
            .Path = confFileName
        
            'COMPILATION OPTIONS
            .Section = "COMPILATION"
            'Use alternative compiler
            .Key = "usespecificfenix"
            .Default = "NO"
            m_useOtherFenix = IIf(.Value = "YES", True, False)
            'Alternative compiler dir
            .Key = "fenixpath"
            .Default = ""
            m_fenixDir = .Value
            'Compilation parameters
            .Key = "parameters"
            .Default = ""
            m_compilerArguments = .Value
        End With
    End If
End Sub

'Saves non local configuration (fbp file)
Private Sub SaveProject()
    Dim fnum As Integer
    Dim dev As cDeveloper, tr As cTracker, ti As cTrackerItem
    Dim varStr As Variant, cat As cCatViewFolder
    
    If m_FileName = "" Then Exit Sub
    
    fnum = FreeFile()
    Open m_FileName For Binary Access Write As #fnum
    
    'FILE HEADER
    WriteStrInFile fnum, FBP_MAGIC      'Magic ("fb2project")
    Put fnum, , FBP_VERSION             'File version
    
    'GENERAL
    WriteStrInFile fnum, m_projectName                      'Project name
    WriteStrInFile fnum, m_compilationDir                   'Output dir
    WriteStrInFile fnum, makePathRelative(m_mainSource)     'Main file
    
    'CATVIEW CATEGORIES
    Put fnum, , m_colCategories.count           'Number of categorías
    For Each cat In m_colCategories
        WriteStrInFile fnum, cat.name                       'Category name
        WriteStrInFile fnum, cat.Extensions                 'Extensions
    Next
    
    'FILES
    Put fnum, , m_colFiles.count                'Number of files
    For Each varStr In m_colFiles
        If varType(varStr) = vbString Then
            WriteStrInFile fnum, CStr(varStr)   'File name
        End If
    Next
    
    'DEVELOPERS
    Put fnum, , devcol.count            'Number of developers
    'WriteStrInFile fnum, devcol.defaultDev     'Default dev
    For Each dev In devcol
        WriteStrInFile fnum, dev.name           'Dev name
        WriteStrInFile fnum, dev.RealName       'Real name
        WriteStrInFile fnum, dev.Mail           'Mail
    Next
    
    'TRACKER
    Put fnum, , colTrackers.count       'Number of trackers
    For Each tr In colTrackers
        'GENERAL
        WriteStrInFile fnum, tr.name            'Name
        WriteStrInFile fnum, tr.description     'Descripction
        Put fnum, , tr.IconIndex                'Icon
        
        'COLUMNS
        Put fnum, , tr.VisibleColumns           'Visible columns
        Put fnum, , tr.SortColumn               'Column sorted by
        Put fnum, , tr.SortOrder                'Asc or desc
        'cols arrangement
        
        'BEHAVIOR
        Put fnum, , CInt(tr.AutoExpandSelItems)     'Expand sel items
        Put fnum, , CInt(tr.ColorItemsByPriority)   'Paint priority
        Put fnum, , CInt(tr.ColorOldItems)          'paint old items
        Put fnum, , tr.OldItemsDays                 'Old Items days
        
        'CATEGORIES
        Put fnum, , tr.CategoryCount    'Number of categorías
        For Each varStr In tr.CategoryCol
            If varType(varStr) = vbString Then
                WriteStrInFile fnum, CStr(varStr)
            End If
        Next
        
        'ELEMENTS OF THE TRACKER
        Put fnum, , tr.count            'Number of items
        For Each ti In tr
            'GENERAL
            WriteStrInFile fnum, ti.Summary         'Summary
            WriteStrInFile fnum, ti.DetailedDesc    'Detailed desc
            WriteStrInFile fnum, ti.module          'Module
            WriteStrInFile fnum, ti.Category        'Category
            WriteStrInFile fnum, ti.AssignedTo      'Assigned to
            WriteStrInFile fnum, ti.SubmittedBy     'Send by
            Put fnum, , ti.Priority                 'Priority
            Put fnum, , ti.Completed                '% completed
            Put fnum, , CInt(ti.Closed)             'Closed?
            Put fnum, , CInt(ti.Hidden)             'Hidden?
            Put fnum, , CInt(ti.Locked)             'Locked
            Put fnum, , CDbl(ti.DateCreated)        'Creation date
            Put fnum, , CDbl(ti.DateClosing)        'Close date
            Put fnum, , CDbl(ti.DateModified)       'Last modification date
        Next
    Next
    
    Close fnum
End Sub

Public Function LoadProject() As Boolean
    Dim fnum As Integer
    Dim dev As New cDeveloper, tr As New cTracker, ti As New cTrackerItem, cat As New cCatViewFolder
    Dim strvar As String, intvar As Integer, lngvar As Long, lngvar2 As Long
    Dim bError As Boolean, i As Long, j As Long

    If m_FileName = "" Then Exit Function

    fnum = FreeFile()
    Open m_FileName For Binary Access Read As #fnum
    bError = False

    ' FILE HEADER
    strvar = ReadStrFromFile(fnum)     'Magic ("fb2project")
    If strvar = FBP_MAGIC Then
        Get fnum, , intvar             'File version
        If intvar <> FBP_VERSION Then
            MsgBox m_FileName & vbCrLf & vbCrLf & "Incompatible FBP version!", vbCritical
            bError = True
        End If
    Else
        MsgBox m_FileName & vbCrLf & vbCrLf & "Unrecognized file format!", vbCritical
        bError = True
    End If

    If bError = False Then
        'GENERAL
        m_projectName = ReadStrFromFile(fnum)                       'Project name
        m_compilationDir = ReadStrFromFile(fnum)                    'Output dir
        m_mainSource = makePathForProject(ReadStrFromFile(fnum))    'Main file
        If (FSO.FileExists(m_mainSource) = False) Then m_mainSource = ""

        'CATVIEW CATEGORIES
        ClearCategories
        Get fnum, , lngvar              'Number of categorías
        For i = 1 To lngvar
            strvar = ReadStrFromFile(fnum)          'Category name
            cat.name = strvar
            strvar = ReadStrFromFile(fnum)          'Extensions
            cat.Extensions = strvar
            AddCategory cat.name, cat.Extensions    'Add the category
        Next
        
        'FILES
        ClearFiles
        Get fnum, , lngvar              'Number of files
        For i = 1 To lngvar
            strvar = ReadStrFromFile(fnum)          'File name
            If Not Dir(makePathForProject(strvar)) = "" Then
                AddFile makePathForProject(strvar), False
            Else 'File doesn't exists
                MsgBox "Impossible to find " & strvar & ". The file will be removed from the " _
                        & "project.", vbCritical
            End If
        Next

        'DEVELOPERS
        devcol.Clear
        Get fnum, , lngvar              'Number of developers
        'strvar = ReadStrFromFile(fnum)             'Default dev
        For i = 1 To lngvar
            dev.name = ReadStrFromFile(fnum)        'Dev name
            dev.RealName = ReadStrFromFile(fnum)    'Real name
            dev.Mail = ReadStrFromFile(fnum)        'Mail
            devcol.Add dev.name, dev.RealName, dev.Mail
        Next
        'devcol.defaultDev = strvar
        'Set default developer
        With iniManager
            .Path = confFileName
            .Section = "OTHER"
            .Key = "defaultdev"
            devcol.defaultDev = .Value
        End With
    
        'TRACKER
        colTrackers.Clear
        Get fnum, , lngvar              'Number of trackers
        For i = 1 To lngvar
            'GENERAL
            tr.name = ReadStrFromFile(fnum)         'Name
            tr.description = ReadStrFromFile(fnum)  'Description
            tr.IconIndex = ReadLngFromFile(fnum)    'Icon

            'COLUMNAS
            tr.VisibleColumns = ReadLngFromFile(fnum)   'Visible columns
            tr.SortColumn = ReadLngFromFile(fnum)       'Column sorted by
            tr.SortOrder = ReadLngFromFile(fnum)        'Asc or desc
            'cols arrangement

            'COMPORTAMIENTO
            tr.AutoExpandSelItems = CBool(ReadIntFromFile(fnum)) 'Expandir sel items
            tr.ColorItemsByPriority = CBool(ReadIntFromFile(fnum)) 'Colorear prioridad
            tr.ColorOldItems = CBool(ReadIntFromFile(fnum)) 'Colorear old items
            tr.OldItemsDays = ReadIntFromFile(fnum) 'Old Items days

            'BEHAVIOR
            tr.CategoryClear
            Get fnum, , lngvar2         'Number of categorías
            For j = 1 To lngvar2
                strvar = ReadStrFromFile(fnum)
                tr.AddCategory strvar
            Next

            'ELEMENTS OF THE TRACKER
            tr.Clear
            Get fnum, , lngvar2     'Number of items
            For j = 1 To lngvar2
                'GENERAL
                ti.Summary = ReadStrFromFile(fnum)          'Summary
                ti.DetailedDesc = ReadStrFromFile(fnum)     'Detailed desc
                ti.module = ReadStrFromFile(fnum)           'Module
                ti.Category = ReadStrFromFile(fnum)         'Category
                ti.AssignedTo = ReadStrFromFile(fnum)       'Assigned to
                ti.SubmittedBy = ReadStrFromFile(fnum)      'Send by
                ti.Priority = ReadLngFromFile(fnum)         'Priority
                ti.Completed = ReadIntFromFile(fnum)        '% completed
                ti.Closed = CBool(ReadIntFromFile(fnum))    'Closed?
                ti.Hidden = CBool(ReadIntFromFile(fnum))    'Hidden?
                ti.Locked = CBool(ReadIntFromFile(fnum))    'Locked
                ti.DateCreated = CDate(ReadDblFromFile(fnum))   'Creation date
                ti.DateClosing = CDate(ReadDblFromFile(fnum))   'Close date
                ti.DateModified = CDate(ReadDblFromFile(fnum))  'Last modification date
                tr.AddIndirect ti
            Next

            colTrackers.AddIndirect tr
        Next

        LoadProject = True 'Sucessfull load
    Else
        LoadProject = False 'Not sucessfull load
    End If

    Close fnum
    Set ti = Nothing
    Set tr = Nothing
    Set dev = Nothing
    Set cat = Nothing
End Function

Private Sub Class_Initialize()
    Set iniManager = New cInifile
    Set devcol = New cDeveloperCollection
    Set colTrackers = New cTrackerCollection
    Set m_colFiles = New Collection
    Set m_colCategories = New Collection

'    'Categories by default
'    AddCategory "Source files", "*.prg|*.inc|*.h"
'    AddCategory "Image files", "*.map|*.fbm|*.png"
'    AddCategory "Graphics Libraries", "*.fpg|*.fgc"
    
End Sub

Private Sub Class_Terminate()
    Set m_colFiles = Nothing
    Set iniManager = Nothing
    Set devcol = Nothing
    Set colTrackers = Nothing
    Set m_colCategories = Nothing
End Sub

'These functions write and read a character string in a file
'Open opened by using a 4-byte varspace for length
'chain. The chain is only written if it is empty
Private Sub WriteStrInFile(ByVal fileNum As Integer, ByVal str As String)
    Dim lenStr As Long
    lenStr = Len(str)
    Put fileNum, , lenStr
    If lenStr > 0 Then Put fileNum, , str
End Sub
Private Function ReadStrFromFile(ByVal fileNum As Integer) As String
    Dim lenStr As Long, str As String
    str = ""
    Get fileNum, , lenStr
    If lenStr > 0 Then
        str = String(lenStr, " ")
        Get fileNum, , str
    End If
    ReadStrFromFile = CStr(str)
End Function
'The following function reads a file variable
'It has been created because VB can not read directly into a property of a class
'They could have made a general by VarType function, but subtract efficiency
'having to check the type of variable
Private Function ReadIntFromFile(ByVal fileNum As Integer) As Integer
    Dim var As Integer
    Get fileNum, , var
    ReadIntFromFile = var
End Function
Private Function ReadLngFromFile(ByVal fileNum As Integer) As Long
    Dim var As Long
    Get fileNum, , var
    ReadLngFromFile = var
End Function
Private Function ReadDblFromFile(ByVal fileNum As Integer) As Double
    Dim var As Double
    Get fileNum, , var
    ReadDblFromFile = var
End Function

