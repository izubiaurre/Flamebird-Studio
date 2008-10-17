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

Public devcol As cDeveloperCollection 'Colección de desarrolladores
Public colTrackers As cTrackerCollection

Private m_FileName As String 'Ruta completa del archivo de proyecto
Private m_projectName As String 'Nombre del proyecto

'Variables de configuración del proyecto no locales (para compatir)
Private m_compilationDir As String 'Ruta del directorio de compilación (output)
Private m_mainSource As String 'Archivo principal del proyecto
Private m_colFiles As Collection 'Colección de los ficheros que contiene el proyecto
Private m_colCategories As Collection 'Colección de las categorías para el catview

'Variables de configuración local
Private m_useOtherFenix As Boolean 'Determina si se debe usar un fxc distinto al configurado
Private m_fenixDir As String 'Directorio de fenix alternativo
Private m_compilerArguments As String 'Argumentos de compilación

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
    

    
    ' Lee el contenido del archivo.
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
'Importante: En la colección m_colFiles, Key guarda la ruta del archivo COMPLETA
Public Function AddFile(ByVal file As String, Optional AskForNonLocalFiles As Boolean = True) As Boolean
    Dim msgResult As VbMsgBoxResult
    Dim sRootFolder As String, sFile As String
    
    On Error GoTo ErrHandle
    
    sFile = makePathRelative(file)
    sRootFolder = FSO.GetParentFolderName(m_FileName) & "\"
    If Not FileExist(sFile) Then  'No pertenece el archivo al proyecto
        'Miramos si el archivo no es local (contiene la letra de la unidad)
        If FSO.GetDriveName(sFile) <> "" And AskForNonLocalFiles Then
            msgResult = MsgBox("The file is out of the project root folder and it could be " _
                                & "inaccessible when moving the project to a different location." _
                                & vbCrLf & vbCrLf & "The project root folder is:" & vbCrLf _
                                & sRootFolder & vbCrLf & vbCrLf _
                                & "Would you like to copy the file there?", vbQuestion + vbYesNoCancel)
            If msgResult = vbYes Then
                FSO.CopyFile sFile, sRootFolder, False  'Copia el archivo
                sFile = makePathRelative(sRootFolder & FSO.GetFileName(sFile))
            ElseIf msgResult = vbCancel Then 'Cancela
                AddFile = False
                Exit Function
            End If
        End If
        'Agregamos el fichero a la colección
        m_colFiles.Add sFile, makePathForProject(sFile)
        AddFile = True
    Else
        MsgBox "File '" & file & "' is already in the project!", vbExclamation
        AddFile = False
    End If
    
    Exit Function
ErrHandle:
    If Err.Number = 58 Then 'El archivo ya existe
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
    
    'Si el path no es relativo (contiene la unidad), lo hace relativo
    If FSO.GetDriveName(sFile) <> "" Then sFile = makePathRelative(sFile)
    
    FileExist = False
    For Each s In m_colFiles
        If varType(s) = vbString Then
            If s = sFile Then
                FileExist = True
                Exit For 'Existe el archivo
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

'Borra todos los ficheros
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
    End With
End If
End Sub

'Recupera el estado en el que se encontraba el IDE (archivos abiertos...)
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
        End With
    
        If cant > 0 Then 'Había formularios abiertos
            FileArray = Split(lista, "|", , vbTextCompare) 'Obtenemos un array con los elementos separados
    
            For i = LBound(FileArray) To UBound(FileArray)
                OpenFileByExt (FileArray(i))
            Next i
    
            If activo <> "" Then 'Había algún formulario activo
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
    ' busca el direcorio donde se encuentra el proyecto en
    ' el path del archivo que se quiere agregar
    Dim directorioBase As String
    directorioBase = LCase(FSO.GetParentFolderName(m_FileName))
    Dim num As Long
    num = InStr(LCase(sFilename), directorioBase)
    
    ' si es encontrado se elimina del path
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
    ' asigna valor nuevo
    m_FileName = vData
End Property
Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
    Filename = m_FileName
End Property

Public Property Get confFileName() As String
    confFileName = Left(m_FileName, Len(m_FileName) - 3) & CONFFILEEXT
End Property

'Añade una categoría
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

'Guarda el proyecto
Public Sub Save()
    SaveProjectConf
    SaveProject
    SaveIdeStatus
End Sub

'Guarda la configuración local del proyecto
Private Sub SaveProjectConf()
    If Not confFileName = "" Then
        With iniManager
            .Path = confFileName
            ' OPCIONES DE COMPILACION
            .Section = "COMPILATION"
            'Usar fenix alternativo
            .Key = "usespecificfenix"
            .Default = "NO"
            .Value = IIf(m_useOtherFenix, "YES", "NO")
            'Ruta del fenix alternativo
            .Key = "fenixpath"
            .Default = ""
            .Value = m_fenixDir
            'Argumentos de compilacion
            .Key = "parameters"
            .Default = ""
            .Value = m_compilerArguments
            'OTRAS
            .Section = "OTHER"
            'Default Developer
            .Key = "defaultdev"
            .Default = ""
            .Value = devcol.defaultDev
        End With
    End If
End Sub

'Carga la configuración del proyecto desde el archivo local
Public Sub LoadProjectConf()
    If Not confFileName = "" Then
        With iniManager
            .Path = confFileName
        
            ' OPCIONES DE COMPILACION
            .Section = "COMPILATION"
            'Usar fenix alternativo
            .Key = "usespecificfenix"
            .Default = "NO"
            m_useOtherFenix = IIf(.Value = "YES", True, False)
            'Ruta del fenix alternativo
            .Key = "fenixpath"
            .Default = ""
            m_fenixDir = .Value
            'Argumentos de compilacion
            .Key = "parameters"
            .Default = ""
            m_compilerArguments = .Value
        End With
    End If
End Sub

'Guarda la configuración no local (archivo fbp)
Private Sub SaveProject()
    Dim fnum As Integer
    Dim dev As cDeveloper, tr As cTracker, ti As cTrackerItem
    Dim varStr As Variant, cat As cCatViewFolder
    
    If m_FileName = "" Then Exit Sub
    
    fnum = FreeFile()
    Open m_FileName For Binary Access Write As #fnum
    
    'CABECERA DEL ARCHIVO
    WriteStrInFile fnum, FBP_MAGIC 'Magic ("fb2project")
    Put fnum, , FBP_VERSION 'Versión de archivo
    
    'GENERAL
    WriteStrInFile fnum, m_projectName 'Nombre del proyecto
    WriteStrInFile fnum, m_compilationDir 'Directorio output
    WriteStrInFile fnum, makePathRelative(m_mainSource) 'Archivo principal
    
    'CATEGORÍAS DEL CATVIEW
    Put fnum, , m_colCategories.count
    For Each cat In m_colCategories
        WriteStrInFile fnum, cat.name
        WriteStrInFile fnum, cat.Extensions
    Next
    
    'FICHEROS
    Put fnum, , m_colFiles.count 'Número de ficheros
    For Each varStr In m_colFiles
        If varType(varStr) = vbString Then
            WriteStrInFile fnum, CStr(varStr) 'Nombre del fichero
        End If
    Next
    
    'DEVELOPERS
    Put fnum, , devcol.count 'Número de developers
    'WriteStrInFile fnum, devcol.defaultDev 'Defaultr dev
    For Each dev In devcol
        WriteStrInFile fnum, dev.name 'Dev name
        WriteStrInFile fnum, dev.RealName 'Real name
        WriteStrInFile fnum, dev.Mail 'Mail
    Next
    
    'TRACKER
    Put fnum, , colTrackers.count 'Número de trackers
    For Each tr In colTrackers
        'GENERAL
        WriteStrInFile fnum, tr.name 'Nombre
        WriteStrInFile fnum, tr.description 'Descripción
        Put fnum, , tr.IconIndex 'Icono
        
        'COLUMNAS
        Put fnum, , tr.VisibleColumns 'Columnas visibles
        Put fnum, , tr.SortColumn 'Columna por la cual se ordena
        Put fnum, , tr.SortOrder 'Ascendente o descendente
        'cols arrangement
        
        'COMPORTAMIENTO
        Put fnum, , CInt(tr.AutoExpandSelItems) 'Expandir sel items
        Put fnum, , CInt(tr.ColorItemsByPriority) 'Colorear prioridad
        Put fnum, , CInt(tr.ColorOldItems) 'Colorear old items
        Put fnum, , tr.OldItemsDays 'Old Items days
        
        'CATEGORÍAS
        Put fnum, , tr.CategoryCount 'Número de categorías
        For Each varStr In tr.CategoryCol
            If varType(varStr) = vbString Then
                WriteStrInFile fnum, CStr(varStr)
            End If
        Next
        
        'ELEMENTOS DEL TRACKER
        Put fnum, , tr.count 'Número de items
        For Each ti In tr
            'GENERAL
            WriteStrInFile fnum, ti.Summary 'Summary
            WriteStrInFile fnum, ti.DetailedDesc 'Descripción detallada
            WriteStrInFile fnum, ti.module 'Módulo
            WriteStrInFile fnum, ti.Category 'Categoría
            WriteStrInFile fnum, ti.AssignedTo 'Assignado a
            WriteStrInFile fnum, ti.SubmittedBy 'Enviado por
            Put fnum, , ti.Priority 'Prioridad
            Put fnum, , ti.Completed '% completado
            Put fnum, , CInt(ti.Closed) 'Cerrado?
            Put fnum, , CInt(ti.Hidden)  'Oculto?
            Put fnum, , CInt(ti.Locked) 'Bloqueado
            Put fnum, , CDbl(ti.DateCreated) 'Fecha de creación
            Put fnum, , CDbl(ti.DateClosing) 'Fecha de cierre
            Put fnum, , CDbl(ti.DateModified) 'Fecha de última modificación
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

    'CABECERA DEL ARCHIVO
    strvar = ReadStrFromFile(fnum)  'Magic
    If strvar = FBP_MAGIC Then
        Get fnum, , intvar 'Versión
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
        m_projectName = ReadStrFromFile(fnum) 'Nombre del proyecto
        m_compilationDir = ReadStrFromFile(fnum)  'Directorio output
        m_mainSource = makePathForProject(ReadStrFromFile(fnum)) 'Archivo principal
        If (FSO.FileExists(m_mainSource) = False) Then m_mainSource = ""

        'CATEGORÍAS DEL CATVIEW
        ClearCategories
        Get fnum, , lngvar 'Número de categorías
        For i = 1 To lngvar
            strvar = ReadStrFromFile(fnum) 'Nombre de la categoría
            cat.name = strvar
            strvar = ReadStrFromFile(fnum) 'Extensiones
            cat.Extensions = strvar
            AddCategory cat.name, cat.Extensions 'Añadimos la categoría
        Next
        
        'FICHEROS
        ClearFiles
        Get fnum, , lngvar 'Número de ficheros
        For i = 1 To lngvar
            strvar = ReadStrFromFile(fnum) 'Nombre del fichero
            If Not Dir(makePathForProject(strvar)) = "" Then
                AddFile makePathForProject(strvar), False
            Else 'El fichero no existe
                MsgBox "Impossible to find " & strvar & ". The file will be removed from the " _
                        & "project.", vbCritical
            End If
        Next

        'DEVELOPERS
        devcol.Clear
        Get fnum, , lngvar 'Número de developers
        'strvar = ReadStrFromFile(fnum) 'Default dev
        For i = 1 To lngvar
            dev.name = ReadStrFromFile(fnum) 'Dev name
            dev.RealName = ReadStrFromFile(fnum) 'Real name
            dev.Mail = ReadStrFromFile(fnum) 'Mail
            devcol.Add dev.name, dev.RealName, dev.Mail
        Next
        'devcol.defaultDev = strvar
        'Establecemos el desarrollador por defecto
        With iniManager
            .Path = confFileName
            .Section = "OTHER"
            .Key = "defaultdev"
            devcol.defaultDev = .Value
        End With
    
        'TRACKER
        colTrackers.Clear
        Get fnum, , lngvar 'Número de trackers
        For i = 1 To lngvar
            'GENERAL
            tr.name = ReadStrFromFile(fnum) 'Nombre
            tr.description = ReadStrFromFile(fnum)  'Descripción
            tr.IconIndex = ReadLngFromFile(fnum) 'Icono

            'COLUMNAS
            tr.VisibleColumns = ReadLngFromFile(fnum) 'Columnas visibles
            tr.SortColumn = ReadLngFromFile(fnum) 'Columna por la cual se ordena
            tr.SortOrder = ReadLngFromFile(fnum) 'Ascendente o descendente
            'cols arrangement

            'COMPORTAMIENTO
            tr.AutoExpandSelItems = CBool(ReadIntFromFile(fnum)) 'Expandir sel items
            tr.ColorItemsByPriority = CBool(ReadIntFromFile(fnum)) 'Colorear prioridad
            tr.ColorOldItems = CBool(ReadIntFromFile(fnum)) 'Colorear old items
            tr.OldItemsDays = ReadIntFromFile(fnum) 'Old Items days

            'CATEGORÍAS
            tr.CategoryClear
            Get fnum, , lngvar2 'Número de categorías
            For j = 1 To lngvar2
                strvar = ReadStrFromFile(fnum)
                tr.AddCategory strvar
            Next

            'ELEMENTOS DEL TRACKER
            tr.Clear
            Get fnum, , lngvar2 'Número de items
            For j = 1 To lngvar2
                'GENERAL
                ti.Summary = ReadStrFromFile(fnum) 'Summary
                ti.DetailedDesc = ReadStrFromFile(fnum) 'Descripción detallada
                ti.module = ReadStrFromFile(fnum) 'Módulo
                ti.Category = ReadStrFromFile(fnum) 'Categoría
                ti.AssignedTo = ReadStrFromFile(fnum) 'Assignado a
                ti.SubmittedBy = ReadStrFromFile(fnum) 'Enviado por
                ti.Priority = ReadLngFromFile(fnum) 'Prioridad
                ti.Completed = ReadIntFromFile(fnum) '% completado
                ti.Closed = CBool(ReadIntFromFile(fnum)) 'Cerrado?
                ti.Hidden = CBool(ReadIntFromFile(fnum)) 'Eliminado?
                ti.Locked = CBool(ReadIntFromFile(fnum)) 'Bloqueado
                ti.DateCreated = CDate(ReadDblFromFile(fnum)) 'Fecha de creación
                ti.DateClosing = CDate(ReadDblFromFile(fnum)) 'Fecha de cierre
                ti.DateModified = CDate(ReadDblFromFile(fnum)) 'Fecha de última modificación
                tr.AddIndirect ti
            Next

            colTrackers.AddIndirect tr
        Next

        LoadProject = True 'Carga satisfactoria
    Else
        LoadProject = False 'Carga no satisfactoria
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

'    'Categorías por defecto
'    AddCategory "Source files", "*.prg|*.inc|*.h"
'    AddCategory "Image files", "*.map|*.fbm|*.png"
'    AddCategory "Graphics Libraries", "*.fpg|*.fgc"
    
''Unos desarrolladores de ejemplo (a borrar)
'    devcol.Add "Danko", "Darío Cutillas", "el_danko@hotmail.com"
'    devcol.Add "JaViS", "Javier Arias", "javisarias@gmail.com"
'    devcol.Add "Viator", "Darío Cutillas", "Viator@yahoo.es"
'    devcol.Add "Pepito Grillo", "Pepe Grill Osnado", "elpepe@yosoyyo.net"
End Sub

Private Sub Class_Terminate()
    Set m_colFiles = Nothing
    Set iniManager = Nothing
    Set devcol = Nothing
    Set colTrackers = Nothing
    Set m_colCategories = Nothing
End Sub

'Estas funciones escriben y leen una cadena de caracteres en un archivo
'abierto mediante Open usando un varspace de 4 bytes para la longitud
'de la cadena. La cadena solo es escrita si no está vacía
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
'Las siguientes funciones lee una variable de un archivo
'Se ha creado porque VB no permite leer directamente en una propiedad de una clase
'Se podría haber hecho una función general mediante VarType, pero restaría efeiciencia
'el tener que comprobar el tipo de variable
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

