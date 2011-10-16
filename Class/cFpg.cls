VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Option Compare Binary
Option Base 0

Private Const MSG_FPG_FILENOTFOUND = "File not found!"
Private Const MSG_FPG_DEPTHUNSUPPORTED = "Unsupported depth!"
Private Const MSG_FPG_UNRECOGNIZEDFILE = "Unrecognized file format!"
Private Const MSG_FPG_COULDNOTOPEN = "Impossible to open the file!"
Private Const MSG_FPG_COULDNOTLOADPALETTE = "Could not read palette!"
Private Const MSG_FPG_ERRORLOADINGMAP = "An error ocurred when trying to load the maps"
Private Const MSG_FPG_FPGNOTAVAILABLE = "There is no FPG info available"
Private Const MSG_FPG_CANNOTADD_UNKNOWN = "Unknown reason "
Private Const MSG_FPG_CANNOTADD_DIFFDEPTH = "Map is not of the same depth"

Private Const PALETTE_SIZE = 768
Private Const PALETTE_USELESS = 576
Private Const F_NCPOINTS = &HFFF
Private Const F_ANIMATION = &H1000
Private Const MAX_MAPS = 1000

'Collection of maps
Private m_Maps As Collection
'FPG member variables
Private m_Palette() As Byte
Private m_Depth As Integer
'Informative member variables
'Private m_FilePath As String 'Read-only
Private m_Available As Boolean
Private m_LastError As String
Private m_Codes(MAX_MAPS - 1) As Integer 'To determine if a code already exists
Private m_NextFreeCodeStart As Long

Public Function New16() As Long
    Dispose
    
    m_Depth = 16
    Set m_Maps = New Collection
    m_Available = True
    
    New16 = -1
End Function

'-------------------------------------------------------------------------------------
'PUBLIC PROPERTIES
'-------------------------------------------------------------------------------------
Public Property Get Available() As Boolean
    Available = m_Available
End Property

Public Property Get map(lIndex As Long) As cMap
    If Not (lIndex + 1) > m_Maps.count Then
        Set map = m_Maps.item(lIndex + 1)
    Else
      '  MsgBox "Item " & CStr(i) & " doesn't exist", vbCritical
    End If
End Property

'Public Property Get FilePath() As String
'    FilePath = m_FilePath
'End Property

Public Property Get MapCount() As Long
    MapCount = m_Maps.count
End Property

Public Property Get Depth() As Integer
    Depth = m_Depth
End Property

'-------------------------------------------------------------------------------------
'FUNCTION: getAfterIndex()
'DESCRIPTION: Given a code, looks for the Index in the collection after which a map
'   with that code should be inserted to keep maps sorted.
'RETURNS: The index after which map should be inserted
'-------------------------------------------------------------------------------------
Private Function getAfterIndex(startCode As Long) As Long
    Dim i As Long
    Dim lAfterIndex As Long
    
    lAfterIndex = 0
    For i = startCode To 9 Step -1
        If m_Codes(i) = -1 Then
            lAfterIndex = IndexByCode(i)
            Exit For
        End If
    Next
    
    getAfterIndex = lAfterIndex
End Function


Private Function getBeforeIndex(startCode As Long) As Long
    Dim i As Long
    Dim lBeforeIndex As Long
    
    lBeforeIndex = 0
    For i = startCode To MAX_MAPS - 1
        If m_Codes(i) = -1 Then
            lBeforeIndex = IndexByCode(i)
            Exit For
        End If
    Next
    
    getBeforeIndex = lBeforeIndex
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: FreeCode()
'DESCRIPTION: Look for next free code after the start
'RETURNS: The next free code (0 a MAX_MAPS-1) or -1 if there is none
'-------------------------------------------------------------------------------------
Public Function FreeCode(Optional start As Long = 0) As Long
    Dim i As Integer, lCode As Long
    lCode = -1
    If start < 0 Then start = 0
    If start > MAX_MAPS - 1 Then start = MAX_MAPS - 1
    If m_Maps.count > 0 And m_Maps.count < MAX_MAPS Then
        For i = start To MAX_MAPS - 1
            If Not Exists(i) Then lCode = i: Exit For
        Next
    End If
    
    FreeCode = lCode
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: IndexByCode()
'DESCRIPTION: Get the index of the map whose code is lCode
'RETURNS: Index of the map in the collection (base 1) or 0 if not found
'-------------------------------------------------------------------------------------
Public Function IndexByCode(lCode As Long) As Long
    Dim l As Long
    Dim lIndex As Long
    
    lIndex = 0
    'Look for the map in the collection
    For l = 1 To m_Maps.count
        If m_Maps.item(l).Code = lCode Then
            lIndex = l
            Exit For
        End If
    Next
    IndexByCode = lIndex
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: Add()
'DESCRIPTION: Adds a Map to the fpg.
'RETURNS: -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Add(map As cMap, Optional replace As Boolean = False) As Long
    Dim lSucceded As Long
    Dim afterIndex As Long
    
    lSucceded = 0
    If Not map Is Nothing Then
        If map.Available And map.Depth = m_Depth Then
            If Not Exists(map.Code) Then
                afterIndex = getAfterIndex(map.Code)
                If afterIndex <> 0 Then
                    m_Maps.Add map, , , afterIndex
                Else
                    If m_Maps.count > 1 Then
                        m_Maps.Add map, , 1
                    Else
                        m_Maps.Add map
                    End If
                End If
                lSucceded = -1
                m_Codes(map.Code) = -1
            ElseIf replace = True Then 'The map already exists and replace=true
                'Note that the funcion IndexByCode will search for the first map with the
                'code given and so the new map must be added after the old one
                m_Maps.Add map, , , IndexByCode(map.Code)  'Add the new map after the old one
                m_Maps.item(IndexByCode(map.Code)).Dispose 'Unload the old map
                m_Maps.Remove (IndexByCode(map.Code)) 'Remove the old map from the collection
                lSucceded = -1
            End If
        End If
    End If
    If lSucceded = 0 Then
        If map.Depth = m_Depth Then
            SetLastError MSG_FPG_CANNOTADD_UNKNOWN
        Else
            SetLastError MSG_FPG_CANNOTADD_DIFFDEPTH
        End If
    End If
    Add = lSucceded
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: Remove()
'DESCRIPTION: Removes the map whose code is lCode.
'RETURNS: -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Sub Remove(lCode As Long)
    Dim l As Long
    If Exists(lCode) Then
        l = IndexByCode(lCode)
        If l > 0 Then m_Maps.Remove (l): m_Codes(lCode) = 0
    End If
End Sub
'-------------------------------------------------------------------------------------
'FUNCTION: Exists()
'DESCRIPTION: Check the existence of a map whose code is iCode
'RETURNS: True or false
'-------------------------------------------------------------------------------------
Public Function Exists(ByVal lCode As Long) As Boolean
    If lCode > MAX_MAPS - 1 Then lCode = MAX_MAPS - 1
    If lCode < 0 Then lCode = 0
    Exists = IIf(m_Codes(lCode) = 0, False, True)
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: Dispose()
'DESCRIPTION: Resets the fpg data
'RETURNS: Nothing
'-------------------------------------------------------------------------------------
Public Sub Dispose()
    Dim i As Integer
    'Remove all maps
    While (m_Maps.count > 0)
        m_Maps.Remove (1)
    Wend
    Erase m_Palette
    Erase m_Codes
    'm_FilePath = ""
    SetLastError ""
    m_Depth = 0
    m_Available = False
End Sub
'-------------------------------------------------------------------------------------
'FUNCTION: Load()
'DESCRIPTION: Loads an FPG file
'RETURNS: -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Load(sFile As String) As Long
    Dim lFile As Long
    Dim afterIndex As Long, beforeIndex As Long
    Dim mapclen As Long, cparraylen As Long
    'MAP INFO
    Dim lMCode As Long, lMLen As Long
    Dim sMDesc As String * 32, sMFileName As String * 12
    Dim lMWidth As Long, lMHeight As Long
    Dim lMFlags As Long, iMCPoints() As Integer
    Dim iMData() As Integer, bMData() As Byte
    
    Dim magic As String * 8, paluseless(PALETTE_USELESS) As Byte
    Dim m As cMap
    Dim lSucceded As Long, lError As Long
    
    On Error GoTo ErrHandler
    
    lFile = gzopen(sFile, "rb")
    
    If Not (lFile = 0) Then
        Dispose
        'm_FilePath = sFile
        
        gzread lFile, ByVal magic, Len(magic)
        ' HEADER check
        If (StrComp(Left(magic, 3), F16_MAGIC) = 0) Then 'm16bits
            m_Depth = 16
        ElseIf StrComp(Left(magic, 3), FPG_MAGIC) = 0 Then 'm8bits
            m_Depth = 8
            'Read the palette
            ReDim m_Palette(PALETTE_SIZE) As Byte
            If (gzread(lFile, m_Palette(0), PALETTE_SIZE) < PALETTE_SIZE) Then lError = -1
            If (gzread(lFile, paluseless(0), PALETTE_USELESS) < PALETTE_USELESS) Then lError = -1
            If lError Then SetLastError MSG_FPG_COULDNOTLOADPALETTE
        Else
            lError = -1
            SetLastError MSG_FPG_UNRECOGNIZEDFILE
        End If
        
        If lError = 0 Then
            m_Available = True 'Assume everything will go ok
            'Read all maps
            Do
                'MAP HEADER
                sMDesc = ""
                gzread lFile, lMCode, 4
                If (gzeof(lFile) <> 0) Then Exit Do 'EOF
                gzread lFile, lMLen, 4
                gzread lFile, ByVal sMDesc, 32
                gzread lFile, ByVal sMFileName, 12
                gzread lFile, lMWidth, 4
                gzread lFile, lMHeight, 4
                
                'FLAGS & CPOINTS
                gzread lFile, lMFlags, 4
                mapclen = (lMFlags And F_NCPOINTS) * 4 'Each control point are 4 bytes (two shorts)
                cparraylen = IIf(mapclen > 0, mapclen / 2 - 1, 1) '0 CP means that the default CP must be set to the center of the image
                ReDim iMCPoints(cparraylen) As Integer
                iMCPoints(0) = lMWidth \ 2  'Default X Coord.
                iMCPoints(1) = lMHeight \ 2  'Default Y Coord.
                If (lMFlags And F_NCPOINTS) Then 'Read CPs when there are more than 1 cp
                    If (gzread(lFile, iMCPoints(0), mapclen) < mapclen) Then lError = 1
                End If
                
                'ANIMATION (DISABLED)
                
                'DATA
                Set m = New cMap
                If m_Depth = 16 Then
                    ReDim iMData(lMWidth * lMHeight - 1) As Integer
                    gzread lFile, iMData(0), lMWidth * lMHeight * 2
                    'Create the map
                    If m.CreateFromStream16(AsciiZToString(sMDesc), lMCode, _
                                    lMWidth, lMHeight, iMData) = 0 Then
                        lError = -1
                    End If
                ElseIf m_Depth = 8 Then
                    ReDim bMData(lMWidth * lMHeight - 1) As Byte
                    gzread lFile, bMData(0), lMWidth * lMHeight
                    'Create the map
                    If m.CreateFromStream8(AsciiZToString(sMDesc), lMCode, _
                                    lMWidth, lMHeight, bMData, m_Palette) = 0 Then
                        lError = -1
                    End If
                End If
                If lError = 0 Then 'Set the MAP Control points data
                    If Not m.SetCPData(iMCPoints) Then lError = -1
                End If
                'Add the map to the collection
                If lError = 0 Then
                    afterIndex = getAfterIndex(m.Code)
                    beforeIndex = getBeforeIndex(m.Code)
                    If afterIndex <> 0 Then
                        m_Maps.Add m, , , afterIndex
                    ElseIf beforeIndex <> 0 Then
                        m_Maps.Add m, , beforeIndex
                    Else
                        m_Maps.Add m
                    End If
                    'An error could occur here if the Code of the map is negative
                    'or greater than MAX_MAPS-1 but this should never happen since we
                    'assume that the fpg file is well created
                    m_Codes(m.Code) = -1
                Else
                    SetLastError MSG_FPG_ERRORLOADINGMAP & ": " & m.GetLastError
                    Set m = Nothing
                    Dispose
                    Exit Do
                End If
                Set m = Nothing
            Loop While (gztell(lFile) <> -1)
        End If
        gzclose (lFile)
    Else 'File not found
        SetLastError MSG_FPG_FILENOTFOUND
    End If
    
    Load = IIf(lError = 0, -1, 0)
ErrHandler:
    If Err.Number > 0 Then ShowError "cFpg.Load"
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: Save()
'DESCRIPTION: Saves an FPG file
'RETURNS: -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Save(sFile As String) As Long
    Dim lFile As Long
    Dim mapclen As Long, mapLen As Long, cparraylen As Long
    'MAP INFO
    Dim lMCode As Long, lMLen As Long
    Dim sMDesc As String * 32, sMFileName As String * 12
    Dim lMWidth As Long, lMHeight As Long
    Dim lMFlags As Long, iMCPoints() As Integer
    Dim iMData() As Integer, bMData() As Byte
    Dim b As Byte
    
    Dim i As Integer
    Dim magic As String * 3, paluseless(PALETTE_USELESS) As Byte
    Dim m As cMap
    Dim lSucceded As Long, lError As Long
    
    If m_Available = True Then
        lFile = gzopen(sFile, "wb")
        
        If Not (lFile = 0) Then
            lError = 0 'Supose no error
            'm_FilePath = sFile
            'HEADER
            'File type
            If m_Depth = 16 Then
                magic = "f16"
            ElseIf m_Depth = 8 Then
                magic = "fpg"
            End If
            'Hex data
            Call gzwrite(lFile, ByVal magic, 3)
            b = &H1A: Call gzwrite(lFile, b, 1)
            b = &HD: Call gzwrite(lFile, b, 1)
            b = &HA: Call gzwrite(lFile, b, 1)
            b = &H0: Call gzwrite(lFile, b, 1)
            b = &H0: Call gzwrite(lFile, b, 1)
            'PALETTE
            If m_Depth = 8 Then '8bits-->Write palette
                gzwrite lFile, m_Palette(0), PALETTE_SIZE
                gzwrite lFile, paluseless(0), PALETTE_USELESS
            End If
            'MAPS
            For i = 1 To m_Maps.count
                Set m = m_Maps(i)
                'MAP HEADER
                gzwrite lFile, m.Code, 4 'Code
                lMLen = 64 + (m.CPointsCount - 1) * 4 + m.Width * m.Height
                gzwrite lFile, lMLen, 4 'Len
                sMDesc = StringToAsciiZ(m.description, 32)
                gzwrite lFile, ByVal sMDesc, 32 'Description
                sMFileName = StringToAsciiZ("SINARCHIVO", 12)
                gzwrite lFile, ByVal sMFileName, 12 'FileName
                gzwrite lFile, m.Width, 4 'Width
                gzwrite lFile, m.Height, 4 'Height
                
                'FLAGS & CPOINTS
                lMFlags = m.CPointsCount
                If m.CPointsCount = 1 Then 'If the only CP is the center of the image, don't write it
                    If (m.ControlPoint(0)(0) = m.Width / 2) And (m.ControlPoint(0)(1) = m.Height / 2) Then lMFlags = 0
                End If
                gzwrite lFile, lMFlags, 4
                mapclen = (lMFlags And F_NCPOINTS) * 4
                If (lMFlags And F_NCPOINTS) Then
                    ReDim iMCPoints(m.CPointsCount * 4) As Integer
                    If m.GetCPData(iMCPoints) Then 'Write CP
                        gzwrite lFile, iMCPoints(0), mapclen
                    Else 'Error obtaining CP
                        lError = -1
                        SetLastError "Error obtaining CP Data: " & m.GetLastError
                        Exit For
                    End If
                End If
                
                'ANIMATION (DISABLED)
                
                'DATA
                If m_Depth = 16 Then
                    mapLen = m.Width * m.Height * 2
                    ReDim iMData(mapLen - 1) As Integer
                    If m.GetData16(iMData()) Then
                        gzwrite lFile, iMData(0), mapLen
                    Else
                        lError = -1
                    End If
                ElseIf m_Depth = 8 Then
                    mapLen = m.Width * m.Height
                    ReDim bMData(mapLen - 1) As Byte
                    If m.GetData8(bMData()) Then
                        gzwrite lFile, bMData(0), mapLen
                    Else
                        lError = -1
                    End If
                End If
                If lError Then 'Could not write mapdata
                    SetLastError "Cannot set map data: " & m.GetLastError
                    Exit For
                End If
            Next
            gzclose (lFile)
        Else 'File not found
            SetLastError MSG_FPG_FILENOTFOUND
            lError = -1
        End If
    Else 'Unexisting FPG
        SetLastError MSG_FPG_FPGNOTAVAILABLE
        lError = -1
    End If
    
    Save = IIf(lError = 0, -1, 0)
End Function

'---------------------------------------------------------------------
Private Sub SetLastError(sError As String)
    m_LastError = sError
End Sub

Public Function GetLastError() As String
    GetLastError = m_LastError
End Function
'--------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Maps = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Maps = Nothing
End Sub
