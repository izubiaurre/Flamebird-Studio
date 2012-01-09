VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMap"
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
Option Compare Binary
Option Base 0

' NOTE: MapBlt, MakeBitmap, and other functions returns 0 on error, -1 otherwise
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Error constants
Private Const ERR_MAP_ALREADY16BPP = "This map is already 16bpp depth"
Private Const ERR_MAP_COULDNOTMAKEBMP = "Bitmap could not be created"
Private Const ERR_MAP_NOTLOADED = "Map not loaded"
Private Const ERR_MAP_CANNOTBEDRAWN = "The map cannot be drawn"
Private Const ERR_MAP_FILENOTFOUND = "file not found"
Private Const ERR_MAP_FILETRUNCATED = "file truncated"
Private Const ERR_MAP_FILEINCOMPATIBLE = "incompatible file"
Private Const ERR_MAP_LASTCPOINT = "Last control point cannot be removed"
Private Const ERR_MAP_INVALIDCPOINTINDEX = "Invalid control point index"
Private Const ERR_MAP_ALREADY8BPP = "This map is already 8bpp depth"
Private Const ERR_MAP_CANNOTSELECTDC = "The device context could not be selected"
Private Const ERR_MAP_CANNOTCREATECOPYBMP = "Could not create a copy of the bitmap"
Private Const ERR_MAP_BMPNOTAVAILABLE = "Cannot create a map from an empty bitmap"
Private Const ERR_MAP_MAPNOTAVAILABLE = "The map isn't ready"
Private Const ERR_MAP_STREAMDATAINVALID = "Not a valid stream data"

Private Const PALETTE_SIZE = 768
Private Const PALETTE_USELESS = 576

Private Const F_NCPOINTS = &HFFF
Private Const F_ANIMATION = &H1000

Private Const BI_RGB = 0&

'class member variables (with property implementation)
Private m_Width As Long
'm_Height and m_Width are long here but uint16 in map file format!
Private m_Height As Long
Private m_Code As Long
Private m_Description As String * 32    'AsciiZ format
'Private m_FilePath As String 'Read-only
Private m_Depth As Integer
Private m_AnimationLength As Integer    'Read-only
Private m_AnimationSpeed As Integer     'Read-only

'member variables (without property implementation)
Private m_Palette() As Byte
Private m_Data16() As Integer           'Map data in 565 format
Private m_Data8() As Byte               'Map data in indexed format
Private m_CPoints() As Integer
Private m_Animation() As Integer
Private m_NFlags As Integer
Private m_Frames As Integer

'Other
'The cMap object store, MAP information ready to be written,
'the information graphic object (called object bitmap), or both.
'In order to perform read and write operations of MAPs, the variable m_MapAvailable
'must be true. In order to perform painting operations graphics and open/save
'the map in different formats, m_BitmapAvailable must be true
Private m_MapAvailable As Boolean       'Informs if is builded the info in MAP format
Private m_BitmapAvailable As Boolean

Private m_Bitmap As cBitmap
Private BitmapData() As Byte            'For use of 1, 4 and 8 bit maps
Private BitmapData555() As Integer      'For use of 16 bit maps
Private sLastError As String            'Last error-message

Private WithEvents gdip As cGDIP
Attribute gdip.VB_VarHelpID = -1

'--------------------------------------------------------------------------------------
'Properties implementation
'--------------------------------------------------------------------------------------



'Returns width of the map
Property Get Width() As Long
    Width = m_Width
End Property

'Returns height of the map
Property Get Height() As Long
    Height = m_Height
End Property

'Returns code of the map
Property Get Code() As Long
    Code = m_Code
End Property
Property Let Code(ByVal newVal As Long)
    m_Code = newVal
End Property

'Returns description of the map
Property Get description() As String
    description = AsciiZToString(m_Description)
End Property
Property Let description(ByVal newVal As String)
    m_Description = StringToAsciiZ(newVal, Len(m_Description))
End Property

'Property Get FilePath() As String
'    FilePath = m_FilePath
'End Property

'Returns depth (bpp) of the map
Property Get Depth() As Integer
    Depth = m_Depth
End Property

'Property Get IsLoaded() As Boolean
'    IsLoaded = (m_BitmapAvailable Or m_MapAvailable) 'retorna verdadero si se ha cargado un mapa
'End Property

Property Get Available() As Boolean
    Available = m_MapAvailable
End Property

Property Get ControlPoint(Index As Integer) As Integer()
    Dim res(1) As Integer
    
    If m_MapAvailable Then
        If Index < ((UBound(m_CPoints) + 1) / 2) And Not (Index < 0) Then
            res(0) = m_CPoints(Index * 2)
            res(1) = m_CPoints(Index * 2 + 1)
            ControlPoint = res
        End If
    End If
End Property

Property Get CPointsCount() As Integer
    If m_MapAvailable Then
        CPointsCount = (UBound(m_CPoints) + 1) / 2
    Else
        CPointsCount = 0
    End If
End Property

Property Let ControlPoint(Index As Integer, cp() As Integer)
    If m_MapAvailable Then
        If (Index < ((UBound(m_CPoints) + 1) / 2)) And ((UBound(cp) - LBound(cp)) = 1) Then
            If Not (Index < 0) Then
                m_CPoints(Index * 2) = cp(LBound(cp))
                m_CPoints(Index * 2 + 1) = cp(UBound(cp))
            End If
        End If
    End If
End Property

Public Function NewCPoint(cpX As Integer, cpY As Integer) As Long
    Dim lSucceded As Long
    If m_MapAvailable = True Then
        ReDim Preserve m_CPoints(UBound(m_CPoints) + 2) As Integer
        m_CPoints(UBound(m_CPoints) - 1) = cpX
        m_CPoints(UBound(m_CPoints)) = cpY
        lSucceded = 1
    Else
        lSucceded = 0
        SetLastError ERR_MAP_MAPNOTAVAILABLE
    End If
    NewCPoint = lSucceded
End Function

Public Function RemoveCPoint(Index As Integer) As Long
    Dim aux() As Integer
    Dim cnt As Long, lSucceded As Long
    
    If m_MapAvailable Then
        If CPointsCount > 1 Then
            If Index < ((UBound(m_CPoints) + 1) / 2) And Not (Index < 0) Then
                'Sorts the rest elements in an auxiliar array
                ReDim aux(UBound(m_CPoints) - 2) As Integer
                For cnt = 0 To Index * 2 - 1
                    aux(cnt) = m_CPoints(cnt)
                Next
                For cnt = Index * 2 + 2 To UBound(m_CPoints)
                    aux(cnt - 2) = m_CPoints(cnt)
                Next
                
                ReDim m_CPoints(UBound(aux)) As Integer
                m_CPoints = aux
                lSucceded = -1
            Else
                lSucceded = 0
                SetLastError ERR_MAP_INVALIDCPOINTINDEX
            End If
        Else
            lSucceded = 0
            SetLastError ERR_MAP_LASTCPOINT
        End If
    Else
        lSucceded = 0
        SetLastError ERR_MAP_MAPNOTAVAILABLE
    End If
    
    RemoveCPoint = lSucceded
End Function

Property Get palette() As Byte()
    Dim palcopy(PALETTE_SIZE - 1) As Byte
    If UBound(m_Palette) >= LBound(m_Palette) Then
        CopyMemory ByVal VarPtr(palcopy(0)), ByVal VarPtr(m_Palette(0)), PALETTE_SIZE
    End If
    palette = palcopy
End Property
'--------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------
'Methods
Public Sub Dispose()
    Erase m_Data8
    Erase m_Data16
    Erase BitmapData
    Erase BitmapData555
    Erase m_Palette
    Erase m_CPoints
    
    m_MapAvailable = False
    m_BitmapAvailable = False
    Set m_Bitmap = Nothing
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION:      Draw()
'DESCRIPTION:   Draws the map into a DC
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Draw(hdc As Long, lX As Long, lY As Long, Optional lWidth As Long = 0, _
                        Optional lHeight As Long = 0, Optional bTrans As Boolean = False) As Long
    Dim graphics As cGraphics
    Dim lSucceded As Long, lError As Long
    Dim bmp As cBitmap      'Stores a copy of the bitmap to paint
    
    If m_MapAvailable = True And m_BitmapAvailable = False Then 'Map exists but the bitmap has not been builded
        m_BitmapAvailable = MakeBitmap
    End If
    
    If m_BitmapAvailable Then
        lError = 1
        
        Set graphics = New cGraphics    'Graphic object
        Set graphics.gdip = gdip
        Set bmp = New cBitmap
        
        lError = bmp.CreateFromClone(m_Width, m_Height, m_Bitmap.Handle, m_Bitmap.PixelFormat)
        If Not lError Then
            If bTrans = True Then 'Sets trans
                If m_Depth = 8 Then
                    bmp.SetTransColor (0) 'First palette color
                ElseIf m_Depth = 16 Then
                    bmp.MakeTransparent (0) 'Black transparent color
                End If
                drawTransBack hdc, lWidth, lHeight
            End If
        
            lError = graphics.CreateFromHdc(hdc) 'Select the DC to draw
            If Not lError Then
                If Not (bmp Is Nothing) Then 'Draw the map
                    graphics.SetInterpolationMode InterpolationModeNearestNeighbor
                    lError = graphics.DrawImageRectI(bmp.Handle, lX, lY, lWidth, lHeight)
                End If
                
                graphics.Dispose
                bmp.Dispose
                Set bmp = Nothing
                
                If Not lError Then
                    lSucceded = -1 'Success
                Else
                    lSucceded = 0
                    SetLastError ERR_MAP_CANNOTBEDRAWN
                End If
            Else
                lSucceded = 0
                SetLastError ERR_MAP_CANNOTSELECTDC
            End If
        Else
            lSucceded = 0
            SetLastError ERR_MAP_CANNOTCREATECOPYBMP
        End If
        Set graphics = Nothing
    Else
        lSucceded = 0
        SetLastError ERR_MAP_MAPNOTAVAILABLE
    End If
    
    Draw = lSucceded
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      MakeBitmap()
'DESCRIPTION:   Create a bitmap (windows format) from map info
'RETURNS:       -1 if no error, otherwise 0. Check GetLastError
'-------------------------------------------------------------------------------------
Private Function MakeBitmap() As Long
    Dim bInfo256 As BITMAPINFO256, bInfo As BITMAPINFO, bInfo16 As BITMAPINFO16, bInfo2 As BITMAPINFO2
    Dim cnt As Long, cnt2 As Long
    Dim scanline As Long            'Size of a scanline
    Dim lSucceded As Long, lError As Long
    
    If Not (m_Bitmap Is Nothing) Then Set m_Bitmap = Nothing
    
    lError = -1 'Supose error
    If m_MapAvailable = True Then
        Set m_Bitmap = New cBitmap 'Create new bitmap
        Set m_Bitmap.gdip = gdip   'Reference for errors
        
        Select Case m_Depth
        Case 8
            With bInfo256.bmiHeader
                .biBitCount = 8: .biCompression = BI_RGB: .biPlanes = 1
                .biClrUsed = 256: .biClrImportant = 0: .biSize = LenB(bInfo256.bmiHeader)
                .biWidth = m_Width: .biHeight = -m_Height
            End With
            
            'Prepare the palette
            For cnt = 0 To 255
                bInfo256.bmiColors(cnt).rgbRed = m_Palette(cnt * 3) * 4
                bInfo256.bmiColors(cnt).rgbGreen = m_Palette(cnt * 3 + 1) * 4
                bInfo256.bmiColors(cnt).rgbBlue = m_Palette(cnt * 3 + 2) * 4
            Next
            
            'Calculate the scanline in aligned in 4 bytes
            scanline = (m_Width + 4 * Abs(CLng((m_Width Mod 4) > 0))) - (m_Width And 3) 'The size of the scanline must be aligned in 4 bytes
            ReDim BitmapData(CLng(scanline) * CLng(m_Height) - 1) As Byte
            If m_Width = scanline Then 'La información del mapa está alineada a 4 bytes
                BitmapData = m_Data8
            Else 'The map info is not aligned to 4 bytes
                'Prepara el mapa
                For cnt = 0 To m_Height - 1
                    CopyMemory BitmapData(cnt * scanline), m_Data8(cnt * m_Width), m_Width
                Next
            End If
            lError = m_Bitmap.CreateFromGdiDib8bpp(bInfo256, VarPtr(BitmapData(0)))    'Create bitmap
            
        Case 16
            With bInfo.bmiHeader
                .biBitCount = 16: .biCompression = BI_RGB: .biPlanes = 1
                .biSize = LenB(bInfo.bmiHeader)
                .biWidth = m_Width: .biHeight = -m_Height
            End With
            
            'We must align the data to 4 bytes.
            ' As each pixel are 2 bytes,only in case of odd map width
            scanline = (m_Width Mod 2) + m_Width
            'Convert from 16 bits 565 format to 555
            ReDim BitmapData555(CLng(scanline) * CLng(m_Height) - 1) As Integer
            For cnt = 0 To m_Height - 1
                For cnt2 = 0 To m_Width - 1
                    BitmapData555(cnt2 + cnt * scanline) = convtab_565to555(m_Data16(cnt2 + cnt * m_Width))
                Next
            Next
            
            lError = m_Bitmap.CreateFromGdiDib16bpp(bInfo, VarPtr(BitmapData555(0))) 'Create the bitmap
        End Select
    End If
    
    If lError Then
        lSucceded = 0
        SetLastError ERR_MAP_COULDNOTMAKEBMP 'There was an error
        Dispose
    Else
        lSucceded = -1
    End If
    
    MakeBitmap = lSucceded
End Function

Public Function ConvertTo16bpp() As Long
    Dim cnt As Long, cnt2 As Long
    Dim aux As Long
    Dim lSucceded As Long
    
    If m_MapAvailable Then
        If m_Depth = 8 Then
            'Calculates data of the 16 bits map
            ReDim m_Data16(CLng(m_Width) * CLng(m_Height) - 1)
            For cnt = 0 To m_Height - 1
                For cnt2 = 0 To m_Width - 1
                    aux = m_Data8(cnt2 + cnt * m_Width) * 3
                    m_Data16(cnt2 + cnt * m_Width) = convtab_666to565(m_Palette(aux), m_Palette(aux + 1), m_Palette(aux + 2))
                Next
            Next
            m_Depth = 16
            lSucceded = MakeBitmap
            If lSucceded Then
                Erase BitmapData
                Erase m_Palette
                Erase m_Data8
            End If
        Else
            lSucceded = 0
            SetLastError ERR_MAP_ALREADY16BPP
        End If
    Else
        lSucceded = 0
        SetLastError ERR_MAP_MAPNOTAVAILABLE
    End If
    
    ConvertTo16bpp = lSucceded
End Function

Public Function ConvertTo8bpp() As Long
    Dim i As Long, j As Long
    Dim clrs() As Long, clrsfreq() As Long
    Dim clrCnt As Integer, clrUsed As Integer
    Dim ordenado As Byte
    Dim aux As Long
    Dim dicClr As New Dictionary
    Dim dicNearestClr As Dictionary 'Hash table to find fast the near colors of a palette
    Dim lSucceded As Long
    Dim lMaplen As Long
    
    If m_MapAvailable Then
        If m_Depth = 16 Then
            lMaplen = CLng(m_Width) * CLng(m_Height)
        
            Set dicNearestClr = New Dictionary
            
            'Calculates the frequency fo the colors in a image
            For i = 0 To m_Height - 1
                For j = 0 To m_Width - 1
                    If dicClr.Exists(m_Data16(i * m_Width + j)) Then
                        dicClr.item(m_Data16(i * m_Width + j)) = dicClr.item(m_Data16(i * m_Width + j)) + 1
                    Else
                        dicClr.Add m_Data16(i * m_Width + j), 1
                    End If
                Next
            Next
            
            'Creates the color arrays and frequencies with so much elements as they are
            clrCnt = dicClr.count
            ReDim clrs(clrCnt - 1) As Long
            ReDim clrsfreq(clrCnt - 1) As Long
            
            Dim var As Variant
            j = 0
            For Each var In dicClr.keys
                clrs(j) = SInt16ToUint32(CInt(var)) 'Conversion: color to unsigned integer
                clrsfreq(j) = CLng(dicClr(var))
                j = j + 1
            Next
              
            Set dicClr = Nothing
        
            'Sorts the elements by freq (BUBLE SORT! very slow)
            ordenado = 0: i = 1
            While (ordenado = 0)
                ordenado = 1 'Supose it's sorted
                For j = 0 To clrCnt - i - 1
                    If clrsfreq(j) < clrsfreq(j + 1) Then
                        aux = clrs(j)
                        clrs(j) = clrs(j + 1)
                        clrs(j + 1) = aux
                        aux = clrsfreq(j)
                        clrsfreq(j) = clrsfreq(j + 1)
                        clrsfreq(j + 1) = aux
                        ordenado = 0
                    End If
                Next
                i = i + 1
            Wend
            
            Erase clrsfreq
            
            'Create palette from more freq. colors
            clrUsed = IIf(clrCnt > 256, 256, clrCnt) 'Colores usados
            
            ReDim m_Palette(PALETTE_SIZE - 1) As Byte
            For i = 0 To clrUsed - 1
                m_Palette(i * 3) = ((clrs(i) And &HF800) \ 2 ^ 11) * 2
                m_Palette(i * 3 + 1) = (clrs(i) And &H7E0) \ 2 ^ 5
                m_Palette(i * 3 + 2) = (clrs(i) And &H1F) * 2
            Next
            
            'Create hash table with the already known colors
            For i = 0 To clrUsed - 1
                dicNearestClr.Add UInt32ToSInt16(clrs(i)), i
            Next
            Erase clrs
            
            'Creates the info. of the map looking the closest colors for each pixel in the palette
            'If there are less than 256 colors, we use the hashtable without check if the element is in this table
            ReDim m_Data8(lMaplen - 1) As Byte
            If clrUsed <> clrCnt Then 'There are more colors than the used ones
                For i = 0 To lMaplen - 1
                    If Not dicNearestClr.Exists(m_Data16(i)) Then
                        dicNearestClr.Add m_Data16(i), NearestColor(m_Palette, clrUsed, m_Data16(i))
                    End If
                    m_Data8(i) = dicNearestClr(m_Data16(i))
                Next
            Else
                For i = 0 To lMaplen - 1
                    m_Data8(i) = dicNearestClr(m_Data16(i))
                Next
            End If
            
            Set dicNearestClr = Nothing
            
            m_Depth = 8
            lSucceded = MakeBitmap
            If lSucceded Then
                Erase m_Data16
                Erase BitmapData555
            End If
        Else
            lSucceded = 0
            SetLastError ERR_MAP_ALREADY8BPP
        End If
    Else
        lSucceded = 0
        SetLastError ERR_MAP_MAPNOTAVAILABLE
    End If
    
    ConvertTo8bpp = lSucceded
End Function

Private Function NearestColor(pal() As Byte, ByVal nColors As Integer, ByVal clr565S As Integer) As Byte
    Dim distance As Long, mindistance As Long
    Dim clrIndex As Integer
    Dim i As Long
    Dim r As Byte, g As Byte, b As Byte
    Dim lClr565U As Long
    
    lClr565U = SInt16ToUint32(clr565S)
    
    r = ((lClr565U And &HF800) \ 2 ^ 11) * 2
    g = (lClr565U And &H7E0) \ 2 ^ 5
    b = (lClr565U And &H1F) * 2
    
    mindistance = &H3000 + 1 'MAximun distance, 64^2*3
    For i = 0 To nColors - 1
        distance = (CLng(pal(i * 3)) - r) ^ 2 + (CLng(pal(i * 3 + 1)) - g) ^ 2 + (CLng(pal(i * 3 + 2)) - b) ^ 2
        If distance < mindistance Then
            mindistance = distance
            clrIndex = i
        End If
    Next
    
    NearestColor = clrIndex
End Function

Private Sub drawTransBack(hdc As Long, lWidth As Long, lHeight As Long)
    Dim bmp As New cBitmap
    Dim graphics As New cGraphics
    Dim i As Long, j As Long
    Dim resFile As String
    Dim cx As Integer, cy As Integer
    
    'If FSO.FileExists(FSO.BuildPath(App.Path, "\Resources\backmaps.bmp")) Then
    bmp.LoadFromFile FSO.BuildPath(App.Path, "\Resources\backmaps.bmp")

    graphics.CreateFromHdc hdc
    graphics.Clear
    
    cx = lWidth \ bmp.Width + 1
    cy = lHeight \ bmp.Height + 1
    For i = 0 To cx
        For j = 0 To cy
            graphics.DrawImageRectI bmp.Handle, i * bmp.Width, j * bmp.Height, bmp.Width, bmp.Height
        Next
    Next
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION:      MakeMap()
'DESCRIPTION:   Creates a 16bpp map data from the pixels of a bitmap
'RETURNS:       -1 if no error, otherwise 0. Check GetLastError
'-------------------------------------------------------------------------------------
Private Function MakeMap() As Long
    Dim lSucceded As Long
    Dim i As Long, j As Long, clr As ARGB
    
    If Not m_Bitmap Is Nothing Then
        'Initial properties of the map
        m_Height = m_Bitmap.Height
        m_Width = m_Bitmap.Width
        m_Code = 0
        m_Description = StringToAsciiZ("Unnamed", Len(m_Description))
        ReDim m_CPoints(1) As Integer 'Initial Control Point
        m_CPoints(0) = m_Width / 2
        m_CPoints(1) = m_Height / 2
        'For now, load the image at 16 bits and convert to 8 if necessary
        m_Depth = 16
        'Convert info to 16 bits
        ReDim m_Data16(CLng(m_Height) * CLng(m_Width) - 1) As Integer
        For i = 0 To m_Height - 1
            For j = 0 To m_Width - 1
                clr = m_Bitmap.GetPixelARGB(j, i)
                If clr.Alpha < 255 / 2 Then
                    clr.Red = 0: clr.Green = 0: clr.Blue = 0
                End If
                m_Data16(i * m_Width + j) = UInt32ToSInt16(((clr.Red \ 8) * 2 ^ 11) Or ((clr.Green \ 4) * 2 ^ 5) Or CLng(clr.Blue \ 8))
            Next
        Next
        m_MapAvailable = True
        If m_Bitmap.PixelFormat = PixelFormat8bppIndexed Then lSucceded = ConvertTo8bpp
        
        lSucceded = -1
    Else
        lSucceded = 0
        SetLastError ERR_MAP_BMPNOTAVAILABLE
    End If
    
    m_MapAvailable = CBool(lSucceded)
    MakeMap = lSucceded
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      LoadBitmap()
'DESCRIPTION:   Loads a standar GDI+ file and store it as bitmap data
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Import(sFile As String) As Long
    Dim lError As Long
    
    Dispose 'Resets
    
    Set m_Bitmap = New cBitmap
    lError = m_Bitmap.LoadFromFile(sFile)
    
    If Not lError Then
        'm_FilePath = sFile
        MakeMap
        MakeBitmap
    Else
        m_BitmapAvailable = False
    End If
    
    Import = IIf(lError = 0, -1, 0)
End Function


'-------------------------------------------------------------------------------------
'FUNCTION:      LoadMap()
'DESCRIPTION:   Loads a map file
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Load(sFile As String) As Long
    Dim lFile As Long
    Dim mapLen As Long, mapclen As Long
    Dim magic As String * 8, paluseless(PALETTE_USELESS - 1) As Byte
    Dim cparraylen As Long
    Dim lError As Long, lSucceded As Long
    Dim iAux As Integer
    
    lError = 0
    
    lFile = gzopen(sFile, "rb")
    
    If Not (lFile = 0) Then
        Dispose 'Resets
        'HEADER
        gzread lFile, ByVal magic, Len(magic)
        gzread lFile, iAux, Len(iAux)
        m_Width = SInt16ToUint32(iAux)
        gzread lFile, iAux, Len(iAux)
        m_Height = SInt16ToUint32(iAux)
        gzread lFile, m_Code, Len(m_Code)
        gzread lFile, ByVal m_Description, Len(m_Description)
        
        'HEADER CHECK
        If (StrComp(Left(magic, 3), M16_MAGIC) = 0) Then    'm16bits
            mapLen = m_Width * m_Height * 2
            m_Depth = 16
        ElseIf StrComp(Left(magic, 3), MAP_MAGIC) = 0 Then  'm8bits
            mapLen = m_Width * m_Height
            m_Depth = 8
            ReDim m_Palette(PALETTE_SIZE - 1) As Byte
            If (gzread(lFile, m_Palette(0), PALETTE_SIZE) < PALETTE_SIZE) Then lError = -1
            If (gzread(lFile, paluseless(0), PALETTE_USELESS) < PALETTE_USELESS) Then lError = -1 'Useless info
        Else
            lError = -1
        End If
        
        If lError = 0 Then
            'NFLAGS explained
            'The integer of 16 bits m_NFlags, shares info about the Control Points and the animation
            'So, the first 12 bits are for the number od Control Points
            'and the 13th bit if exists animation (1) or not (0).
            
            'CONTROL POINTS
            gzread lFile, m_NFlags, 2
            mapclen = (m_NFlags And F_NCPOINTS) * 4     'Each Control Point is 4 bytes(2 short integers)
            cparraylen = IIf(mapclen > 0, mapclen / 2 - 1, 1) 'When number of CPs is 0, it's the center of the image
            ReDim m_CPoints(cparraylen)
            m_CPoints(0) = m_Width \ 2          'Coord. X of the cpoint by default
            m_CPoints(1) = m_Height \ 2         'Coord. Y of the cpoint by default
            If (m_NFlags And F_NCPOINTS) Then
                If (gzread(lFile, m_CPoints(0), mapclen) < mapclen) Then lError = -1
            End If
            
            'ANIMATION
            'By default this is disabled
            m_AnimationLength = 0
            m_AnimationSpeed = 50
            m_Frames = 1
            If (m_NFlags And F_ANIMATION) Then
                If (gzread(lFile, m_Frames, 2) < 2) Then lError = -1
                If (gzread(lFile, m_AnimationLength, 2) < 2) Then lError = -1
                If (gzread(lFile, m_AnimationSpeed, 2) < 2) Then lError = -1
                mapclen = 2 * m_AnimationLength
                ReDim m_Animation(mapclen + 2)
                If (gzread(lFile, m_Animation(0), mapclen) < mapclen) Then lError = -1
                mapLen = mapLen * m_Frames
            End If
            
            'MAP
            If m_Depth = 8 Then '8 bits maps
                ReDim m_Data8(mapLen - 1) As Byte
                If (gzread(lFile, m_Data8(0), mapLen) < mapLen) Then lError = -1
            ElseIf m_Depth = 16 Then 'Read the data if depth=16
                ReDim m_Data16(mapLen \ 2 - 1) As Integer
                If (gzread(lFile, m_Data16(0), mapLen) < mapLen) Then lError = -1
            End If
                
            If Not lError = 0 Then 'Opening errors
                Erase m_Palette: Erase m_Animation: Erase m_Data8: Erase m_Data16
                lSucceded = 0
                SetLastError ERR_MAP_FILETRUNCATED
            Else
                'm_FilePath = sFile
                lSucceded = -1 'Succeded
            End If
                
        Else 'Incompatible file
            lSucceded = 0
            SetLastError ERR_MAP_FILEINCOMPATIBLE
        End If
        
        Call gzclose(lFile)
    Else 'File not found
        lSucceded = 0
        SetLastError ERR_MAP_FILENOTFOUND
    End If
    
    If Not lSucceded Then
        m_MapAvailable = False
    Else
        m_MapAvailable = True
    End If
    
    Load = lSucceded
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      CreateFromStream16()
'DESCRIPTION:   Creates a 16bpp map from an stream
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function CreateFromStream16(ByVal sDescription As String, ByVal lCode As Long _
                , ByVal lWidth As Long, ByVal lHeight As Long, mapData() As Integer) As Long
    Dim lError As Long, mapLen As Long
    lError = 0
    mapLen = lWidth * lHeight * 2
    'Check Len of the data
    If lError = 0 Then
        If (UBound(mapData) - LBound(mapData)) < (mapLen / 2 - 1) Then
            lError = -1
            SetLastError "Stream data len insufficient"
        End If
    End If
    'Create the map
    If lError = 0 Then
        Dispose 'Reset
        m_Depth = 16
        m_Code = lCode
        m_Width = lWidth
        m_Height = lHeight
        ReDim m_Data16(mapLen) As Integer
        m_Data16 = mapData
        m_Description = StringToAsciiZ(sDescription, 32)
        'FLAGS
        m_NFlags = 0
        ReDim m_CPoints(1) As Integer
        m_CPoints(0) = m_Width \ 2      'Coord. X of the cpoint by default
        m_CPoints(1) = m_Height \ 2     'Coord. Y of the cpoint by default
        'ANIMATION
        
        m_MapAvailable = True
    End If
    
    CreateFromStream16 = Not lError
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      CreateFromStream8()
'DESCRIPTION:   Creates a map from an stream
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function CreateFromStream8(ByVal sDescription As String, ByVal lCode As Long _
                , ByVal lWidth As Long, ByVal lHeight As Long, mapData() As Byte, mapPalette() As Byte) As Long
                
    Dim lError As Long, mapLen As Long
    lError = 0
    'Check palette
    If (UBound(mapPalette) - LBound(mapPalette)) < (PALETTE_SIZE - 1) Then
        lError = -1
        SetLastError "Stream palette len insufficient"
    End If
    'Check Len of the data
    mapLen = lWidth * lHeight
    If lError = 0 Then
        If (UBound(mapData) - LBound(mapData)) < (mapLen - 1) Then
            lError = -1
            SetLastError "Stream data len insufficient"
        End If
    End If
    'Create the map
    If lError = 0 Then
        Dispose 'Reset
        m_Depth = 8
        m_Code = lCode
        m_Width = lWidth
        m_Height = lHeight
        ReDim m_Data8(mapLen) As Byte
        m_Data8 = mapData
        ReDim m_Palette(PALETTE_SIZE - 1) As Byte
        m_Palette = mapPalette
        m_Description = StringToAsciiZ(sDescription, 32)
        'FLAGS
        m_NFlags = 0
        ReDim m_CPoints(1) As Integer
        m_CPoints(0) = m_Width \ 2      'Coord. X of the cpoint by default
        m_CPoints(1) = m_Height \ 2     'Coord. Y of the cpoint by default
        'ANIMATION
        
        m_MapAvailable = True
    End If
    
    CreateFromStream8 = Not lError
End Function



'-------------------------------------------------------------------------------------
'FUNCTION:      Save()
'DESCRIPTION:   Saves a Map file
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Save(sFile As String) As Long
    Dim lFile As Long
    Dim mapLen As Long, mapclen As Long 'Len of the data of the map and the CP data
    Dim magic As String * 8, newNFlags As Integer
    Dim paluseless(PALETTE_USELESS - 1) As Byte, b As Byte, iAux As Integer
    Dim lSucceded As Long

    If m_MapAvailable = False And m_BitmapAvailable Then 'If exists info of the bitmap but not from the map,create the map
        MakeMap
    End If
    
    If m_MapAvailable = True Then
        lFile = gzopen(sFile, "wb1")
        
        If Not (lFile = 0) Then
            'm_FilePath = sFile
            
            'Write map header
            If m_Depth = 16 Then
                magic = "m16" '& &H1A & &HD & &HA
            ElseIf m_Depth = 8 Then
                magic = "map" '& &H1A & &HD & &HA
            End If
            
            Call gzwrite(lFile, ByVal magic, 3)
            b = &H1A: Call gzwrite(lFile, b, 1)
            b = &HD: Call gzwrite(lFile, b, 1)
            b = &HA: Call gzwrite(lFile, b, 1)
            b = &H0: Call gzwrite(lFile, b, 1)
            b = &H0: Call gzwrite(lFile, b, 1)
            'Width, height and description
            iAux = UInt32ToSInt16(m_Width)
            Call gzwrite(lFile, iAux, 2)
            iAux = UInt32ToSInt16(m_Height)
            Call gzwrite(lFile, iAux, 2)
            Call gzwrite(lFile, m_Code, 4)
            Call gzwrite(lFile, ByVal m_Description, 32)
        
            If m_Depth = 8 Then         '8bits-->Write palette
                gzwrite lFile, m_Palette(0), PALETTE_SIZE
                gzwrite lFile, paluseless(0), PALETTE_USELESS
                mapLen = m_Width * m_Height
            ElseIf m_Depth = 16 Then    '16bits
                mapLen = m_Width * m_Height * 2
            End If
        
            'Control points
            newNFlags = CPointsCount
            If CPointsCount = 1 Then 'If exists only one CPoint, look if it is the center (in this case don't write)
                If (m_CPoints(0) = m_Width / 2) And (m_CPoints(1) = m_Height / 2) Then newNFlags = 0
            End If
            gzwrite lFile, newNFlags, 2
            mapclen = (newNFlags And F_NCPOINTS) * 4
            If (newNFlags And F_NCPOINTS) Then
                gzwrite lFile, m_CPoints(0), mapclen
            End If
        
            'ANIMATION is disabled

            If m_Depth = 8 Then 'write the data if depth=8
                gzwrite lFile, m_Data8(0), mapLen
            Else                'write the data if depth=16
                gzwrite lFile, m_Data16(0), mapLen
            End If
 
            Call gzclose(lFile)
            
            m_NFlags = newNFlags
            'm_FilePath = sFile
            lSucceded = -1
        Else 'File not found
            lSucceded = 0
            SetLastError ERR_MAP_FILENOTFOUND
        End If
    Else 'Map not laoded
        lSucceded = 0
        SetLastError ERR_MAP_NOTLOADED
    End If
    
    Save = lSucceded
End Function

'-------------------------------------------------------------------------------------
'FUNCTION:      GetData8() & GetData16
'DESCRIPTION:   Retrieves the map data (copying it into an array arg)
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function GetData16(ByRef Data() As Integer) As Long
    Dim lSucceded As Long
    If (UBound(Data) - LBound(Data)) < m_Width * m_Height - 1 Then
        lSucceded = 0
        SetLastError "Insufficient array len to allocate data"
    Else
        lSucceded = -1
        Data = m_Data16
    End If
    GetData16 = lSucceded
End Function
Public Function GetData8(ByRef Data() As Byte) As Long
    Dim lSucceded As Long
    If (UBound(Data) - LBound(Data)) < m_Width * m_Height - 1 Then
        lSucceded = 0
        SetLastError "Insufficient array len to allocate data"
    Else
        lSucceded = -1
        Data = m_Data8
    End If
    GetData8 = lSucceded
End Function
'-------------------------------------------------------------------------------------
'FUNCTION:      GetCPData
'DESCRIPTION:   Retrieves the array of the CP
'RETURNS:       -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function GetCPData(ByRef cpData() As Integer) As Long
    Dim lSucceded As Long
    If UBound(cpData) - LBound(cpData) < UBound(m_CPoints) - LBound(m_CPoints) Then
        lSucceded = 0
        SetLastError "Insufficient array len to allocate CP data"
    Else
        cpData = m_CPoints
        lSucceded = -1
    End If
    GetCPData = lSucceded
End Function
Public Function SetCPData(ByRef cpData() As Integer) As Long
    Dim lSucceded As Long
    Dim lLenCP As Long
    lLenCP = (UBound(cpData) - LBound(cpData) + 1)
    'Error if the len of the array is not an even number
    If (lLenCP Mod 2) <> 0 Then
        lSucceded = 0
        SetLastError "CPdata len must be even"
    Else
        m_NFlags = lLenCP / 2
        ReDim m_CPoints(lLenCP - 1) As Integer
        m_CPoints = cpData
        lSucceded = -1
    End If
    SetCPData = lSucceded
End Function

Private Sub SetLastError(sError As String)
    sLastError = sError
End Sub

Public Function GetLastError() As String
    GetLastError = sLastError
End Function

Private Sub gdip_Error(ByVal lGdiError As Long, ByVal sErrorDesc As String, ByVal sClassName As String)
    MsgBox "GDIP Error in: " & sClassName & vbTab & "Error Number: " & CStr(lGdiError) & vbTab & "Error Description: " & sErrorDesc
End Sub

Private Sub Class_Initialize()
    Set gdip = New cGDIP 'Initialize GDI+
End Sub

Private Sub Class_Terminate()
    Dispose
    Set gdip = Nothing
End Sub
