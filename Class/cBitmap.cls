VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' GDI+ Wrapper written by Paul
' http://www.syix.com/wpsjr1/index.html
'
' Wrapper will Dispose on setting equal to Nothing (terminate event)
' And you are responsible for Disposing before any new bitmap is created
' Failure to Dispose properly will lead to various problems/crashes.
'
' Currently, functions that make a new bitmap are:
' CreateFromScan0, CreateFromHBITMAP and LoadFromFile

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As picDesc, riid As CLSID, ByVal fOwn As Long, ByRef lplpvObj As Any)

Public gdip As cGDIP ' Reference for Error Handling

Private bitmap          As Long
Private lCurErr         As Long
Private lPixelFormat    As Long
Private m_lFrameCount   As Long
Private m_lWidth        As Long
Private m_lHeight       As Long
Private m_lNumCodecs    As Long
Private rc              As RECT
Private tGuids()        As CLSID ' array of GUIDs for codecs
Private colCodecs       As Collection   ' search into codec index on file extension
Private stream          As IUnknown
Private bProperties()   As Byte
Private pi              As PROPERTYITEM
Private m_lPropSize     As Long
Private b()             As Byte

Private Function BitmapFromStream(ByVal IStream As IUnknown) As Long
  If Not (IStream Is Nothing) Then
    BitmapFromStream = Execute(GdipLoadImageFromStream(IStream, bitmap))
    If BitmapFromStream = status.Ok Then
      GetPixelFormat
      GetDimension
    End If
  End If
End Function

Public Function CreateFromScan0(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal ePixelFormats As PixelFormats, ByRef lDataPtr As Long)
  CreateFromScan0 = Execute(GdipCreateBitmapFromScan0(lWidth, lHeight, lStride, ePixelFormats, ByVal lDataPtr, bitmap))
  GetPixelFormat
  GetDimension
End Function

' Assumes a zero-based byte array containing an entire image is passed
Public Function CreateFromStream(ByRef bIn() As Byte) As Long
  CreateStreamOnHGlobal bIn(0), False, stream
  CreateFromStream = BitmapFromStream(stream)
End Function

Public Function CreateFromHBitmap(ByVal hBitmap As Long) As Long
  CreateFromHBitmap = Execute(GdipCreateBitmapFromHBITMAP(hBitmap, 0, bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Function CreateFromGraphics(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lGraphicHandle As Long) As Long
  CreateFromGraphics = Execute(GdipCreateBitmapFromGraphics(lWidth, lHeight, lGraphicHandle, bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Function CreateFromClone(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBitmapHandle As Long, ByVal ePixelFormat As PixelFormats) As Long
  CreateFromClone = Execute(GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, ePixelFormat, lBitmapHandle, bitmap))
  GetPixelFormat
  GetDimension
End Function

'ADDED FEATURE FOR FB2 PROJECT
Public Function CreateFromGdiDib8bpp(bInfo256 As BITMAPINFO256, ByVal lData As Long)
    CreateFromGdiDib8bpp = GdipCreateBitmapFromGdiDib(bInfo256, ByVal lData, bitmap)
    GetPixelFormat
    GetDimension
End Function

'ADDED FEATURE FOR FB2 PROJECT
Public Function CreateFromGdiDib16bpp(bInfo As BITMAPINFO, ByVal lData As Long)
    CreateFromGdiDib16bpp = GdipCreateBitmapFromGdiDib(bInfo, ByVal lData, bitmap)
    GetPixelFormat
    GetDimension
End Function


' I purposely did not use GdipCreateBitmapFromResource for now,
' since it doesn't work on anything but bitmap resources.
' This should work on all GDI+ supported image types
' Store images as "CUSTOM" resources with the resource editor addin
Public Function CreateFromResource(ByVal lResID As Long) As Long
  Dim b() As Byte ' Changed to module wide scope, 02/14/2004
  b = LoadResData(lResID, "CUSTOM")
  CreateStreamOnHGlobal b(0), False, stream
  CreateFromResource = BitmapFromStream(stream)
End Function

Public Function CreateThumbFromImage(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBitmapHandle As Long) As Long
  CreateThumbFromImage = Execute(GdipGetImageThumbnail(lBitmapHandle, lWidth, lHeight, bitmap))
  GetPixelFormat
  GetDimension
End Function

Public Property Get CurErr() As Long
  CurErr = lCurErr
End Property

Public Sub Dispose()
  If bitmap Then GdipDisposeImage bitmap
  Set stream = Nothing
End Sub

Private Function EnumEncoders() As Long
  Dim lNumEncoders  As Long
  Dim lEncoderSize  As Long
  Dim lError        As Long
  Dim b()           As Byte
  Dim i             As Long
  Dim codecs()      As ImageCodecInfo
  
  lError = GdipGetImageEncodersSize(lNumEncoders, lEncoderSize)
  If lError = 0 Then
    ReDim codecs(lNumEncoders - 1)
    ReDim b(lEncoderSize - 1)
    
    lError = GdipGetImageEncoders(lNumEncoders, lEncoderSize, b(0))
    If lError = 0 Then
      RtlMoveMemory codecs(0), b(0), lNumEncoders * LenB(codecs(0))
      ReDim tGuids(lNumEncoders - 1)
      m_lNumCodecs = lNumEncoders
      Set colCodecs = Nothing
      Set colCodecs = New Collection
      
      Do While lNumEncoders
        lNumEncoders = lNumEncoders - 1
        tGuids(lNumEncoders) = codecs(lNumEncoders).CLSID
        ParseOnChar StringFromPointerW(codecs(lNumEncoders).pwszFilenameExtension), ";", lNumEncoders
      Loop
    Else
      Error lError
    End If
  Else
    Error lError
  End If
End Function

Friend Property Get EncoderGuid(ByVal lIndex As Long) As CLSID
  EncoderGuid = tGuids(lIndex)
End Property

Private Function Execute(ByVal lReturn As Long) As Long
  If lReturn Then
    lCurErr = lReturn
    If Not gdip Is Nothing Then gdip.Error lReturn, "cBitmap"
  Else
    lCurErr = 0
  End If
  Execute = lCurErr
End Function

' do not compare this to a boolean
'  returns -1 for not found, 0-positive GUID index for found
Public Function ExtensionExists(ByRef sKey As String) As Long
  On Error GoTo errorhandler
  ExtensionExists = True ' invalid index
  
  If Not colCodecs Is Nothing Then
    ExtensionExists = colCodecs.item(sKey)
  End If
  
  Exit Function
errorhandler:
  ' exit silently
End Function

' number of frames in a multi-page tiff or animated gif
Public Property Get FrameCount() As Long
  FrameCount = m_lFrameCount
End Property

Public Function LoadFromFile(ByVal sFilename As String) As Long
  LoadFromFile = Execute(GdipLoadImageFromFile(sFilename, bitmap))
  GetPixelFormat
  GetDimension
End Function

' Returns the number of bytes from the stream in the array
' 0 if failure
Friend Function GetBitmapStream(ByRef bBmpBits() As Byte, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
  Dim stream As IStream
  Dim lBytesRead As Long
  Dim curMax As Currency
  
  ' True allows the global memory to be freed automatically,
  ' after the destruction of the IStream
  If CreateStreamOnHGlobal(ByVal 0, True, stream) = 0 Then
    If Execute(GdipSaveImageToStream(bitmap, stream, tEncoder, ByVal lEncParamPointer)) = 0 Then
      
      ' Here is an alternate method of getting information from the stream
      ' It's not nearly as versitile, and liable to be inaccurate,
      ' as GlobalSize might be rounded.
    
      ' I thought I'd include it, as the IStream interface need not be
      ' defined to use this code, simply; Dim stream as IUnknown
    
      ' Dim p As Long
      ' Dim hGlob As Long
      ' Dim lSizeGlob As Long
    
      ' GetHGlobalFromStream stream, p
      ' RtlMoveMemory hGlob, ByVal p, 4 ' could also GlobalLock here to deref and lock, remember to unlock later
      ' lSizeGlob = GlobalSize(hGlob)
      ' If lSizeGlob Then
      '   ReDim bBmpBits(lSizeGlob - 1)
      '   RtlMoveMemory bBmpBits(0), ByVal hGlob, lSizeGlob
      ' End If

      ' Note: you could seek anywhere in the stream and get any portion
      ' of the bitmap stream.
      stream.Seek 0@, STREAM_SEEK_END, curMax ' get the size of the stream
      curMax = curMax * 10000 ' adjust for int64

      stream.Seek 0@, STREAM_SEEK_SET, 0@ ' return to the beginning

      If curMax > 0 Then
        ReDim bBmpBits(curMax - 1)
        stream.Read bBmpBits(0), curMax, lBytesRead ' note that lBytesRead could overflow, please fix this if you're reading more than 2 Gigs :P
        GetBitmapStream = lBytesRead
      End If
    End If
  End If

  Set stream = Nothing
End Function

Private Sub GetDimension()
  Dim sngWidth  As Single
  Dim sngHeight As Single
  
  Execute GdipGetImageDimension(bitmap, sngWidth, sngHeight)
  m_lWidth = sngWidth
  m_lHeight = sngHeight
  rc.Right = m_lWidth
  rc.Bottom = m_lHeight
End Sub

Friend Sub GetPalette(ByVal lPalPointer As Long, ByVal lNumColors As Long)
  Dim lSize As Long
  
  lSize = (lNumColors * 4) + 8
  Execute GdipGetImagePalette(bitmap, ByVal lPalPointer, lSize)
End Sub

Private Sub GetPixelFormat()
  Execute GdipGetImagePixelFormat(bitmap, lPixelFormat)
End Sub

Public Property Get Handle() As Long
  Handle = bitmap
End Property

' So far this should only be changed from cGraphics.MakeTransparent
Public Property Let Handle(ByVal lHandle As Long)
  bitmap = lHandle
End Property

Public Property Get Height() As Long
  Height = m_lHeight
End Property

' Useful for Animated GIF, gets the delay for a given frame (0 based)
Public Function GetDelay(ByVal lFrame As Long) As Long
  Dim vDelay As Variant
  
  GetPropertyItem PropertyTagFrameDelay, vDelay, lFrame
  GetDelay = vDelay
End Function

Public Function GetLoopCount() As Long
  Dim vLoopCount As Variant
  
  If GetPropertyItem(PropertyTagLoopCount, vLoopCount, 0) = status.Ok Then
    GetLoopCount = vLoopCount
  End If
End Function

Public Function GetPropertyItem(ByVal eTag As PropertyTags, ByRef vValue As Variant, ByVal lPropertyIndex As Long) As Long
  Dim lSize As Long
  
  GetPropertyItem = Execute(GdipGetPropertyItemSize(bitmap, eTag, lSize))
  If GetPropertyItem = status.Ok Then
    ReDim bProperties(lSize - 1)
    GetPropertyItem = Execute(GdipGetPropertyItem(bitmap, eTag, lSize, bProperties(0)))
    If GetPropertyItem = status.Ok Then
      RtlMoveMemory pi, bProperties(0), LenB(pi) ' 2 bytes padding in memory
      If pi.value And pi.length > 0 Then
        ReDim bProperties(pi.length - 1)
        RtlMoveMemory bProperties(0), ByVal pi.value, pi.length
        m_lPropSize = pi.length
        If pi.Type = PropertyTagTypeByte Or pi.Type = PropertyTagTypeShort Or pi.Type = PropertyTagTypeLong Then
          vValue = GetPropertyNumeric(pi.Type, lPropertyIndex)
        ElseIf pi.Type = PropertyTagTypeASCII Or pi.Type = PropertyTagTypeRational Then
          'vValue = GetPropertyString(pi.Type, lPropertyIndex)
        Else
          Debug.Print "Error - Unhandled PropertyTagType"
        End If
      End If
    End If
  End If
End Function

Private Function GetPropertyNumeric(ByVal eTagtype As PropertyTagTypes, ByVal lPropIndex As Long) As Long
  Dim lSize As Long
  Dim lIndex As Long
  
    Select Case eTagtype
      Case PropertyTagTypeByte
        lSize = 1
      
      Case PropertyTagTypeShort
        lSize = 2
      
      Case PropertyTagTypeLong
        lSize = 4
    End Select
    
    lIndex = lPropIndex * lSize
    If lIndex >= 0 And lIndex < m_lPropSize Then
      RtlMoveMemory GetPropertyNumeric, bProperties(lIndex), lSize
    End If
End Function

'Private Function GetPropertyString(ByVal eTagtype As PropertyTagTypes, ByVal lPropIndex As Long) As String
'  If lPropIndex >= 0 And lPropIndex < m_lPropCount Then
'    ' todo
'  End If
'End Function

Friend Function GetPixelARGB(ByVal X As Long, ByVal Y As Long) As ARGB
  Execute GdipBitmapGetPixel(bitmap, X, Y, GetPixelARGB)
End Function

Public Function GetPixelLong(ByVal X As Long, ByVal Y As Long) As Long
  Execute GdipBitmapGetPixel(bitmap, X, Y, GetPixelLong)
End Function

Private Function IPictureFromBitmap(ByVal lBitmap As Long) As IPicture
  Dim lError  As Long
  Dim hBitmap As Long
  Dim picdes  As picDesc

  If bitmap = 0 Then Exit Function
    
  lError = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
    
  If hBitmap <> 0 Then
    picdes.cbSizeOfStruct = Len(picdes)
    picdes.picType = vbPicTypeBitmap
    picdes.hGdiObj = hBitmap
    OleCreatePictureIndirect picdes, StringToGuid(IPictureCLSID), True, IPictureFromBitmap
  Else
    Error lError
  End If
  ' dispose of bitmap in caller
End Function

Friend Function LockBits(ByRef rc As GDIPlus.RECT, ByVal eMode As ImageLockMode, ByVal eFormat As PixelFormats, ByRef bd As BitmapData) As Long
  LockBits = Execute(GdipBitmapLockBits(bitmap, rc, eMode, eFormat, bd))
End Function

Public Function MakeTransparent(ByVal lColor As Long) As Long
  Dim bmp As cBitmap
  Dim gfx As cGraphics
  Set bmp = New cBitmap
  Set gfx = New cGraphics
  Dim rc As RECT
  Dim lAttrib As Long
  Dim lTmpHandle As Long

  If bmp.CreateFromScan0(m_lWidth, m_lHeight, 0, PixelFormat32bppARGB, 0) = Ok Then
    If gfx.GetGraphicsContext(bmp.Handle) = Ok Then
      If gfx.Clear = Ok Then
        rc.Right = m_lWidth
        rc.Bottom = m_lHeight
        
        If GdipCreateImageAttributes(lAttrib) = Ok Then
          If GdipSetImageAttributesColorKeys(lAttrib, ColorAdjustTypeDefault, 1, lColor, lColor) = Ok Then
            If gfx.DrawImageRectRectIAttr(bitmap, 0, 0, m_lWidth, m_lHeight, 0, 0, m_lWidth, m_lHeight, lAttrib) = Ok Then
              GdipDisposeImageAttributes lAttrib
              lTmpHandle = bitmap
              bitmap = bmp.Handle
              bmp.Handle = lTmpHandle
            End If
          End If
        End If
        
      End If
    End If
  End If
  
  Set gfx = Nothing
  Set bmp = Nothing
End Function

' included instead of using Split() for the VB5 set :)
Private Sub ParseOnChar(ByRef sIn As String, ByRef sChar As String, ByVal lGuidIndex As Long)
  Dim lStartPosition As Long
  Dim lFoundPosition As Long
  Dim sItem          As String
  
  lFoundPosition = InStr(sIn, sChar)
  lStartPosition = 1
  
  Do While lFoundPosition
    sItem = Mid$(sIn, lStartPosition, lFoundPosition - lStartPosition)
    colCodecs.Add lGuidIndex, sItem
    lStartPosition = lFoundPosition + 1
    lFoundPosition = InStr(lStartPosition, sIn, sChar)
  Loop
  
  sItem = Trim$(Mid$(sIn, lStartPosition))
  If LenB(sItem) Then colCodecs.Add lGuidIndex, sItem
End Sub

Public Property Get Picture() As StdPicture
  Set Picture = IPictureFromBitmap(bitmap)
End Property

Public Property Get PixelFormat() As PixelFormats
  PixelFormat = lPixelFormat
End Property

Friend Property Get Rectangle() As RECT
  Rectangle = rc
End Property

Public Function RotateFlip(ByVal eFlipType As RotateFlipType) As Long
  RotateFlip = Execute(GdipImageRotateFlip(bitmap, eFlipType))
End Function


' At the time I wrote this class there were only two multi-page formats supported
' Consider adding an enum if more formats arise
Public Function SelectActiveFrame(ByVal lFrame As Long, Optional ByVal lIsGIF As Boolean = True) As Long
  Dim cid As CLSID
  If lIsGIF Then
    cid = StringToGuid(FrameDimensionTime)
  Else ' TIFF
    cid = StringToGuid(FrameDimensionPage)
  End If
  SelectActiveFrame = Execute(GdipImageSelectActiveFrame(bitmap, cid, lFrame))
  If m_lFrameCount = 0 Then GdipImageGetFrameCount bitmap, cid, m_lFrameCount
  GetDimension ' frames can be different sizes
End Function

Public Sub SetTransColor(ByVal lIndex As Byte)
  Dim gpal As ColorPalette256
  
  If bitmap Then
    GetPalette VarPtr(gpal), 256
    gpal.Entries(lIndex).Alpha = 0
    SetPalette VarPtr(gpal)
  End If
End Sub

Friend Function SaveToFile(ByVal sFilename As String, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
  SaveToFile = Execute(GdipSaveImageToFile(bitmap, sFilename, tEncoder, ByVal lEncParamPointer))
End Function

Public Function SetPalette(ByVal lPalPointer As Long) As Long
  SetPalette = Execute(GdipSetImagePalette(bitmap, ByVal lPalPointer))
End Function

Private Function StringFromPointerW(ByVal lPointer As Long) As String
  Dim lLength As Long
  
  If lPointer Then
    lLength = lstrlenW(lPointer)
    StringFromPointerW = Space$(lLength)
    RtlMoveMemory ByVal StrPtr(StringFromPointerW), ByVal lPointer, lLength * 2
  End If
End Function

Friend Function StringToGuid(ByRef sGuid As String) As CLSID
  CLSIDFromString sGuid, StringToGuid
End Function

Friend Function UnlockBits(ByRef bd As BitmapData) As Long
  UnlockBits = Execute(GdipBitmapUnlockBits(bitmap, bd))
End Function

Public Property Get Width() As Long
  Width = m_lWidth
End Property

Private Sub Class_Initialize()
  EnumEncoders
End Sub

Private Sub Class_Terminate()
  Dispose
End Sub
