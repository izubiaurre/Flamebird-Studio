VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGraphics"
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

Public gdip As cGDIP ' Reference for Error Handling

Private graphics As Long
Private lDC      As Long
Private lCurErr  As Long

Public Function Clear() As Long
  Clear = Execute(GdipGraphicsClear(graphics, Colors.Transparent))
End Function

Public Function CreateFromHdc(ByVal lHdc As Long) As Long
  CreateFromHdc = Execute(GdipCreateFromHDC(lHdc, graphics))
End Function

Public Function CreateFromHWND(ByVal lHwnd As Long) As Long
  CreateFromHWND = Execute(GdipCreateFromHWND(lHwnd, graphics))
End Function

Public Function GetGraphicsContext(ByVal lBitmap As Long) As Long
  GetGraphicsContext = Execute(GdipGetImageGraphicsContext(lBitmap, graphics))
End Function

Public Property Get CurErr() As Long
  CurErr = lCurErr
End Property

Public Sub Dispose()
  GdipDeleteGraphics graphics
  graphics = 0
End Sub

Public Function DrawImageI(ByVal lBmpHandle As Long, ByVal X As Long, ByVal Y As Long) As Long
  DrawImageI = Execute(GdipDrawImageI(graphics, lBmpHandle, X, Y))
End Function

Public Function DrawImageRectI(ByVal lBmpHandle As Long, ByVal X As Long, ByVal Y As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
  DrawImageRectI = Execute(GdipDrawImageRectI(graphics, lBmpHandle, X, Y, lWidth, lHeight))
End Function

Public Function DrawImageRectRectI(ByVal lBmpHandle As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal lDstWidth As Long, ByVal lDstHeight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal lSrcWidth As Long, ByVal lSrcHeight As Long) As Long
  DrawImageRectRectI = Execute(GdipDrawImageRectRectI(graphics, lBmpHandle, dstx, dsty, lDstWidth, lDstHeight, srcx, srcy, lSrcWidth, lSrcHeight, unit.UnitPixel))
End Function

Public Function DrawImageRectRectIAttr(ByVal lBmpHandle As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal lDstWidth As Long, ByVal lDstHeight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal lSrcWidth As Long, ByVal lSrcHeight As Long, ByVal lAttribute As Long) As Long
  DrawImageRectRectIAttr = Execute(GdipDrawImageRectRectI(graphics, lBmpHandle, dstx, dsty, lDstWidth, lDstHeight, srcx, srcy, lSrcWidth, lSrcHeight, unit.UnitPixel, lAttribute))
End Function

Public Function DrawString(ByVal sText As String, ByVal lColor As Long, ByVal oFont As cFont, Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0) As Long
  Dim rcF As RECTF
  Dim format As Long
  Dim brush As Long
  
  rcF.Left = lLeft
  rcF.Top = lTop
  
  GdipCreateSolidFill lColor, brush
  GdipStringFormatGetGenericDefault format
  
  DrawString = Execute(GdipDrawString(graphics, sText, Len(sText), oFont.Handle, rcF, format, brush))
  
  GdipDeleteBrush brush
  GdipDeleteStringFormat format
End Function

Private Function Execute(ByVal lReturn As Long) As Long
  If lReturn Then
    lCurErr = lReturn
    If Not gdip Is Nothing Then gdip.Error lReturn, "cFont"
  Else
    lCurErr = 0
  End If
  Execute = lCurErr
End Function

Public Function GetDC() As Long
  Execute GdipGetDC(graphics, lDC)
End Function

Public Property Get Handle() As Long
  Handle = graphics
End Property

Public Sub ReleaseDC()
  Execute GdipReleaseDC(graphics, lDC)
  lDC = 0
End Sub

Public Function ResetTransform() As Long
  ResetTransform = Execute(GdipResetWorldTransform(graphics))
End Function

Public Function RotateTransform(ByVal sngAngle As Single, Optional ByVal eOrder As MatrixOrder = MatrixOrderPrepend) As Long
  RotateTransform = Execute(GdipRotateWorldTransform(graphics, sngAngle, eOrder))
End Function

Public Function SetInterpolationMode(ByVal eMode As InterpolationMode) As Long
  SetInterpolationMode = Execute(GdipSetInterpolationMode(graphics, eMode))
End Function

Public Function SetTextRenderingHint(ByVal eHint As TextRenderingHint) As Long
  SetTextRenderingHint = Execute(GdipSetTextRenderingHint(graphics, eHint))
End Function

Public Function TranslateTransform(ByVal dx As Single, ByVal dy As Single, Optional ByVal eOrder As MatrixOrder = MatrixOrderPrepend) As Long
  TranslateTransform = Execute(GdipTranslateWorldTransform(graphics, dx, dy, eOrder))
End Function

Private Sub Class_Terminate()
  If graphics Then
    If lDC Then ReleaseDC
    Dispose
  End If
End Sub
