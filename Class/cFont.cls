VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
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

' GDI+ Wrapper written by Paul
' http://www.syix.com/wpsjr1/index.html

' Handles Fonts and Font Families

Public gdip As cGDIP ' Reference for Error Handling

Dim font            As Long
Dim fontFamily      As Long
Dim lCurErr         As Long
Dim m_bBold         As Boolean
Dim m_lHeight       As Long
Dim m_bItalic       As Boolean
Dim m_sName         As String
Dim m_sngSize       As Single
Dim m_eStyle        As FontStyle
Dim m_bUnderline    As Boolean
Dim m_eUnit         As unit
Dim m_bStrikeout    As Boolean

'Public Current_Color As Long
'Public Current_Char As Long
'Public Zoom As Long
'Public Alpha As Boolean
'
'Dim Draw_Rect As RECT
Private sLastError As String



Public Function CreateFont(ByVal oFont As StdFont, Optional ByVal eFontStyle As FontStyle = FontStyleRegular, Optional ByVal eUnit As unit = UnitPoint) As Long
    m_eStyle = eFontStyle
    m_eUnit = eUnit
    m_sName = oFont.name
    m_bBold = oFont.Bold
    m_bItalic = oFont.Italic
    m_bUnderline = oFont.Underline
    m_sngSize = oFont.Size
    m_bStrikeout = oFont.Strikethrough
    'If CreateFontFamily(m_sName) <> Status.Ok Then ' Failed, try Generic
        CreateGenericFont m_sName
    'End If
    If fontFamily Then
        CreateFont = Execute(GdipCreateFont(fontFamily, CSng(oFont.Size), eFontStyle, eUnit, font))
    End If
End Function

Private Function CreateFontFamily(ByVal lFontName As String) As Long
    CreateFontFamily = Execute(GdipCreateFontFamilyFromName(lFontName, 0, fontFamily))
End Function

Private Function CreateGenericFont(ByVal lFontName As String) As Long
    fontFamily = 0
    CreateGenericFont = Execute(GdipGetGenericFontFamilySansSerif(fontFamily))
End Function

Public Function CreateFromHdc(ByVal lHdc As Long, ByVal oFont As StdFont) As Long
    m_sName = oFont.name
    m_bBold = oFont.Bold
    m_bItalic = oFont.Italic
    m_bUnderline = oFont.Underline
    m_sngSize = oFont.Size
    m_bStrikeout = oFont.Strikethrough
    CreateFromHdc = Execute(GdipCreateFontFromDC(lHdc, font))
    If CreateFromHdc = status.Ok Then
        GdipGetFontStyle font, m_eStyle
        GdipGetFontUnit font, m_eUnit
    End If
End Function

Public Sub Dispose()
    If fontFamily Then GdipDeleteFontFamily fontFamily
    If font Then GdipDeleteFont font
End Sub

Private Function Execute(ByVal lReturn As Long) As Long
    If lReturn Then
        lCurErr = lReturn
        If Not gdip Is Nothing Then gdip.Error lReturn, "cFont"
    Else
        lCurErr = 0
    End If
    Execute = lCurErr
End Function

Public Property Get Bold() As Boolean
    Bold = m_bBold
End Property

Public Property Get Handle() As Long
    Handle = font
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Get Italic() As Boolean
    Italic = m_bItalic
End Property

Public Property Get name() As String
    name = m_sName
End Property

Public Property Get Size() As Single
    Size = m_sngSize
End Property

Public Property Get StrikeOut() As Boolean
    StrikeOut = m_bStrikeout
End Property

Public Property Get Style() As FontStyle
    Style = m_eStyle
End Property

Public Property Get Underline() As Boolean
    Underline = m_bUnderline
End Property

Public Property Get unit() As unit
    unit = m_eUnit
End Property

Private Sub SetLastError(sError As String)
    sLastError = sError
End Sub

Public Function GetLastError() As String
    GetLastError = sLastError
End Function

Private Sub Class_Terminate()
    Dispose
End Sub


