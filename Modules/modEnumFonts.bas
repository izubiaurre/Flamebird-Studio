Attribute VB_Name = "modEnumFonts"
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

Private cbo As ComboBox

Private Const FIXED_PITCH As Long = 1
Private Const ANSI_CHARSET As Long = 0
Private Const DEVICE_FONTTYPE As Long = &H2
Private Const RASTER_FONTTYPE As Long = &H1
Private Const TRUETYPE_FONTTYPE As Long = &H4
Private Const LF_FULLFACESIZE As Long = 64
Private Const LF_FACESIZE As Long = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Private Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Private Type ENUMLOGFONT
    elfLogFont As LOGFONT
    elfFullName(LF_FULLFACESIZE) As Byte
    elfStyle(LF_FACESIZE) As Byte
End Type

Private Const TMPF_TRUETYPE As Long = &H4
Private Declare Function EnumFontFamilies Lib "gdi32.dll" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumFonts Lib "gdi32.dll" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long

Private Function EnumFontFamProc(nlf As ENUMLOGFONT, ntm As NEWTEXTMETRIC, ByVal nFontType As Long, ByVal lpData As Long) As Long
    Dim name As String
    
    If (nlf.elfLogFont.lfPitchAndFamily And 3&) = FIXED_PITCH And nlf.elfLogFont.lfCharSet = ANSI_CHARSET Then
        name = GetNameFromByteArray(nlf.elfLogFont.lfFaceName)
        cbo.AddItem name
    End If

    EnumFontFamProc = -1
End Function

Private Function GetNameFromByteArray(src() As Byte) As String
    Dim t As String, zeropos As Integer
    t = StrConv(CStr(src), vbUnicode)
    zeropos = InStr(t, Chr(0))
    If zeropos > 1 Then t = Left(t, zeropos - 1)
    GetNameFromByteArray = t
End Function

Public Function FixedPitchFontsToCombo(hdc As Long, cboWhere As ComboBox)
    Set cbo = cboWhere
    Call EnumFontFamilies(hdc, vbNullString, AddressOf EnumFontFamProc, 0)
    Set cbo = Nothing
End Function
