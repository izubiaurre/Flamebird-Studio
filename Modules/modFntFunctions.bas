Attribute VB_Name = "modFntFunctions"
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

Public Const c_FNT_DIV_Magic = "fnt"
Public Const c_FNT_BENNU_Magic = "fnx"
Public Const c_PAL_Magic = "pal"
Public Const c_MAP_Magic = "map"
Public Const c_FPG_Magic = "fpg"
Public Const c_FNT__1 = &H1A '26        ' MS-DOS header
Public Const c_FNT__2 = &HD  '13
Public Const c_FNT__3 = &HA  '10
Public Const c_FNT__4 = &H0  '0

'0 for old DIV fonts
'Under bennu:
'1 for 1bpp fonts
'8 for 8bpp   "
'16 for 8bpp   "
'32 for 8bpp   "
Public Const c_FNT_version_0 = &H0
Public Const c_FNT_version_1 = &H1
Public Const c_FNT_version_8 = &H8
Public Const c_FNT_version_16 = &H10
Public Const c_FNT_version_32 = &H20


Public Type t_FNT_Magic
    magic As String * 3 ' must be "fnt" or "fnx"
    MS_DOS_header(3) As Byte
    version As Byte
End Type

Public Type t_FNT_Palette_Entry
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type t_FNT_Palette
    Entries(255) As t_FNT_Palette_Entry
    UnusedBytes(579) As Byte        ' gradients
End Type

Public charset As Long              ' charset_cp850 = 1; charset_ISO8859 = 0

Public Type t_FNT_Char_Header       ' old DIV char header
    width As Long
    height As Long
    Vertical_Offset As Long
    File_Offset As Long

    w_Width As Long
End Type

Public Type t_FNT_bennu_Char_HEADER ' new header (for Bennu)
    width As Long
    height As Long
    total_width As Long
    total_height As Long
    x_offset As Long
    y_offset As Long
    data_offset As Long
End Type

Public Type t_FNT_Char
    Header As t_FNT_Char_Header
    Data() As Byte
End Type

Public Type t_FNT_bennu_Char
    Header As t_FNT_bennu_Char_HEADER
    Data() As Byte
End Type

