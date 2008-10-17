Attribute VB_Name = "modFntFunctions"
Public Const c_FNT_Magic = "fnt"
Public Const c_PAL_Magic = "pal"
Public Const c_MAP_Magic = "map"
Public Const c_FPG_Magic = "fpg"
Public Const c_FNT_version_1 = &H1A
Public Const c_FNT_version_2 = &HD
Public Const c_FNT_version_3 = &HA
Public Const c_FNT_version_4 = &H0
Public Const c_FNT_version_5 = &H0

Public Type t_FNT_Magic
    magic As String * 3
    Version(4) As Byte
End Type

Public Type t_FNT_Palette_Entry
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type t_FNT_Palette
    Entries(255) As t_FNT_Palette_Entry
    UnusedBytes(579) As Byte
End Type

Public Type t_FNT_Char_Header
    Width As Long
    Height As Long
    Vertical_Offset As Long
    File_Offset As Long

    w_Width As Long
End Type

Public Type t_FNT_Char
    Header As t_FNT_Char_Header
    data() As Byte
End Type

