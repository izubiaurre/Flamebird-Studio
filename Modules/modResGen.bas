Attribute VB_Name = "modResGen"
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

' **************************************************************
' Metodos y propiedades comunes a todos los archivos de recurso
' **************************************************************
Private Const MSG_ERROR_ADDING_MAP = "The map could not be added: "

Public Const MAP_MAGIC = "map"
Public Const M16_MAGIC = "m16"
Public Const PAL_MAGIC = "pal"
Public Const FNT_MAGIC = "fnt"
Public Const FPG_MAGIC = "fpg"
Public Const F16_MAGIC = "f16"

'Public Const BACKTRANS_WIDTH = 16
'Public Const BACKTRANS_HEIGHT = 16

Public Type MAPINFOHEADER
    sDescription As String
    iDepth As Integer
    lCode As Long
    lWidth As Long
    lHeight As Long
    lFlags As Long
End Type

Public convtab_565to555(-32768 To 32767) As Integer
Public convtab_666to565(63, 63, 63) As Integer

Public Function Convert565to555(ByVal color As Integer) As Integer
    Dim r As Long, g As Long, b As Long
    
    r = color And &HF800&
    g = color And &H7E0&
    b = color And &H1F&
    
    r = ((r \ 2 ^ 1))
    g = ((g \ 2 ^ 6) * 2 ^ 5)
    
    Convert565to555 = (r Or g Or b)
End Function

Public Function Convert666to565(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Integer
    r = r \ 2
    g = g
    b = b \ 2
    Convert666to565 = UInt32ToSInt16((((r * 2 ^ 11) Or (g * 2 ^ 5) Or b)))
End Function

Public Function initConversionTables()
    Dim cnt As Long, cnt2 As Long, cnt3 As Long
    '565to555
    For cnt = LBound(convtab_565to555) To UBound(convtab_565to555)
        convtab_565to555(CInt(cnt)) = Convert565to555(CInt(cnt))
    Next
    '666to555
    For cnt = 0 To UBound(convtab_666to565, 1)
        For cnt2 = 0 To UBound(convtab_666to565, 2)
            For cnt3 = 0 To UBound(convtab_666to565, 3)
                'convtab_666to555(cnt, cnt2, cnt3) = Convert666to555(cnt, cnt2, cnt3)
                convtab_666to565(cnt, cnt2, cnt3) = Convert666to565(cnt, cnt2, cnt3)
            Next
        Next
    Next
    Exit Function
End Function

Public Function UInt32ToSInt16(UInt32 As Long) As Integer
    Dim X%
    If UInt32 > 65535 Then
       MsgBox "You passed a value larger than 65535"
       Exit Function
    End If
    
    X% = UInt32 And &H7FFF
    UInt32ToSInt16 = X% Or -(UInt32 And &H8000)
End Function

Public Function SInt16ToUint32(SInt16 As Integer) As Long
    SInt16ToUint32 = SInt16 And &HFFFF&
End Function

Public Function AsciiZToString(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, Chr(0), vbBinaryCompare)
    If p > 0 Then
        res$ = Left(s, p)
    Else
        res$ = s
    End If
    AsciiZToString = res$
End Function

Public Function StringToAsciiZ(ByVal s As String, lMaxLen As Long) As String
    If Len(s) >= lMaxLen Then
        res$ = Left(s, lMaxLen)
    Else
        res$ = s & String(lMaxLen - Len(s), Chr(0))
    End If
    StringToAsciiZ = res$
End Function

Public Function addMapToFpg(fpg As cFpg, map As cMap) As Boolean
    Dim msgResult As VbMsgBoxResult
    Dim replace As Boolean
    
    'The code already exists
    If fpg.Exists(map.Code) Then
        msgResult = MsgBox("The FPG contains another map with the code: " & map.Code & _
                        ". Replace it?", vbQuestion + vbYesNoCancel)
        If msgResult = vbYes Then
            replace = True
        ElseIf msgResult = vbCancel Then
            addMapToFpg = False
            Exit Function
        Else
            map.Code = fpg.FreeCode
        End If
    End If
    
    'Add the map to the fpg
    If fpg.Add(map, replace) <> -1 Then
        MsgBox MSG_ERROR_ADDING_MAP & fpg.GetLastError, vbCritical
    Else
        addMapToFpg = True
    End If
End Function

'Creates a copy of a cMap object
Public Function getMapCopy(map As cMap) As cMap
    Dim Copy As cMap
    Dim data8() As Byte
    Dim data16() As Integer
    Dim cpData() As Integer
    
    Set Copy = New cMap
    
    If map.Depth = 8 Then
        ReDim data8(map.Width * map.Height - 1) As Byte
        map.GetData8 data8()
        Call Copy.CreateFromStream8(map.description, map.Code, map.Width, map.Height, data8, map.palette)
    ElseIf map.Depth = 16 Then
        ReDim data16(map.Width * map.Height - 1) As Integer
        map.GetData16 data16()
        Call Copy.CreateFromStream16(map.description, map.Code, map.Width, map.Height, data16)
    End If
    ReDim cpData(map.CPointsCount - 1) As Integer
    map.GetCPData cpData()
    Copy.SetCPData cpData
    
    Set getMapCopy = Copy
End Function

