VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIP"
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

' Initialization class
' Seperated off from cBitmap class 02/14/2004

Event Error(ByVal lGdiError As Long, ByVal sErrorDesc As String, ByVal sClassName As String)

Private lToken As Long

Private Sub Class_Initialize()
  Dim gsi As GdiplusStartupInput
  
  gsi.GdiplusVersion = 1
  
  ' next line used when debugging in VC, a la DebugBreak API
  ' gsi.DebugEventCallback = PassAddress(AddressOf DebugEventProc)
  If GdiplusStartup(lToken, gsi, ByVal 0) <> status.Ok Then
    Debug.Print "GdiplusStartup Error!"
  End If
End Sub

Public Function Error(ByVal lGdiError As Long, ByVal sClassName As String)
  RaiseEvent Error(lGdiError, GdiErrorString(lGdiError), sClassName)
End Function

Private Function GdiErrorString(ByVal lError As status) As String
  Dim s As String
  
  Select Case lError
    Case GenericError:              s = "Generic Error"
    Case InvalidParameter:          s = "Invalid Parameter"
    Case OutOfMemory:               s = "Out Of Memory"
    Case ObjectBusy:                s = "Object Busy"
    Case InsufficientBuffer:        s = "Insufficient Buffer"
    Case NotImplemented:            s = "Not Implemented"
    Case Win32Error:                s = "Win32 Error"
    Case WrongState:                s = "Wrong State"
    Case Aborted:                   s = "Aborted"
    Case FileNotFound:              s = "File Not Found"
    Case ValueOverflow:             s = "Value Overflow"
    Case AccessDenied:              s = "Access Denied"
    Case UnknownImageFormat:        s = "Unknown Image Format"
    Case FontFamilyNotFound:        s = "FontFamily Not Found"
    Case FontStyleNotFound:         s = "FontStyle Not Found"
    Case NotTrueTypeFont:           s = "Not TrueType Font"
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version"
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized"
    Case PropertyNotFound:          s = "Property Not Found"
    Case PropertyNotSupported:      s = "Property Not Supported"
    Case Else:                      s = "Unknown GDI+ Error"
  End Select
  
  GdiErrorString = s
End Function

Private Sub Class_Terminate()
  If lToken Then GdiplusShutdown lToken
End Sub
