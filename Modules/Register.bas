Attribute VB_Name = "ModRegister"
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

'**************************************
' Name: ActiveX Dll Register/UnRegister
' Description: This code shows how to register and unregister ActiveX dlls programatically,
'       without relying on regsvr32 f or the task. It's built into a reuseable class that
'       can be put in your own code or compiled into a dll. Based very loosely on code
'       from Vasudevan S.
' By: Robert J May
'
' Inputs:The file name
'
' Returns:7 flags. See the enumeration and sample code for details.
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46775&lngWId=1'for details.
'**************************************

Option Explicit


Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long


Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long


Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long


Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
    Private Const STATUS_WAIT_0 = &H0
    Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
    Private Const NOERRORS As Long = 0


Private Enum stRegisterStatus
    stFileCouldNotBeLoadedIntoMemorySpace = 1
    stNotAValidActiveXComponent = 2
    stActiveXComponentRegistrationFailed = 3
    stActiveXComponentRegistrationSuccessful = 4
    stActiveXComponentUnRegisterSuccessful = 5
    stActiveXComponentUnRegistrationFailed = 6
    stNoFileProvided = 7
End Enum


Public Function Register(ByVal p_sFileName As String) As Variant
    Dim lLib As Long
    Dim lProcAddress As Long
    Dim lThreadID As Long
    Dim lSuccess As Long
    Dim lExitCode As Long
    Dim lThreadHandle As Long
    Dim lRet As Long
    On Error GoTo errorhandler


    If lRet = NOERRORS Then


        If p_sFileName = "" Then
            lRet = stNoFileProvided
        End If
    End If


    If lRet = NOERRORS Then
        lLib = LoadLibraryRegister(p_sFileName)


        If lLib = 0 Then
            lRet = stFileCouldNotBeLoadedIntoMemorySpace
        End If
    End If


    If lRet = NOERRORS Then
        lProcAddress = GetProcAddressRegister(lLib, "DllRegisterServer")


        If lProcAddress = 0 Then
            lRet = stNotAValidActiveXComponent
        Else
            lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)


            If lThreadHandle <> 0 Then
                lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)


                If lSuccess = 0 Then
                    Call GetExitCodeThread(lThreadHandle, lExitCode)
                    Call ExitThread(lExitCode)
                    lRet = stActiveXComponentRegistrationFailed
                Else
                    lRet = stActiveXComponentRegistrationSuccessful
                End If
            End If
        End If
    End If
ExitRoutine:
    Register = lRet


    If lThreadHandle <> 0 Then
        Call CloseHandle(lThreadHandle)
    End If


    If lLib <> 0 Then
        Call FreeLibraryRegister(lLib)
    End If
    Exit Function
errorhandler:
    lRet = Err.Number
    GoTo ExitRoutine
End Function


Public Function UnRegister(ByVal p_sFileName As String) As Variant
    Dim lLib As Long
    Dim lProcAddress As Long
    Dim lThreadID As Long
    Dim lSuccess As Long
    Dim lExitCode As Long
    Dim lThreadHandle As Long
    Dim lRet As Long
    On Error GoTo errorhandler


    If lRet = NOERRORS Then


        If p_sFileName = "" Then
            lRet = stNoFileProvided
        End If
    End If


    If lRet = NOERRORS Then
        lLib = LoadLibraryRegister(p_sFileName)


        If lLib = 0 Then
            lRet = stFileCouldNotBeLoadedIntoMemorySpace
        End If
    End If


    If lRet = NOERRORS Then
        lProcAddress = GetProcAddressRegister(lLib, "DllUnregisterServer")


        If lProcAddress = 0 Then
            lRet = stNotAValidActiveXComponent
        Else
            lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)


            If lThreadHandle <> 0 Then
                lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)


                If lSuccess = 0 Then
                    Call GetExitCodeThread(lThreadHandle, lExitCode)
                    Call ExitThread(lExitCode)
                    lRet = stActiveXComponentUnRegistrationFailed
                Else
                    lRet = stActiveXComponentUnRegisterSuccessful
                End If
            End If
        End If
    End If
ExitRoutine:
    UnRegister = lRet


    If lThreadHandle <> 0 Then
        Call CloseHandle(lThreadHandle)
    End If


    If lLib <> 0 Then
        Call FreeLibraryRegister(lLib)
    End If
    Exit Function
errorhandler:
    lRet = Err.Number
    GoTo ExitRoutine
End Function

