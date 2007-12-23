Attribute VB_Name = "modExecCMD"
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

Public Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)

Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Public Const SW_SHOWMINNOACTIVE = 7
Public Const STARTF_USESHOWWINDOW = &H1
Public Const INFINITE = -1&
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Private Const STATUS_WAIT_0 As Long = &H0
Private Const STATUS_USER_APC As Long = &HC0
Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_IO_COMPLETION As Long = STATUS_USER_APC
Private Const WAIT_ABANDONED_0 As Long = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_FAILED As Long = &HFFFFFFFF


Public Function ExecCmdPipe(ByVal CmdLine As String) As String

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    
    DoEvents
    'creamos un buffer (string)
    mybuff = String(30 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 1048576)
    DoEvents
    If ret = 0 Then
        '===Error
        ExecCmdPipe = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    
    GetStartupInfo start 'Obtenemos los valores de inicializacion segun el SO
    
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SW_HIDE
    DoEvents
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    DoEvents
    If ret <> 1 Then
        ExecCmdPipe = "Error: CreateProcess failed. " & Err.LastDllError
        Exit Function
    End If
    DoEvents
    'EXBUG: aca se clava cuando no alcansa el buffer del pipe
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    
    If ret = WAIT_OBJECT_0 Then
        
        DoEvents
        'aca lee y lo devuelve en mybuff
        bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
        DoEvents
        If bSuccess = 1 Then ' si se pudo leer
            'y lo pasa a returnstr
            sReturnStr = Left(mybuff, bytesread)
        Else
            ExecCmdPipe = "Error: ReadFile failed. " & Err.LastDllError
            Exit Function
        End If
    End If
    DoEvents
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    DoEvents
    
    ExecCmdPipe = sReturnStr
End Function
