Attribute VB_Name = "modExecution"
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

'MSG Constants (for future multi-language support)
Private Const MSG_COMPILE_NOFENIXDIR = "Compiler directory has not been configured or does not exist"
Private Const MSG_COMPILE_FILENOTFOUND = "The file you are trying to compile does not exist"
Private Const MSG_COMPILE_NOTALREADYSAVED = "The file has not been saved yet. Save the file before compile"
Private Const MSG_COMPILE_NOSTDOUT = "Flamebird is waintg for the compiler to finish compiling for getting its output but " _
        & "nothing seems to happen. Do you want to wait a bit more?"
Private Const MSG_COMPILE_NODCB = "Flamebird is waintg for the compiler to create the dcb file but " _
        & "nothing seems to happen. Do you want to wait a bit more?"
Private Const MSG_RUN_DBCNOTFOUND = "DCB file not found. You must compile first"

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal Hwnd As Long, ByVal nIDEvent As Long) As Long
    
Private objDOS As New DOSOutputs
Private mTimerId As Long
Private mFxiDir As String, mFileDir As String

Public Sub SaveBeforeCompiling()
    Dim i As Integer
    Dim f As Form
    Dim ff As IFileForm
    
    'Save files if user has selected it
    Select Case R_SaveBeforeCompiling
    Case 1 'Save current
        Set f = frmMain.ActiveFileForm
        If Not f Is Nothing Then
            Set ff = f
            If ff.Identify = FF_SOURCE Then
                SaveFileOfFileForm ff
            End If
        End If
    Case 2 'Save project
        If Not openedProject Is Nothing Then
            For Each f In Forms
                If TypeOf f Is IFileForm Then
                    Set ff = f
                    If ff.IsDirty And openedProject.FileExist(ff.FilePath) = True Then
                        SaveFileOfFileForm ff
                    End If
                End If
            Next
        End If
    Case 3 'Save all
        For Each f In Forms
            If TypeOf f Is IFileForm Then
                Set ff = f
                If ff.IsDirty Then
                    SaveFileOfFileForm ff
                End If
            End If
        Next
    End Select
End Sub
    
    
Private Function isFxiRunning() As Boolean
    Dim sTitulo As String
    Dim Hwnd As Long
    Dim found As Boolean
    Dim procId As Long
    
    'Get the first window
    Hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    
    found = False
    'Search in the rest of the windows
    Do While Hwnd <> 0&
        If GetWindowThreadProcessId(Hwnd, procId) = objDOS.ThreadID Then
            found = True
            Exit Do
        End If
        'Get the next window
        Hwnd = GetWindow(Hwnd, GW_HWNDNEXT)
    Loop
    
    isFxiRunning = found
End Function

Private Sub ReadFxiOutputAndErrors()
    Dim stdoutFile As String, stdout As String
    Dim stderrFile As String, stderr As String
    Dim f As file
    Dim fxcDir As String
    Dim textStream As textStream
    
    
    On Error GoTo errors:
    'Look for the output files in the Fenix dir and the file dir
    If FSO.FileExists(mFxiDir & "\stddebug.txt") Then
        stdoutFile = mFxiDir & "\stddebug.txt"
    ElseIf FSO.FileExists(mFileDir & "\stddebug.txt") Then
        stdoutFile = mFileDir & "\stddebug.txt"
    End If
    
    If FSO.FileExists(mFxiDir & "\stderr.txt") Then
        stderrFile = mFxiDir & "\stderr.txt"
    ElseIf FSO.FileExists(mFileDir & "\stderr.txt") Then
        stderrFile = mFileDir & "\stderr.txt"
    End If
    
    'Read stdout
    If stdoutFile <> "" Then
        Set f = FSO.GetFile(stdoutFile)
        Set textStream = f.OpenAsTextStream(ForReading, TristateFalse)
        stdout = textStream.ReadAll()
        textStream.Close
    End If
    'Read stderr
    If stderrFile <> "" Then
        Set f = FSO.GetFile(stderrFile)
        Set textStream = f.OpenAsTextStream(ForReading, TristateFalse)
        stderr = textStream.ReadAll()
        textStream.Close
    End If
    
    'Set the frmDebug and frmErrors text...
    frmDebug.txtOutput.text = stdout
    frmErrors.txtOutput.text = stderr
    
    frmMain.StatusBar.RedrawPanel ("FXI_OUTPUT_INFO")
    Exit Sub
errors:
    MsgBox "Execution terminated abnormally"
    Exit Sub
End Sub

Private Sub TimerProc(ByVal Hwnd As Long, ByVal nIDEvent As Long, _
              ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    'Check if the fxi is not running and
    If isFxiRunning() = False Then
        KillTimer 0&, mTimerId 'End timer
        ReadFxiOutputAndErrors
    End If
End Sub


'-------------------------------------------------------------------------------------
'FUNCTION: SearchErrorLine()
'DESCRIPTION: Searches the string ":error:" in the CString.
'             which is supposed to be the output of the fxc.
'             The name of the file where the error takes place is stored in sFileError
'             and the error message is placed in the sError string.
'RETURNS: -1 if no error was found, otherwise the error line number.
'-------------------------------------------------------------------------------------
Private Function SearchErrorLine(ByVal Cstring As String, ByRef sFileError As String, ByRef sError As String) As Long
    Dim sLine As String, i As Integer
    Dim lines() As String
    Dim numLine As Integer
    Dim searchStart As Long, searchEnd As Long
    Dim result As Long
    
    result = -1
    lines = Split(Cstring, vbNewLine, , vbTextCompare) 'Splits the string into separated lines
    'Read line by line
    For i = 0 To UBound(lines)
        sLine = Trim(lines(i))
'        Old Mode
'        If InStr(1, sLine, "Error in file", vbTextCompare) > 0 Then
'            'Get the file where the error occurrs
'            sFileError = Trim(Mid(sLine, 15, InStr(sLine, " at line ") - Len(" at line ") - 5))
'            'Get Error line number
'            searchStart = InStr(sLine, "at line") '+ 7
'            searchEnd = InStr(Mid(sLine, searchStart), ":") - 1
'            numLine = CInt(Trim(Mid(sLine, searchStart, searchEnd)))
'            result = numLine
'            'Get the error message
'            sError = Trim(Mid(sLine, InStrRev(sLine, ":") + 1))
'            Exit For
'        End If
        

        If InStr(1, sLine, ": error:", vbTextCompare) > 0 Then
             'Get the file where the error occurrs
             sFileError = Trim(Left(sLine, InStr(3, sLine, ":", vbTextCompare) - 1))
                'MsgBox sFileError
             'Get Error line number
             searchStart = InStr(3, sLine, ":") + 1
                'MsgBox searchStart
             searchEnd = InStr(sLine, ": error:")
                'MsgBox searchEnd
             numLine = CInt(Trim(Mid(sLine, searchStart, searchEnd - searchStart)))
                'MsgBox numLine
             result = numLine
             'Get the error message
             sError = Trim(Right(sLine, Len(sLine) - (InStr(sLine, ": error") + 8)))
             Exit For
        End If
    Next
    SearchErrorLine = result
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: CompileFFSource()
'DESCRIPTION: Call compile function for a fileform
'RETURNS: true if compilation success, otherwise false.
'-------------------------------------------------------------------------------------
Public Function CompileFFSource(ff As IFileForm)
    Dim bResult As Boolean
    
    bResult = False
    If ff.Identify = FF_SOURCE Then
        If ff.AlreadySaved = True Then 'Saved at least once
            bResult = Compile(ff.FilePath)
        Else
            MsgBox MSG_COMPILE_NOTALREADYSAVED, vbExclamation
        End If
    End If
    
    CompileFFSource = bResult
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: CompileFFSource()
'DESCRIPTION: Call run function for a fileform
'RETURNS: true if compilation success, otherwise false.
'-------------------------------------------------------------------------------------
Public Function RunFFSource(ff As IFileForm)
    Dim bResult As Boolean
    
    bResult = False
    If ff.Identify = FF_SOURCE Then
        bResult = Run(ff.FilePath)
    End If
    RunFFSource = bResult
End Function
'-------------------------------------------------------------------------------------
'FUNCTION: Compile()
'DESCRIPTION: Executes FXC for the specified sFile
'RETURNS: true if no error, otherwise false.
'-------------------------------------------------------------------------------------
Public Function Compile(ByVal sFile As String) As Boolean
    Dim bResult As Boolean
    Dim dcbFile As String, sCommand As String, stdoutFile As String
    Dim fxcDir As String
    Dim sFileError As String
    Dim errorLine As Long
    Dim sError As String
    Dim stdout As String
    Dim bCancel As Boolean
    Dim timeStart As Long
    Dim msgResult As VbMsgBoxResult
    Dim textStream As textStream
    Dim f As file
    Dim fDoc As frmDoc
    
    On Error GoTo errhandler
    
    bResult = False
    
    'Determine which compilation Fenix Directory to use
    fxcDir = fenixDir
    If Not openedProject Is Nothing Then
        If openedProject.FileExist(sFile) And openedProject.useOtherFenix = True Then
            fxcDir = openedProject.fenixDir
        End If
    End If
    
    If FSO.FolderExists(fxcDir) Then
        If FSO.FileExists(sFile) Then
            ChDir FSO.GetParentFolderName(sFile)
            'Verify if the DCB already exists and delete it if so
            dcbFile = FSO.GetParentFolderName(sFile) & "\" & FSO.GetBaseName(sFile) & ".dcb"
            If FSO.FileExists(dcbFile) Then
                FSO.DeleteFile dcbFile, True
            End If
            
            'Delete stdout file if exists
            stdoutFile = fxcDir & "\stdout.txt"
            If FSO.FileExists(fxcDir & "\stdout.txt") = True Then
                FSO.DeleteFile fxcDir & "\stdout.txt"
            End If
            
            'Execute Compiler
            If R_Compiler = 0 Then
                sCommand = Chr(34) & fxcDir & "\fxc.exe" & Chr(34) & " "
            Else
                sCommand = Chr(34) & fxcDir & "\bgdc.exe" & Chr(34) & " "
            End If
            If R_Debug Then 'Debug mode on (compile using parameters)
                sCommand = sCommand + " -g"
            End If
            If R_MsDos Then
                sCommand = sCommand + " -c"
            End If
            If R_Stub Then
                'sCommand = sCommand + " -s " & fxcDir & "\bgdi.exe"
                sCommand = sCommand + " -s " & "bgdi.exe"
            End If
            If R_AutoDeclare Then
                sCommand = sCommand + " -Ca"
            End If
            sCommand = sCommand & " " & Chr(34) & sFile & Chr(34) '& " > " & Chr(34) & stdoutFile & Chr(34)
                   'MsgBox sCommand
                   'Clipboard.Clear
                   'Clipboard.SetText sCommand
            stdout = objDOS.ExecuteCommand(sCommand)
            
            'MsgBox stdout
                        
            'The output can be empty if the fxc does not produce an stdout stream
            'in that case, the output should be in the stdout.txt in the fenix folder
            'so we just wait for this file to be created
            If (stdout = "") Then
                'MsgBox "a"
                bCancel = False
                If R_Stub Then
                    'Wait for the stdout file to be created
                    timeStart = GetTickCount()
                    While FSO.FileExists(stdoutFile) = False And bCancel = False
                        DoEvents
                        If GetTickCount() - timeStart > 4000 Then
                            msgResult = MsgBox(MSG_COMPILE_NOSTDOUT, vbYesNo + vbQuestion)
                            If msgResult = vbNo Then
                                bCancel = True
                            Else
                                timeStart = GetTickCount()
                            End If
                        End If
                    Wend
                End If
                
                If bCancel = False Then
                    'Wait for the stdout file to be filled
                    bCancel = False
                    timeStart = GetTickCount()
                    While FSO.GetFile(stdoutFile).Size = 0 And bCancel = False
                        DoEvents
                        If GetTickCount() - timeStart > 4000 Then
                            msgResult = MsgBox(MSG_COMPILE_NOSTDOUT, vbYesNo + vbQuestion)
                            If msgResult = vbNo Then
                                bCancel = True
                            Else
                                timeStart = GetTickCount()
                            End If
                        End If
                    Wend
                End If
                
                If bCancel = True Then
                    MsgBox "Compilation aborted by user!", vbInformation
                    Compile = False
                    Exit Function
                End If
                
                Set f = FSO.GetFile(stdoutFile)
                Set textStream = f.OpenAsTextStream(ForReading, TristateFalse)
                If Not textStream.AtEndOfStream Then
                    stdout = textStream.ReadAll
                End If
                'I think the stdout file is encoded as UTF without BOM and I cannot manage to read
                'its chars correctly directly, so I make the character replacement myself...
                stdout = replace(replace(replace(stdout, "Ã©", "é"), "Ã¡", "á"), "Ã¼", "ü")
                textStream.Close
            End If
            
            stdout = Right(stdout, Len(stdout) - 267)
            
            frmOutput.txtOutput.text = stdout
            'Search for the error line
            errorLine = SearchErrorLine(stdout, sFileError, sError)
            If errorLine = -1 Then 'No error line in debuger
                bCancel = False
                If R_Stub Then
                    dcbFile = Left(dcbFile, Len(dcbFile) - 3) & "exe"
                End If
                'Wait for the dcb to be created
                timeStart = GetTickCount()
                While FSO.FileExists(dcbFile) = False And bCancel = False
                    DoEvents
                    If GetTickCount() - timeStart > 4000 Then
                        msgResult = MsgBox(MSG_COMPILE_NODCB, vbYesNo + vbQuestion)
                        If msgResult = vbNo Then
                            bCancel = True
                        Else
                            timeStart = GetTickCount()
                        End If
                    End If
                Wend
                
                If bCancel = True Then
                    MsgBox "Compilation aborted by user!", vbInformation
                    Compile = False
                    Exit Function
                End If
                
                bResult = True 'Compilation succesfull
            Else 'Show a message and go to the line where the error ocurred
                MsgBox "Error compiling " & sFileError & " at line " & errorLine & vbCrLf & vbCrLf & sError
                sFileError = replace(sFileError, "/", "\")
                Set fDoc = FindFileForm(sFileError)
                If fDoc Is Nothing Then 'The file is not opened
                    Set fDoc = NewFileForm(FF_SOURCE, sFileError)
                End If
                If Not fDoc Is Nothing Then
                    Set fDoc = frmMain.ActiveForm
                    errorLine = IIf(errorLine = 0, 0, errorLine - 1)
                    fDoc.cs.ExecuteCmd cmCmdGoToLine, CInt(errorLine)
                    fDoc.cs.HighlightedLine = CInt(errorLine)
                End If
            End If
        Else
            MsgBox MSG_COMPILE_FILENOTFOUND, vbCritical
        End If
    Else 'Fenix Dir Does not exist
        MsgBox MSG_COMPILE_NOFENIXDIR, vbExclamation
    End If
    
    Compile = bResult
    
    Exit Function
errhandler:
    If Err.Number > 0 Then ShowError ("modExecution.Compile()")
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: Run()
'DESCRIPTION: Executes FXI for the dcb of the specified sFile
'RETURNS: true if no error, otherwise false.
'-------------------------------------------------------------------------------------
Public Function Run(ByVal sFile As String) As Boolean
    Dim dcbFile As String, sCommand As String
    Dim bCancel As Boolean
    Dim bResult As Boolean
    Dim timeStart As Long
    Dim msgResult As VbMsgBoxResult
    Dim fxiDir As String
    Dim stdout As String
    Dim stdoutFile As String
    Dim textStream As textStream
    Dim f As file

    'Determine which compilation Fenix Directory to use
    fxiDir = fenixDir
    If Not openedProject Is Nothing Then
        If openedProject.FileExist(sFile) And openedProject.useOtherFenix = True Then
            fxiDir = openedProject.fenixDir
        End If
    End If

    mFxiDir = fxiDir
    
    bResult = False
    If FSO.FolderExists(fxiDir) Then 'Valid Interpreter Dir
        'Look for the DCB
        dcbFile = FSO.GetParentFolderName(sFile) & "\" & FSO.GetBaseName(sFile) & ".dcb"
        mFileDir = FSO.GetParentFolderName(sFile)
        If FSO.FileExists(dcbFile) Then
            If R_Compiler = 0 Then  ' Fenix
                sCommand = Chr(34) & fxiDir & "\fxi.exe " & Chr(34) & IIf(R_Debug, " -d ", " ") & Chr(34) & dcbFile & Chr(34)
            Else                    ' Bennu
                sCommand = Chr(34) & fxiDir & "\bgdi.exe " & Chr(34) & IIf(R_Debug, " -d ", " ") & Chr(34) & dcbFile & Chr(34)
            End If
            'Params
            'If R_filter Then sCommand = sCommand & " -f "
            'If R_DoubleBuf Then sCommand = sCommand & " -b "
            
            'Delete stdout file if exists
            stdoutFile = fxiDir & "\stddebug.txt"
            FSO.CreateTextFile stdoutFile
'            If FSO.FileExists(stdoutFile) = True Then
'                FSO.DeleteFile stdoutFile
'            Else
'                FSO.CreateTextFile fxiDir & "\stddebug.txt"
'            End If
            
            sCommand = sCommand & " > " & Chr(34) & stdoutFile & Chr(34)
            stdout = objDOS.ExecuteCommand(sCommand)
            bResult = True 'Execution succesful
            
            frmDebug.txtOutput.text = ""
            frmDebug.txtOutput.text = stdout
            
            'Start the end-running check timer
            mTimerId = SetTimer(0&, 0&, 300, AddressOf TimerProc)
        Else
            If R_Stub Then  ' execute .exe file cause doesn't exist dcb file
                sCommand = Left(dcbFile, Len(dcbFile) - 3) & "exe"
                objDOS.ExecuteCommand sCommand
                bResult = True 'Execution succesful
            
                'Start the end-running check timer
                mTimerId = SetTimer(0&, 0&, 300, AddressOf TimerProc)
            Else            'DCB does not exists
                MsgBox MSG_RUN_DBCNOTFOUND, vbExclamation
            End If
        End If
    Else 'Invalid Fenix Dir
        MsgBox MSG_COMPILE_NOFENIXDIR, vbExclamation
    End If
    Run = bResult
End Function
