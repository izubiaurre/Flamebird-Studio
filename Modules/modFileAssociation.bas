Attribute VB_Name = "modRegisterFileType"
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

'===========================================================================================================
'ASSOCIATE ICON WITH FILE - MODULE CODE
'===========================================================================================================

Option Explicit

'===========================================================================================================
'START VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const GWL_WNDPROC = -4

Private Const WM_CLOSE = &H10
Private Const WM_COPYDATA = &H4A

Private nCopyData As COPYDATASTRUCT
Private nBUFFER(1 To 255) As Byte
Private nOldProc As Long
'===========================================================================================================
'END VARIABLES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================

Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const SHCNF_IDLIST = &H0
Private Const SHCNE_ASSOCCHANGED = &H8000000

'ASSOCIATES A FILETYPE WITH YOUR PROGRAM. IT USES THE COMMAND, " %1", TO LOAD FILE IN YOUR PROGRAM
Public Sub RegisterType(ByVal sExt As String, ByVal sName As String, ByVal sConType As String, ByVal sDescription As String, ByVal Icon As String, Optional Executable As String)
    If Executable = "" Then
        Executable = App.Path & "\" & App.EXEName & ".exe %1"
    End If
    
    If Left(sExt, 1) <> "." Then sExt = "." & sExt

    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "", sName)
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "Content Type", sConType)
    Call CreateKey(HKEY_CLASSES_ROOT, sName, "", sDescription)
    'Call CreateKey(HKEY_CLASSES_ROOT, sName & "\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe," & iResIcon)
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\DefaultIcon", "", Icon)
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell", "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open", "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open\Command", "", Executable)
    
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

'UNASSOCIATES A FILETYPE WITH PROGRAM. IT SIMPLY DELETES THE FILENAME KEY.
Public Sub DeleteType(ByVal sExt As String, ByVal sName As String)
    If Left(sExt, 1) <> "." Then sExt = "." & sExt
    
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "", "")
    Call CreateKey(HKEY_CLASSES_ROOT, sExt, "Content Type", "")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\DefaultIcon")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open\Command")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell\Open")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName & "\Shell")
    Call DeleteKey(HKEY_CLASSES_ROOT, sName)
    
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

'CHECKS TO SEE IF YOUR PROGRAM IS THE HANDLER FOR A CERTAIN FILETYPE
Public Function FileAssociated(ByVal sExt As String, ByVal sName As String) As Boolean
    If Left(sExt, 1) <> "." Then sExt = "." & sExt
    
    FileAssociated = CBool(GetString(HKEY_CLASSES_ROOT, sExt, "") = sName)
End Function

'HELPER FUNCTION THAT CHECKS TO SEE IF A FILE IS ASSOCIATED WITH YOUR PROGRAM
Private Function GetString(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String)
    Dim lResult As Long
    Dim lHandle As Long
    Dim sBuffer As String
    Dim lLenBuffer As Long
    Dim lValueType As Long
    Dim iZeroPos As Integer
    
    Call RegOpenKey(hKey, sPath, lHandle)
    
    lResult = RegQueryValueEx(lHandle, sValue, 0&, lValueType, ByVal 0&, lLenBuffer)
    
    If lValueType = REG_SZ Then
        sBuffer = String(lLenBuffer, " ")
        lResult = RegQueryValueEx(lHandle, sValue, 0&, 0&, ByVal sBuffer, lLenBuffer)
        If lResult = ERROR_SUCCESS Then
            iZeroPos = InStr(sBuffer, Chr$(0))
            If iZeroPos > 0 Then
                GetString = Left$(sBuffer, iZeroPos - 1)
            Else
                GetString = sBuffer
            End If
        End If
    End If
End Function

'HELPER PROCEDURE THAT CREATES ALL STRING VALUES IN THE REGISTRY
Private Sub CreateKey(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String, ByVal sData As String)
    Dim lResult As Long
    
    Call RegCreateKey(hKey, sPath, lResult)
    Call RegSetValueEx(lResult, sValue, 0, REG_SZ, ByVal sData, Len(sData))
    Call RegCloseKey(lResult)
End Sub

'HELPER PROCEDIRE THAT DELETES A STRING IN A REGISTRY KEY
Private Sub DeleteKey(ByVal hKey As Long, ByVal sKey As String)
    Call RegDeleteKey(hKey, sKey)
End Sub

'===========================================================================================================
'START PROCEDURES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================
'SUBCLASSES THE FORM
Public Sub SubclassEnqueue(ByVal hwnd As Long)
    nOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf EnqueueProcedure)
End Sub

'UNSUBCLASSES THE FORM
Public Sub UnSubclassEnqueue(ByVal hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, nOldProc)
End Sub

'PROCESSES WHETHER OR NOT IS FIRST INSTANCE - IF IT ISNT, IT SENDS THE FILEPATH OF PATH TO FIRST INSTANCE
Public Sub EnqueueProcess(ByVal hwnd As Long, ByVal sCommand As String)
    If App.PrevInstance And sCommand <> "%1" And sCommand <> "" Then
        Dim lHwnd As Long
        
        lHwnd = CLng(Val(GetSetting(App.Title, "ActiveWindow", "Handle")))
        Call CopyMemory(nBUFFER(1), ByVal sCommand, Len(sCommand))

        With nCopyData
            .dwData = 3
            .cbData = Len(sCommand) + 1
            .lpData = VarPtr(nBUFFER(1))
        End With
        
        Call SendMessage(lHwnd, WM_COPYDATA, lHwnd, nCopyData)
        End
    Else
        SaveSetting App.Title, "ActiveWindow", "Handle", str(hwnd)
    End If
End Sub

'FIRES WHEN ANOTHER FILE IS OPENED WHEN AN INSTANCE IS ALREADY AVAILABLE
Private Function EnqueueProcedure(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_COPYDATA
            Dim sCommand As String
            
            Call CopyMemory(nCopyData, ByVal lParam, Len(nCopyData))
            Call CopyMemory(nBUFFER(1), ByVal nCopyData.lpData, nCopyData.cbData)
            sCommand = StrConv(nBUFFER, vbUnicode)

            'PROCESS FILE FROM ALL OTHER INSTANCES
            MsgBox sCommand
            
            EnqueueProcedure = 0
        Case WM_CLOSE
            Call UnSubclassEnqueue(hwnd)
        Case Else
            EnqueueProcedure = CallWindowProc(nOldProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function
'===========================================================================================================
'END PROCEDURES TO ENQUEUE FILES FROM EXPLORER
'===========================================================================================================

'===========================================================================================================
'HOW TO USE THIS CODE!
'===========================================================================================================
'
'Private Sub Form_Load()
'    Call EnqueueProcess(hWnd, Command$)
'
'    If Not App.PrevInstance And Command$ <> "%1" And Command$ <> "" Then
'        'ONLY USED IF THIS WAS THE FIRST INSTANCE - ALL OTHER INSTANCES GO THROUGH SUBCLASSING
'        MsgBox Command$
'    End If
'
'    Call SubclassEnqueue(hWnd)
'End Sub
'
'Private Sub cmdRegister_Click()
'    Call RegisterType(".mp3", "XamP.File", "Audio/MPEG", "XamP Media File", 0)
'End Sub
'
'Private Sub cmdUnRegister_Click()
'    Call DeleteType(".mp3", "XamP.File")
'End Sub
'
'Private Sub cmdCheckAssociation_Click()
'    MsgBox FileAssociated(".mp3", "XamP.File"), vbInformation, "Associated?"
'End Sub
'
'===========================================================================================================
'CODE MODIFIED BY MICHAEL DOMBROWSKI
'===========================================================================================================
