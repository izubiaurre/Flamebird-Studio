Attribute VB_Name = "modFunctions"
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

'Used in SetMinWindowSize
Public OldWindowProc As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long

Public Const GWL_WNDPROC = (-4)
Public MinWidth As Integer
Public MinHeight As Integer
Type WINDOWPOS
  hwnd As Long
  hWndInsertAfter As Long
  X As Long
  Y As Long
  cx As Long
  cy As Long
  flags As Long
End Type
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

'FormDrag Stuff
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'WebPage Launcher
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

'Discover Directories
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetFolderPath Lib "ShFolder" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal CSIDL As Long, ByVal TOKENHANDLE As Long, ByVal flags As Long, ByVal lpPath As String) As Long
Public Const CSIDL_DESKTOP = &H0                 '{desktop}
Public Const CSIDL_INTERNET = &H1                'Internet Explorer (icon on desktop)
Public Const CSIDL_PROGRAMS = &H2                'Start Menu\Programs
Public Const CSIDL_CONTROLS = &H3                'My Computer\Control Panel
Public Const CSIDL_PRINTERS = &H4                'My Computer\Printers
Public Const CSIDL_PERSONAL = &H5                'My Documents
Public Const CSIDL_FAVORITES = &H6               '{user}\Favourites
Public Const CSIDL_STARTUP = &H7                 'Start Menu\Programs\Startup
Public Const CSIDL_RECENT = &H8                  '{user}\Recent
Public Const CSIDL_SENDTO = &H9                  '{user}\SendTo
Public Const CSIDL_BITBUCKET = &HA               '{desktop}\Recycle Bin
Public Const CSIDL_STARTMENU = &HB               '{user}\Start Menu
Public Const CSIDL_DESKTOPDIRECTORY = &H10       '{user}\Desktop
Public Const CSIDL_DRIVES = &H11                 'My Computer
Public Const CSIDL_NETWORK = &H12                'Network Neighbourhood
Public Const CSIDL_NETHOOD = &H13                '{user}\nethood
Public Const CSIDL_FONTS = &H14                  'windows\fonts
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16       'All Users\Start Menu
Public Const CSIDL_COMMON_PROGRAMS = &H17        'All Users\Programs
Public Const CSIDL_COMMON_STARTUP = &H18         'All Users\Startup
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
Public Const CSIDL_APPDATA = &H1A                '{user}\Application Data
Public Const CSIDL_PRINTHOOD = &H1B              '{user}\PrintHood
Public Const CSIDL_LOCAL_APPDATA = &H1C          '{user}\Local Settings \ Application Data (non roaming)
Public Const CSIDL_ALTSTARTUP = &H1D             'non localized startup
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E      'non localized common startup
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23          'All Users\Application Data
Public Const CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()
Public Const CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
Public Const CSIDL_PROGRAM_FILES = &H26           'C:\Program Files
Public Const CSIDL_MYPICTURES = &H27              'C:\Program Files\My Pictures
Public Const CSIDL_PROFILE = &H28                 'USERPROFILE
Public Const CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
Public Const CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B    'C:\Program Files\Common
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
Public Const CSIDL_COMMON_TEMPLATES = &H2D        'All Users\Templates
Public Const CSIDL_COMMON_DOCUMENTS = &H2E        'All Users\Documents
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F       'All Users\Start Menu\Programs

Private Type SHORTITEMID
    cb As Long
    abID As Integer
End Type
Private Type ITEMIDLIST
    mkid As SHORTITEMID
End Type

'Used with Pause()
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetInputState Lib "user32" () As Long

'Used with PlaySound()
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Sub PlaySound(filename As String)
    Dim wFlags%, X, SoundName As String
    
    SoundName$ = filename
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Function LoadWebPage(ByVal vPage As String, f As Form)
  
  On Error Resume Next
  vPage = Trim(vPage)
  ShellExecute f.hwnd, "open", vPage, vbNullString, vbNullString, SW_SHOW
  On Error GoTo 0
  
End Function

Public Sub FormDrag(TheForm As Form)
    
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
    
End Sub

Public Function fGetSpecialFolder(ByVal lngCSIDL As Long, f As Form) As String
    Dim udtIDL As ITEMIDLIST
    Dim lngRtn As Long
    Dim strFolder As String
    Dim Path As String * 260
    
    lngRtn = SHGetSpecialFolderLocation(f.hwnd, lngCSIDL, udtIDL)
    If lngRtn = 0 Then
        strFolder = Space$(260)
        lngRtn = SHGetPathFromIDList( _
        ByVal udtIDL.mkid.cb, ByVal strFolder)
        If lngRtn Then
            fGetSpecialFolder = Left$(strFolder, _
            InStr(1, strFolder, Chr$(0)) - 1) & "\"
        End If
    Else
        lngRtn = SHGetFolderPath(f.hwnd, lngCSIDL, 0, 0, Path)
        If lngRtn = 0 Then
            strFolder = Space$(260)
            lngRtn = SHGetPathFromIDList( _
            ByVal udtIDL.mkid.cb, ByVal strFolder)
            If lngRtn Then
                fGetSpecialFolder = Left$(strFolder, _
                InStr(1, strFolder, Chr$(0)) - 1) & "\"
            End If
        End If
    End If

End Function

Public Sub Pause(numSeconds As Single) 'Pauses for numSeconds Seconds (Decimals are OK)
    Dim t As Single
    Dim T2 As Single
    Dim num As Single
    
    num = numSeconds * 1000
    t = GetTickCount()
    T2 = GetTickCount()
    Do Until T2 - t >= num
        If GetInputState <> 0 Then DoEvents
        T2 = GetTickCount()
    Loop
End Sub

Public Function CountChar(text As String, Find As String) As Integer 'Counts occurances of a single character is a string
    Dim X As Integer
    For X = 1 To Len(text)
        
        If Mid(text, X, 1) = Find Then
            CountChar = CountChar + 1
        End If
        
    Next
End Function

Public Function KillFile(ByVal vFilename As String) 'Delete a Given File
  
  On Error Resume Next
  Kill vFilename
  On Error GoTo 0

End Function

Public Function Rand(min As Integer, max As Integer) As Integer 'Returns a random integer between the Min and Max Values
TryAgain:
    Call Randomize(Timer)
    Rand = Int((Rnd * max) + min)
    
    'This shouldn't be needed, but I like to play it safe.
    If Rand > max Or Rand < min Then GoTo TryAgain
End Function

Public Function RandomColor() As Long 'Returns a random color
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    Call Randomize(Timer)
    Red = Rand(1, 255)
    Green = Rand(1, 255)
    Blue = Rand(1, 255)
    RandomColor = RGB(Red, Green, Blue)
End Function

Public Function getAppPath() As String 'Returns Application Path
    getAppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function

Public Function GetWinDir() As String 'Returns Windows Folder
    Dim nSize As Long
    Dim tmp As String

    tmp = Space$(256)
    nSize = Len(tmp)
    Call GetWindowsDirectory(tmp, nSize)
    GetWinDir = TrimNull(tmp) & "\"
End Function

Public Function GetTempDir() As String 'Returns Windows Temp Folder
    Dim nSize As Long
    Dim tmp As String

    tmp = Space$(256)
    nSize = Len(tmp)
    Call GetTempDirectory(tmp, nSize)
    GetTempDir = TrimNull(tmp) & "\"
End Function

Public Function GetSysDir() As String 'Returns Windows System Folder
    Dim nSize As Long
    Dim tmp As String
    
    tmp = Space$(256)
    nSize = Len(tmp)
    Call GetSystemDirectory(tmp, nSize)
    GetSysDir = TrimNull(tmp) & "\"
End Function

Private Function TrimNull(item As String) 'Used in get___Dir() Functions
    Dim Pos As Integer
    Pos = InStr(item, Chr$(0))

    If Pos Then
        TrimNull = Left$(item, Pos - 1)
    Else
        TrimNull = item
    End If
End Function

Function InDesignMode() As Boolean 'Returns running state of application
    'True = Running in VB IDE
    'False = Running as compiled EXE
    On Error GoTo Err
    Debug.Print 1 / 0
    InDesignMode = False
    Exit Function

Err:
    InDesignMode = True
End Function

Public Function SetMinWindowSize(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
  
  ' Keep the dimensions in bounds.
  If msg = WM_WINDOWPOSCHANGING Then
    If lParam.cx < MinWidth Then lParam.cx = MinWidth
    If lParam.cy < MinHeight Then lParam.cy = MinHeight
  End If
  
  ' Continue normal processing. VERY IMPORTANT!
  SetMinWindowSize = CallWindowProc(OldWindowProc, hwnd, msg, wParam, lParam)
  
End Function

Public Function SimpleEncrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    
    On Error GoTo errorhandler
    
    Dim i As Integer
    Dim Char As String
    SimpleEncrypt = ""
    
    If StringToEncrypt = "" Then Exit Function

    For i = 1 To Len(StringToEncrypt)
        
        Char = Asc(Mid(StringToEncrypt, i, 1))
        SimpleEncrypt = SimpleEncrypt & Len(Char) & Char
    
    Next i
    


    If AlphaEncoding Then
        
        StringToEncrypt = SimpleEncrypt
        SimpleEncrypt = ""


        For i = 1 To Len(StringToEncrypt)
            
            SimpleEncrypt = SimpleEncrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        
        Next i
    
    End If
    
    Exit Function

errorhandler:
    
    SimpleEncrypt = ""

End Function

Public Function SimpleDecrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    
    Dim i As Integer
    
    On Error GoTo errorhandler
    
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    
    If StringToDecrypt = "" Then Exit Function

    If AlphaDecoding Then
        
        SimpleDecrypt = StringToDecrypt
        StringToDecrypt = ""


        For i = 1 To Len(SimpleDecrypt)
            
            StringToDecrypt = StringToDecrypt & (Asc(Mid(SimpleDecrypt, i, 1)) - 147)
        
        Next i
    
    End If
    
    SimpleDecrypt = ""

    Do Until StringToDecrypt = ""
        
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        SimpleDecrypt = SimpleDecrypt & Chr(CharCode)
    
    Loop
    
    Exit Function
    
errorhandler:
    
    SimpleDecrypt = ""

End Function

Function CheckForNulls(text As Variant) As String
    'I know this function looks retardedly useless.. but I often use it when working with
    'Databases. Sometimes You say: Text1.Text = RS.Fields("UserID")
    'If that field is null, it doesnt display "", it crashes. So now you can say:
    'Text1.Text = CheckForNulls(RS.Fields("UserID"))
    'Since the variant type can handle null expressions, it will convert Null to ""
    
    If IsNull(text) Then
        CheckForNulls = ""
    Else
        CheckForNulls = text
    End If
    
End Function
