VERSION 5.00
Begin VB.Form frmMSDOSCommand 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FlameCommander"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMSDOSCommand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOutput 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4800
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   6915
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   4800
         Width           =   6255
      End
      Begin VB.PictureBox pic1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   449
         TabIndex        =   14
         Top             =   0
         Width           =   6735
         Begin VB.Label lblMain 
            BackStyle       =   0  'Transparent
            Caption         =   "Usual command call are stored here to run them fastly without get out of the program and type in the command-line."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   540
            TabIndex        =   16
            Top             =   240
            Width           =   5475
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5100
         TabIndex        =   8
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   6375
         Begin VB.CheckBox chkShowOutput 
            Caption         =   "Show command output"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   6
            Top             =   2760
            Width           =   2595
         End
         Begin VB.CheckBox chkMSDOSWindow 
            Caption         =   "Show MS-DOS window"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   2760
            Width           =   3015
         End
         Begin VB.ComboBox cmbFolder 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   2160
            Width           =   4575
         End
         Begin VB.CheckBox chkFolder 
            Caption         =   "Run in this path"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   1800
            Width           =   2415
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   4800
            ScaleHeight     =   2415
            ScaleWidth      =   1455
            TabIndex        =   13
            Top             =   720
            Width           =   1455
            Begin VB.CommandButton cmdBrowseFolder 
               Caption         =   "Browse"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   60
               TabIndex        =   4
               Top             =   1440
               Width           =   1335
            End
            Begin VB.CommandButton cmdBrowseFile 
               Caption         =   "Browse"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   60
               TabIndex        =   1
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.ComboBox cmbCommand 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmMSDOSCommand.frx":000C
            Left            =   120
            List            =   "frmMSDOSCommand.frx":000E
            TabIndex        =   0
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label lblHeader 
            Caption         =   "Insert  ""%f"" in the call as parameter, to enter automatically the active file"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   60
            TabIndex        =   15
            Top             =   240
            Width           =   6195
         End
         Begin VB.Label lblCommand 
            Caption         =   "Command:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "frmMSDOSCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const CREATE_NO_WINDOW As Long = &H8000000

Private Const MSG_EXE_ERR1 As String = "The process has not been launched sucessfully."
Private Const MSG_EXE_ERR2 As String = "Possible cause: program doesn't exists or path is not correct"
Private Const MSG_EXE_ERR_TITLE As String = "Command execution"

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

Dim FSO As New FileSystemObject
Public WithEvents c As cBrowseForFolder
Attribute c.VB_VarHelpID = -1

'Private Type CALL_LIST
'    commands(100) As String
'    lastCommandIndex As Integer
'    paths(100) As String
'    lastPathIndex As Integer
'    ' to run out this form the last command
'    publicLastCommand As String
'    publicLastPath As String
'End Type
'
'Public callList As CALL_LIST
'Dim lastCommandEnabled As Boolean

Private Sub chkFolder_Click()
    If chkFolder.Value Then
        cmbFolder.Enabled = True
        cmdBrowseFolder.Enabled = True
    Else
        cmbFolder.Enabled = False
        cmdBrowseFolder.Enabled = True
    End If
End Sub

Private Sub chkShowOutput_Click()
    If chkShowOutput.Value Then
        Height = 7800
    Else
        Height = 5235
    End If
End Sub

Private Sub cmdRun_Click()
    Dim lResult As Long
    Dim lShow As Long
    Dim sParams As String, sDir As String
    
    On Error GoTo errhandler
    
    lastCommandEnabled = True
    
    If cmbCommand.text <> "" Then
        callList.commands(callList.lastCommandIndex) = cmbCommand.text
        cmbCommand.AddItem cmbCommand.text
        callList.lastCommandIndex = callList.lastCommandIndex + 1
        callList.publicLastCommand = cmbCommand.text
    Else
        Exit Sub
    End If
    
    If chkFolder.Value Then
        If cmbFolder.text <> "" Then
            sDir = cmbFolder.text
            callList.paths(callList.lastPathIndex) = sDir
            cmbFolder.AddItem sDir
            callList.lastPathIndex = callList.lastPathIndex + 1
            callList.publicLastPath = cmbFolder.text
        Else
            sDir = vbNullString
            callList.publicLastPath = ""
        End If
    End If
    
    If chkMSDOSWindow.Value Then
        lShow = CREATE_NEW_CONSOLE
    Else
        lShow = CREATE_NO_WINDOW
    End If
    
    ' run current command
    If chkShowOutput.Value Then
        If ExecuteCommand(cmbCommand.text & " > " & Chr(34) & App.Path & "\output.txt" & Chr(34), lShow, sDir) = 0 Then
            txtOutput.text = FSO.OpenTextFile(App.Path & "\output.txt").ReadAll
            txtOutput.Visible = True
            Kill (App.Path & "\output.txt")
        Else
           MsgBox MSG_EXE_ERR1 & vbCrLf & MSG_EXE_ERR2, 16, MSG_EXE_ERR_TITLE
        End If
    Else
        If ExecuteCommand(cmbCommand.text, lShow, sDir) = -1 Then
            MsgBox MSG_EXE_ERR1 & vbCrLf & MSG_EXE_ERR2, 16, MSG_EXE_ERR_TITLE
        End If
    End If
    
    If callList.lastCommandIndex = 100 Or callList.lastPathIndex = 100 Then
        refreshCallList
    End If
    
    Exit Sub
    
errhandler:
    Resume Next
    
End Sub

Private Sub cmdBrowseFile_Click()
    Dim sExe As String
    
    On Error GoTo errhandler

    sExe = OpenFile32.OpenDialog(Me, "All executable files (*.exe; *.cmd; *.bat; *.com)|*.exe;*.cmd;*.bat;*.com|", "Choose executable file", callList.paths(callList.lastPathIndex))

    If sExe <> "" Then
        If vbYes = MsgBox("Do you want to use this folder as running path?", vbYesNo + vbExclamation) Then
            chkFolder.Value = 1
            cmbFolder.Enabled = True
            cmdBrowseFolder.Enabled = True
            cmbFolder.text = FSO.GetParentFolderName(sExe)
        End If
        cmbCommand.text = FSO.GetBaseName(sExe)
    End If
    
    Exit Sub
    
errhandler:
    Resume Next
End Sub

Private Sub cmdBrowseFolder_Click()
    Dim sExe As String
    Dim s As String

    On Error GoTo errhandler
    
    c.hwndOwner = Me.Hwnd
    c.InitialDir = App.Path
    c.FileSystemOnly = True
    c.StatusText = True
    c.UseNewUI = True
    sExe = c.BrowseForFolder

    If sExe <> "" Then
        cmbFolder.text = FSO.GetParentFolderName(sExe)
    End If
    
    Exit Sub
    
errhandler:
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    saveCommandHistory
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set c = New cBrowseForFolder
    
    Height = 5235
    
    pic1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")
    ' populate comboBoxes from ini file
    populateHistory
End Sub

'Public Function callLast()
'    If executeCommand(callList.commands(callList.lastCommandIndex), CREATE_NO_WINDOW, callList.paths(callList.lastPathIndex)) = -1 Then
'        MsgBox MSG_EXE_ERR1 & vbCrLf & MSG_EXE_ERR2, 16, MSG_EXE_ERR_TITLE
'    End If
'End Function
'
Public Function ExecuteCommand(file As String, showConsole As Long, InitDir As String)

    Dim Data As Integer
    Dim i As Integer

    Dim Value As Long
    Dim start As STARTUPINFO
    Dim process As PROCESS_INFORMATION

    'data = 0
    start.cb = Len(start)


    Value = CreateProcessA(0&, file, 0&, 0&, 1&, showConsole, 0&, 0&, start, process)

    Value = WaitForSingleObject(process.hProcess, -1&)

    If Value = -1 Then
        'data = 1
        ExecuteCommand = -1
    Else
        Value = CloseHandle(process.hProcess)
        ExecuteCommand = 0
    End If

End Function

' if callList is full, we delete 10 oldest elements
Private Sub refreshCallList()
    Dim i As Integer
    
    ' Clear oldest 10 elements in commands list
    If callList.lastCommandIndex = 100 Then
        For i = 10 To 100
            callList.commands(i - 10) = callList.commands(i)
        Next i
        callList.lastCommandIndex = 90
    End If
    
    ' Clear oldest 10 elements in paths list
    If callList.lastPathIndex = 100 Then
        For i = 10 To 100
            callList.paths(i - 10) = callList.paths(i)
        Next i
        callList.lastPathIndex = 90
    End If
End Sub

Public Function isEnabledLastCommand()
    If lastCommandEnabled Then
        isEnabledLastCommand = True
    Else
        isEnabledLastCommand = False
    End If
End Function

Public Sub disableLastCommand()
    lastCommandEnabled = False
End Sub

' Saves commands history
Public Sub saveCommandHistory()
    Dim sCommands As String, sPaths As String
    Dim i As Integer
    
    On Error GoTo errhandler

    With Ini
        .Path = App.Path & "\Conf\command.ini"
        .Section = "General"
        .Key = "lastCommand"
        .Default = "0"
        .Value = callList.lastCommandIndex
        .Key = "lastPath"
        .Default = "0"
        .Value = callList.lastPathIndex
        
        .Section = "Commands"
        For i = 0 To 100
            .Key = "cmd_" & i
            .Default = ""
            If i <= callList.lastCommandIndex Then
                .Value = callList.commands(i)
            Else
                .Value = ""
            End If
        Next i

        .Section = "Paths"
        For i = 0 To 100
            .Key = "path_" & i
            .Default = ""
            If i <= callList.lastPathIndex Then
                .Value = callList.paths(i)
            Else
                .Value = ""
            End If
        Next i
        
        If Not (.Success) Then
           MsgBox "Failed to save value.", vbInformation
        End If
    End With
    
    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("frmMSDOS.saveCommandHistory")
End Sub
'
'' Loads commands history
'Public Sub loadCommandHistory()
'    Dim sCommands As String, sPaths As String
'    Dim i As Integer
'
'    On Error GoTo errhandler:
'
'       With Ini
'        .Path = App.Path & "\Conf\command.ini"
'        .Section = "General"
'        .Key = "lastCommand"
'        callList.lastCommandIndex = .Value
'        .Key = "lastPath"
'
'        callList.lastPathIndex = .Value
'
'        .Section = "Commands"
'        For i = 0 To 100
'            .Key = "cmd_" & i
'            If i <= callList.lastCommandIndex Then
'                callList.commands(i) = .Value
'            Else
'                callList.commands(i) = ""
'            End If
'        Next i
'
'        .Section = "Paths"
'        For i = 0 To 100
'            .Key = "path_" & i
'            If i <= callList.lastPathIndex Then
'                callList.paths(i) = .Value
'            Else
'                callList.paths(i) = ""
'            End If
'        Next i
'    End With
'
'    Exit Sub
'errhandler:
'    If Err.Number > 0 Then ShowError ("frmMSDOSCommand.loadCommandHistory")
'End Sub

Public Sub populateHistory()
    Dim i As Integer
    
    For i = 0 To callList.lastCommandIndex
        If callList.commands(i) <> "" Then
            cmbCommand.AddItem callList.commands(i)
        End If
    Next i
    For i = 0 To callList.lastPathIndex
        If callList.paths(i) <> "" Then
            cmbFolder.AddItem callList.paths(i)
        End If
    Next i
End Sub

Public Sub callLastCommand()
    If isEnabledLastCommand Then
        ExecuteCommand callList.publicLastCommand, CREATE_NO_WINDOW, callList.publicLastPath
    Else
        If vbYes = MsgBox("Last command is disabled or doesn't exists" & vbCrLf & "Open FlameCommander?", vbInformation + vbYesNo) Then
            frmMSDOSCommand.Show 1
        End If
    End If
End Sub

Public Sub clearCommandHistory()
    Dim i As Integer
    
    callList.lastCommandIndex = 0
    callList.lastPathIndex = 0
    callList.publicLastCommand = 0
    callList.publicLastPath = 0
    
    For i = 0 To 100
        callList.commands(i) = ""
    Next i
    For i = 0 To 100
        callList.paths(i) = ""
    Next i
    
    disableLastCommand
End Sub
