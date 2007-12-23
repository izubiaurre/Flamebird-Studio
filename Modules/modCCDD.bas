Attribute VB_Name = "modCCDD"
Option Explicit
Option Base 0


Public Type T_ExternalTool
    Title As String
    Command As String
    Params As String
    UseForFileAssoc As Boolean
End Type

Private m_ExternalTools() As T_ExternalTool
Private m_ExternalToolsCount As Integer

Public Property Get ExternalTools(index As Integer) As T_ExternalTool
    If index < m_ExternalToolsCount And index >= 0 Then
        ExternalTools = m_ExternalTools(index)
    End If
End Property

Public Property Let ExternalTools(index As Integer, newVal As T_ExternalTool)
    If index < m_ExternalToolsCount And index >= 0 Then
        m_ExternalTools(index) = newVal
    End If
End Property

Public Property Get ExternalToolsCount() As Integer
    ExternalToolsCount = m_ExternalToolsCount
End Property

Public Sub AddExternalTool(tool As T_ExternalTool)
    If m_ExternalToolsCount > 0 Then
        ReDim Preserve m_ExternalTools(UBound(m_ExternalTools) + 1) As T_ExternalTool
        With m_ExternalTools(UBound(m_ExternalTools))
            .Title = tool.Title
            .Command = tool.Command
            .UseForFileAssoc = tool.UseForFileAssoc
        End With
    Else
        ReDim m_ExternalTools(0) As T_ExternalTool
        m_ExternalTools(0) = tool
    End If
    m_ExternalToolsCount = m_ExternalToolsCount + 1
End Sub

Public Sub RemoveExternalTool(index As Integer)
    Dim temp() As T_ExternalTool
    Dim i As Integer, j As Integer
    
    If index < 0 Then Exit Sub
    
    If index < m_ExternalToolsCount - 1 Then
        ReDim temp(m_ExternalToolsCount - (index + 2)) As T_ExternalTool
        j = 0
        For i = index + 1 To m_ExternalToolsCount - 1
            temp(j) = m_ExternalTools(i)
            j = j + 1
        Next
        ReDim Preserve m_ExternalTools(m_ExternalToolsCount - 2) As T_ExternalTool
        j = 0
        For i = index To m_ExternalToolsCount - 2
            m_ExternalTools(i) = temp(j)
            j = j + 1
        Next
    Else 'The last item
        If m_ExternalToolsCount > 1 Then
            ReDim Preserve m_ExternalTools(m_ExternalToolsCount - 2) As T_ExternalTool
        Else
            Erase m_ExternalTools
        End If
    End If
    
    m_ExternalToolsCount = m_ExternalToolsCount - 1
End Sub

Public Sub LoadExternalTools()
    Dim n As Integer
    Dim varInteger As Integer, l As Integer
    Dim i As Integer
    Dim tool As T_ExternalTool
    Dim sFile As String

    Erase m_ExternalTools
    m_ExternalToolsCount = 0

    sFile = App.Path & "\Conf\tools.fbconf"
    
    If FSO.FileExists(sFile) Then
        n = FreeFile()
        Open sFile For Binary Access Read As n
        Get n, , m_ExternalToolsCount 'number of tools
        If m_ExternalToolsCount > 0 Then
            ReDim m_ExternalTools(m_ExternalToolsCount) As T_ExternalTool
            Get n, , m_ExternalTools
        End If
    End If

    Close n
End Sub

Public Sub SaveExternalTools()
    Dim n As Integer
    Dim sFile As String
    
    sFile = App.Path & "\Conf\tools.fbconf"
    n = FreeFile()
    Open sFile For Binary Access Write As n
    Put n, , m_ExternalToolsCount
    If m_ExternalToolsCount > 0 Then Put n, , m_ExternalTools
    Close n
End Sub


Public Sub OpenMultipleFiles(ByRef sFiles() As String)
    Dim i As Integer
    For i = LBound(sFiles) To UBound(sFiles)
        OpenFileByExt sFiles(i)
    Next
End Sub

'Shows an Open Dialog and return the selected files in the sFiles param
'NOTE: sFiles() will be a 1-based indexed array when multiple files are returned
Public Function ShowOpenDialog(ByRef sFiles() As String, Optional ByVal filter As String, _
                    Optional showAllFilesFilter As Boolean, Optional multiSelect As Boolean = True) As Integer
                    
    Dim i As Integer
    
    On Error GoTo errhandler

    Dim cdlg As New cCommonDialog
    Dim sFolder As String, fileCount As Long
    
    cdlg.CancelError = True
    cdlg.hwnd = frmMain.hwnd

    If showAllFilesFilter = True Then filter = filter + "|All files (*.*)|*.*"
    cdlg.filter = filter
    
    cdlg.flags = OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST
    If multiSelect = True Then cdlg.flags = cdlg.flags + OFN_ALLOWMULTISELECT
    
    cdlg.ShowOpen
    
    Call cdlg.ParseMultiFileName(sFolder, sFiles, fileCount)
    If fileCount > 1 Then
        For i = LBound(sFiles) To UBound(sFiles)
            sFiles(i) = sFolder & "\" & sFiles(i)
        Next
    Else
        ReDim sFiles(0) As String
        sFiles(0) = Trim(replace(cdlg.filename, Chr(0), " "))
    End If
        
    ShowOpenDialog = fileCount

    Exit Function
    
errhandler:
    If Err.Number = &H7FF3& Then ShowOpenDialog = 0 'User selected cancel
End Function

Public Function ShowSaveDialog(ByVal defaultExt As String, Optional ByVal filter As String, _
            Optional showAllFilesFilter As Boolean = False) As String
    
    Dim cdlg As New cCommonDialog
    
    On Error GoTo errhandler
    
    cdlg.CancelError = True
    If showAllFilesFilter = True Then filter = filter + "|All files (*.*)|*.*"
    cdlg.filter = filter
    cdlg.flags = OFN_OVERWRITEPROMPT Or OFN_NOREADONLYRETURN
    cdlg.defaultExt = defaultExt
    
    cdlg.ShowSave
    ShowSaveDialog = cdlg.filename
    
    Exit Function
errhandler:
    If Err.Number = &H7FF3& Then ShowSaveDialog = 0 'User selected cancel
End Function
