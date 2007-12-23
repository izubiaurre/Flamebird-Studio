VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbaltreeview6.ocx"
Begin VB.Form frmProjectBrowser 
   Caption         =   "Project Browser"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "frmProjectBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin vbalTBar6.cToolbarHost tbrhPB 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BorderStyle     =   0
   End
   Begin vbalTBar6.cToolbar tbrPB 
      Height          =   495
      Left            =   2880
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin vbalIml6.vbalImageList iml 
      Left            =   3720
      Top             =   2280
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   29848
      Images          =   "frmProjectBrowser.frx":038A
      Version         =   131072
      KeyCount        =   26
      Keys            =   $"frmProjectBrowser.frx":7842
   End
   Begin vbalTreeViewLib6.vbalTreeView tvProject 
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProjectBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Implements ITDockMoveEvents

Private Const PROJECTKEY As String = "?p" 'Key for project name node (no file or folder of windows can contains ? in its name)
Private Enum ViewModeConstants
    vmPath
    vmCategories
End Enum
Private m_ViewMode As ViewModeConstants

Private WithEvents pMenu As cMenus
Attribute pMenu.VB_VarHelpID = -1

'-------------------------------------------------------------------------------------
'START PROPERTIES DEFINITION
'-------------------------------------------------------------------------------------
Private Property Let ViewMode(newVal As ViewModeConstants)
    If newVal <> m_ViewMode Then 'If the view mode has changed, clear all nodes
        tvProject.Nodes.Clear
    End If
    m_ViewMode = newVal
    RefreshTree
End Property

Private Property Get ViewMode() As ViewModeConstants
    ViewMode = m_ViewMode
End Property
'-------------------------------------------------------------------------------------
'END PROPERTIES DEFINITION
'-------------------------------------------------------------------------------------


'Popup menu action control
Private Sub pMenu_Click(ByVal index As Long)
    Select Case pMenu.ItemKey(index)
        Case "mnuViewModePath" 'View Mode Path
            ViewMode = vmPath
        Case "mnuViewModeCategories" 'View Mode Categories
            ViewMode = vmCategories
    End Select
End Sub

'User clicked in a toolbar button
Private Sub tbrPB_ButtonClick(ByVal lButton As Long)
    Select Case tbrPB.ButtonKey(lButton)
        Case "Add": 'Add file
            If Not openedProject Is Nothing Then
                mnuProjectAddFile
            End If
        Case "Remove": 'Remove file
            If Not openedProject Is Nothing Then
                If Not tvProject.SelectedItem Is Nothing Then 'There is a selected node
                    Dim sFilePath As String
                    If tvProject.SelectedItem.Tag = "File" Then  'It's a file
                        sFilePath = makePathForProject(tvProject.SelectedItem.Key)
                        If MsgBox("Delete " & sFilePath & " from project?", vbQuestion + vbYesNo) = vbYes Then
                            openedProject.RemoveFile sFilePath
                            RefreshTree
                        End If
                    End If
                End If
            End If
        Case "Refresh": 'Refresh
            If Not openedProject Is Nothing Then
                RefreshTree
            End If
    End Select
End Sub

'User selected a dropdown button
Private Sub tbrPB_DropDownPress(ByVal lButton As Long)
    Dim X As Long, Y As Long, rc As RECT, rc2 As RECT
    Dim lX As Long, lY As Long
    
    GetWindowRect tbrhPB.hwnd, rc 'Toolbar dimensions
    GetWindowRect Me.hwnd, rc2 'Form dimensions
    
    With tbrPB
        .GetDropDownPosition lButton, X, Y 'Button position
        lX = rc2.Left + X / Screen.TwipsPerPixelX
        lY = rc.Bottom
        If .ButtonKey(lButton) = "Viewmode" Then
            pMenu.PopupMenu "mnuViewMode", lX, lY, TPM_LEFTALIGN 'Shows mnuViewMode popup menu
        End If
    End With
End Sub

'User clicked on a node
Private Sub tvProject_NodeDblClick(node As vbalTreeViewLib6.cTreeViewNode)
    Dim file As String
    If node.Tag = "File" Then  'Open the node if it's a file
        file = node.Key
        file = makePathForProject(file) 'The key stores the relative path so we must construct the absolute path
        OpenFileByExt file
    End If
End Sub

'Form loads
Private Sub Form_Load()
    'Configure tree view
    With tvProject
        .PathSeparator = "\"
        .ImageList = iml
        .FullRowSelect = False
        .HistoryStyle = False
        .HotTracking = True
        .TabStop = False
        .Style = etvwTreelinesPlusMinusPictureText
        .NoCustomDraw = False
        .LineStyle = etvwRootLines
    End With
    'Configure toolbar
    With tbrPB
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .CreateToolbar 16, False, False, True
        .SetImageList iml.hIml
        
        .AddButton "Add file to project", iml.ItemIndex("ADD") - 1, , , , , "Add"
        .AddButton "Remove file from project", iml.ItemIndex("REMOVE") - 1, , , , , "Remove"
        .AddButton "Refresh list", iml.ItemIndex("REFRESH") - 1, , , , , "Refresh"
'        .AddButton "Move up", iml.ItemIndex("MOVE_UP"), , , , , "MoveUp"
'        .AddButton "MoveDown", iml.ItemIndex("MOVE_DOWN"), , , , , "MoveDown"
        .AddButton , , , , , CTBSeparator
        .AddButton "View Mode", iml.ItemIndex("VIEW_MODE") - 1, , , , CTBDropDownArrow, "Viewmode"
    End With
    'Configure toolbar host
    With tbrhPB
        .Capture tbrPB
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar.bmp"
        End If
    End With
    'Create the popup menu for the ViewMode button
    Dim pId As Long
    Set pMenu = New cMenus
    With pMenu
        .CreateFromNothing (tbrhPB.hwnd)
        .DrawStyle = M_Style
        Set .ImageList = iml
        pId = .AddItem(0, Key:="mnuViewMode")
        .AddItem pId, "Path view", Key:="mnuViewModePath", Image:=iml.ItemIndex("PATH_MODE")
        .AddItem pId, "Categories view", Key:="mnuViewModeCategories", Image:=iml.ItemIndex("CATEGORY_MODE")
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pMenu = Nothing
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION: GetIconForExt()
'DESCRIPTION: Gets the appropiated icon index for an extension
'RETURNS:     The appropiated icon index (in the image list) for the given extension
'-------------------------------------------------------------------------------------
Private Function GetIconForExt(sExt As String) As Integer
    Dim Icon As Integer
    Select Case LCase(sExt)
         Case "prg", "inc", "h": 'Source file
             Icon = iml.ItemIndex("SOURCE")
         Case "map", "fbm", "png": 'Graphic file
             Icon = iml.ItemIndex("GRAPHIC")
         Case "fpg", "fgc": 'Graphic library
             Icon = iml.ItemIndex("GRAPHIC_COL")
         Case "mod", "it", "xm", "mid", "ogg" 'Music file
             Icon = iml.ItemIndex("SOUND")
         Case "pal", "fpl" 'Pal file
             Icon = iml.ItemIndex("PALETTE")
         Case "fnt" 'Font file
             Icon = iml.ItemIndex("FONT")
         Case "wav" 'Sound file
             Icon = iml.ItemIndex("SOUND")
         Case "": 'Folder
             Icon = iml.ItemIndex("FOLDER_CLOSED")
         Case Else: 'Unknown file
             Icon = iml.ItemIndex("OTHER")
     End Select
     GetIconForExt = Icon - 1
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: RefreshTree()
'DESCRIPTION: Forzes the Tree to refresh
'-------------------------------------------------------------------------------------
Public Sub RefreshTree()
    If ViewMode = vmPath Then
        ViewInPathMode
    Else
        ViewInCategoryMode
    End If
    
    'Apperance
    If Not openedProject Is Nothing Then
        With tvProject
            If .Nodes.Exists(PROJECTKEY) Then .Nodes(PROJECTKEY).Bold = True 'Name of the project in bold
            'Main source
            If .Nodes.Exists(openedProject.makePathRelative(openedProject.mainSource)) Then
                .Nodes(openedProject.makePathRelative(openedProject.mainSource)).Bold = True 'Main source in bold
            End If
        End With
    End If
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION: addNode()
'DESCRIPTION: Assures the node which is being added does not exists. In case it does,
'             the function just updates it properties
'-------------------------------------------------------------------------------------
Private Function addNode(Optional Relative As cTreeViewNode, Optional relationship As ETreeViewRelationshipContants = etvwChild, _
                        Optional Key As String, Optional text As String, Optional Image As Long = -1, Optional SelectedImage As Long = -1) As cTreeViewNode
    Dim nodeResult As cTreeViewNode
    'If the node already exists, just change it
    If tvProject.Nodes.Exists(Key) Then
        With tvProject.Nodes(Key)
            .text = text
            .Image = Image
            .SelectedImage = SelectedImage
            .ItemData = Image
        End With
        Set nodeResult = tvProject.Nodes(Key)
    Else 'Add the node if it does not exist
        If Relative Is Nothing Then
            Set nodeResult = tvProject.Nodes.Add(, relationship, Key, text, Image, SelectedImage)
        Else
            Set nodeResult = tvProject.Nodes.Add(Relative, relationship, Key, text, Image, SelectedImage)
        End If
        nodeResult.ItemData = Image
    End If
    Set addNode = nodeResult
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: ViewInPathMode()
'DESCRIPTION: Create tree nodes clasifying files according to its location (relative to
'             the project).
'NOTES:       The Key property of nodes stores the relative path of the file
'-------------------------------------------------------------------------------------
Private Sub ViewInPathMode()
    Dim pId As cTreeViewNode, pId2 As cTreeViewNode
    Dim subitems() As String, Key As String
    Dim Icon As Long, Icon1 As Long
    Dim i As Integer, j As Integer
    Dim bIsNotFile As Boolean
    Dim sProjectTitle As String
    Dim toDeleteKeys() As String 'Stores nodes keys to be deleted
    
    
    If Not openedProject Is Nothing Then
        'If the openedProject has changed, clear all nodes
        If tvProject.Nodes.Exists(PROJECTKEY + openedProject.filename) = False Then
            tvProject.Nodes.Clear
        End If
        'Look for those nodes whose associated file does not exist in project and mark them to delete
        j = -1
        For i = 2 To tvProject.NodeCount
            If openedProject.FileExist(tvProject.Nodes(i).Key) = False And tvProject.Nodes(i).Tag = "File" Then
                j = j + 1
                ReDim Preserve toDeleteKeys(j) As String
                toDeleteKeys(j) = tvProject.Nodes(i).Key
            End If
        Next
        If j >= 0 Then 'Delete those nodes
            Dim node As cTreeViewNode
            For i = 0 To UBound(toDeleteKeys)
                'The node can have a parent folder and this parent folder another one and so on and
                'we must delete these folders if any other file is present at them so look for the
                'top most folder which satisfaces this condition
                Set node = tvProject.Nodes(toDeleteKeys(j))
                While node.Parent.Children.count = 1 And node.Parent.Key <> PROJECTKEY + openedProject.filename
                    Set node = node.Parent
                Wend
                node.Delete
            Next
        End If
        
        'Now add or edit existing nodes
        With openedProject
            sProjectTitle = .projectName & " (" & Dir(.filename) & ")"
            Set pId = addNode(, etvwFirst, PROJECTKEY + openedProject.filename, sProjectTitle, iml.ItemIndex("PROJECT") - 1, iml.ItemIndex("PROJECT") - 1) 'Project node
            
            'For each file in the project
            Dim sFile As Variant
            For Each sFile In .Files
                If varType(sFile) = vbString Then
                    subitems = Split(CStr(sFile), "\", , vbTextCompare) 'Split file path
                    Set pId2 = pId 'Mother node
                    Key = ""
                    
                    For i = 0 To UBound(subitems) 'For each element of the file path
                        bIsNotFile = False
                    
                        Icon = GetIconForExt(FSO.GetExtensionName(LCase(subitems(i)))) 'Icon
                        If Icon = iml.ItemIndex("FOLDER_CLOSED") - 1 Then 'Folder
                            Icon1 = iml.ItemIndex("FOLDER_OPENED") - 1 'Different icon when selected
                            bIsNotFile = True
                        Else
                            Icon1 = Icon
                        End If
                            
                        'Add the node
                        Key = Key & "\" & subitems(i)
                        Set pId2 = addNode(pId2, etvwChild, Right(Key, Len(Key) - 1), subitems(i), Icon, Icon1)      'Nota: el Right() es para quitar el primer \
                        If Not bIsNotFile Then pId2.Tag = "File"
                        'Set node appearance.
                        pId2.Bold = False
                    Next
                End If
            Next
        End With
        
        'Sort items
        Call pId.Sort(etvwItemDataThenAlphabetic)
    End If
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION: ViewInCategoryMode()
'DESCRIPTION: Create tree nodes clasifying files according to its file extension.
'             Since two different categories  can have multiple extensions, a file may
'             appear in two different places
'NOTES:       The Key property of nodes stores the relative path of the file
'-------------------------------------------------------------------------------------
Private Sub ViewInCategoryMode()
    Dim i As Integer, j As Integer
    Dim cat As cCatViewFolder
    Dim sFile As Variant, sExt As String, sName As String
    Dim pId As cTreeViewNode, pId2 As cTreeViewNode, pId3 As cTreeViewNode
    Dim sAllExt As String
    Dim sProjectTitle As String
     
    tvProject.Nodes.Clear
    i = 0
    If Not openedProject Is Nothing Then
        With openedProject
            sProjectTitle = .projectName & " (" & Dir(.filename) & ")"
            'Project node key
            Set pId = tvProject.Nodes.Add(, etvwFirst, PROJECTKEY + openedProject.filename, .projectName, iml.ItemIndex("PROJECT") - 1)
            
            'Add categories and its elements
            For Each cat In .Categories
                i = i + 1
                Set pId2 = tvProject.Nodes.Add(pId, etvwChild, cat.name, cat.name, iml.ItemIndex("FOLDER_CLOSED") - 1, iml.ItemIndex("FOLDER_OPENED") - 1)
                'Filter each archive
                For Each sFile In .Files
                    If varType(sFile) = vbString Then
                        sExt = FSO.GetExtensionName(makePathForProject(CStr(sFile))) 'Extension
                        sName = FSO.GetFileName(makePathForProject(CStr(sFile))) 'File name
                        'If the extension belongs to the category, add a node
                        'TO FIX: IF A FILE BELONGS TO TWO DIFFERENT CATEGORIES, AN ERROR MAY OCCURR
                        If Not sExt = "" And InStr(1, cat.Extensions, "*." & sExt, vbTextCompare) > 0 Then
                            Set pId3 = tvProject.Nodes.Add(pId2, etvwChild, sFile, sName, GetIconForExt(sExt), GetIconForExt(sExt))
                            pId3.Tag = "File"
                        End If
                    End If
                Next
                'sAllExt stores and string containing all file extensions
                sAllExt = sAllExt & "|" & cat.Extensions
            Next
            'Others category
            Set pId2 = tvProject.Nodes.Add(pId, etvwChild, "Other Files", "Other Files", iml.ItemIndex("FOLDER_CLOSED") - 1, iml.ItemIndex("FOLDER_OPENED") - 1)
            'Search those files which do not belong to any category
            For Each sFile In .Files
                If varType(sFile) = vbString Then
                    sExt = FSO.GetExtensionName(makePathForProject(CStr(sFile))) 'Ext
                    sName = FSO.GetFileName(makePathForProject(CStr(sFile))) 'Filename
                    If sExt = "" Or InStr(1, sAllExt, sExt, vbTextCompare) < 1 Then 'Its an "orphan" files
                        Set pId3 = tvProject.Nodes.Add(pId2, etvwChild, sFile, sName, GetIconForExt(sExt), GetIconForExt(sExt))
                        pId3.Tag = "File"
                    End If
                End If
            Next
        End With
    End If
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION: ExpandAll()
'DESCRIPTION: Expands all nodes
'NOTES:       Provisional
'-------------------------------------------------------------------------------------
Private Sub ExpandAll()
    Dim i As Integer
    For i = 1 To tvProject.NodeCount
        If tvProject.Nodes.Exists(i) Then
            tvProject.Nodes(i).expanded = True
        End If
    Next
End Sub

Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant

End Function

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
    On Error GoTo errhandler
    
    tbrhPB.Move Left, Top
    tbrhPB.Width = Right
    tbrhPB.Height = ScaleY(tbrPB.MaxButtonHeight, vbPixels, vbTwips)
    If Bottom - Top > 0 Then
        tvProject.Move Left, Top + tbrhPB.Height, Right, Bottom - Top
    End If
    Exit Function
    
errhandler:
    Debug.Print "ERROR: " & Err.Number & " " & Err.Description
End Function
