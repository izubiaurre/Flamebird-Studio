VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.Form frmFnt 
   Caption         =   "Fnt"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFnt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   4275
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbStringList 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin vbalTBar6.cToolbar tbrFnt 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin vbalTBar6.cReBar rebar 
      Left            =   2880
      Top             =   2280
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin vbalIml6.vbalImageList ilFnt 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   10332
      Images          =   "frmFnt.frx":038A
      Version         =   131072
      KeyCount        =   9
      Keys            =   "��������"
   End
   Begin VB.PictureBox picScrollBox 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin VB.PictureBox picFnt 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1215
         Left            =   720
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmFnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Dar�o Cutillas)
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
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "Map saved succesfully!"
Private Const MSG_PAINTMAP_ERRORPAINTING = "An error occurred when trying to paint the fnt: "
Private Const MSG_LOAD_ERRORLOADING = "An error occurred loading the fnt: "

Private Const FAST_SCROLL_STEPS As Integer = 12 ' Shift move

Private m_ShowTransparent As Boolean
Private m_SizeIndex As Single
Private m_IsDirty As Boolean 'This should never be set directly. Use the IsDirty property instead
Private m_Title

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1
Private fnt As New cFont
Private WithEvents m_FpgsMenu As cMenus
Attribute m_FpgsMenu.VB_VarHelpID = -1
Private m_FilePath As String
Private m_addToProject As Boolean

Public Current_Color As Long
Public current_Char As Long

'Public curChars() As Long
Public textShow As String
'Public zoom As Long
'Public Alpha As Boolean

Private chars(255) As t_FNT_Char
Private bennu_chars(255) As t_FNT_bennu_Char
'Private palette_entries(255) As t_FNT_Palette_Entry

Private t_magic As t_FNT_Magic
Private t_palette As t_FNT_Palette

Private fontInfo As Long

Dim draw_rect As RECT
Dim Draw_hDC As Long
Dim Draw_hBM As Long

Implements IFileForm
Implements IPropertiesForm

Private Property Get ShowTransparent() As Boolean
    ShowTransparent = m_ShowTransparent
End Property

Private Property Get SizeIndex() As Single
    SizeIndex = m_SizeIndex
End Property

Private Property Let SizeIndex(newVal As Single)
    m_SizeIndex = newVal
    drawEntireImage m_ShowTransparent
End Property

Private Sub ToggleTransparency()
    m_ShowTransparent = Not m_ShowTransparent
    drawEntireImage m_ShowTransparent
End Sub


'Zoom in/out
Private Function ZoomFnt(fIncrement As Single)
    Dim newSizeIndex As Single

    newSizeIndex = m_SizeIndex + fIncrement

    If newSizeIndex < 0.1 Then newSizeIndex = 0.1
    If newSizeIndex > 8 Then newSizeIndex = 8
    If Not newSizeIndex = m_SizeIndex Then
        m_SizeIndex = newSizeIndex
        drawEntireImage m_ShowTransparent
        setStatusMessage
    End If
End Function

Public Function EditPalette() As Long

    'RunPlugin ("Pal_Edit.clsPluginInterface")
    
    Dim objPlugIn As Object
    Dim strResponse As String
    Dim posible As Boolean
    ' Run the Plugin

    Set objPlugIn = CreateObject("Pal_Edit.clsPluginInterface")
    'posible = objPlugIn.CreatePaletteFromArray(map.palette, frmMain, 0, m_Title)
    strResponse = objPlugIn.Run(frmMain)
    
    'if the plug-in returns an error, let us know
'        If strResponse <> vbNullString Then
'            MsgBox strResponse
'        End If
    
End Function

Private Sub cmbStringList_Change()
    drawEntireImage m_ShowTransparent
    setStatusMessage
End Sub

Private Sub cmbStringList_Click()
    drawEntireImage m_ShowTransparent
    setStatusMessage
End Sub

Private Sub Form_Activate()
    'drawEntireImage False
    m_ShowTransparent = False
    ZoomFnt 0
    setStatusMessage
    'frmMain.setStatusMessage ("Font " & fnt.name & " (" & GetMedWidth & "x" & GetMedHeigth & ")")
End Sub

Private Sub Form_Load()

    'textShow = "0123AaBbCc!?&" '��"
    current_Char = 65
    
    'Configure toolbar
    With tbrFnt
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList ilFnt.hIml, CTBImageListNormal
        .CreateToolbar 16, True, True, True, 16
        .AddButton "Zoom In", 0, , , , CTBAutoSize, "ZoomIn"
        .AddButton "Restore Zoom", 1, , , , CTBAutoSize, "ZoomRestore"
        .AddButton "Zoom Out", 2, , , , CTBAutoSize, "ZoomOut"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Toogle transparency", 3, , , "", CTBAutoSize, "ToogleTrans"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Edit palette", 5, , , , CTBAutoSize, "EditPalette"
        .AddButton eButtonStyle:=CTBSeparator
        .AddControl cmbStringList.Hwnd, , "cmbStringList"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Write text to map", 6, , , , CTBAutoSize, "WriteTexToMap"
        .AddButton "Import font", 7, , , , CTBAutoSize, "ImportFont"
        .AddButton "Export font", 8, , , , CTBAutoSize, "ExportFont"
        '.AddButton "...", 4, , , "...", CTBDropDownArrow + CTBAutoSize, "AddToFpg"
    End With
    'Create the rebar
    With rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        End If
        .CreateRebar Me.Hwnd
        .AddBandByHwnd Me.tbrFnt.Hwnd, , True, False
    End With
    rebar.RebarSize
    
    m_SizeIndex = 1
    m_ShowTransparent = False
    
    'Fill the strings in the combo
    cmbStringList.AddItem "AaBbCc012?!��"
    cmbStringList.AddItem "0123456789"
    cmbStringList.AddItem "abcdefghijklmnopqrstuvwxyz"
    cmbStringList.AddItem "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    cmbStringList.AddItem " !" & Chr(34) & "$%&'()*+-./:;<=>?@[\]^_`{|}~"
    Dim extended As String
    Dim i As Integer
    For i = 128 To 255
        extended = extended + Chr(i)
    Next i
    cmbStringList.AddItem extended
    cmbStringList.ListIndex = 0     'Select the first item from the combo

    'Set up scroll bars:
    Set m_cScroll = New cScrollBars
    m_cScroll.create picScrollBox.Hwnd
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msgRes As VbMsgBoxResult
    'Ask for saving if the document is dirty
    If IFileForm_IsDirty = True And UnloadMode <> vbFormMDIForm Then
        msgRes = MsgBox("The file '" & IFileForm_Title & "' is modified. " _
                    & "Save it?", vbYesNoCancel + vbQuestion, "Save")
        If msgRes = vbYes Then 'Save
            SaveFileOfFileForm Me
        ElseIf msgRes = vbCancel Then 'Cancel
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        picScrollBox.Move 0, _
                        ScaleY(rebar.RebarHeight, vbPixels, vbTwips)
        picScrollBox.width = Me.ScaleWidth
        picScrollBox.height = Me.ScaleHeight - picScrollBox.Top
        rebar.RebarSize
        drawEntireImage m_ShowTransparent
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rebar.RemoveAllRebarBands 'Just for safety
End Sub

Private Sub picFnt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbMiddleButton Then
        m_ShowTransparent = Not m_ShowTransparent
        drawEntireImage m_ShowTransparent
    End If
End Sub

Private Sub picScrollBox_Resize()
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lProportion As Long

   ' Pixels are the minimum change size for a screen object.
   ' Therefore we set the scroll bars in pixels.

   lHeight = (picFnt.height - picScrollBox.ScaleHeight) '\ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      'lProportion = lHeight \ ((picScrollBox.ScaleHeight \ Screen.TwipsPerPixelY) + 1)
      lProportion = lHeight \ (picScrollBox.ScaleHeight + 1)
      lProportion = IIf(lProportion = 0, 1, lProportion) 'Ensures no to perform a div by zero
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      picFnt.Top = 0
      m_cScroll.Visible(efsVertical) = False
   End If

   lWidth = (picFnt.width - picScrollBox.ScaleWidth) '\ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      'lProportion = lWidth \ (picScrollBox.ScaleWidth \ Screen.TwipsPerPixelX) + 1
      lProportion = lWidth \ (picScrollBox.ScaleWidth + 1)
      lProportion = IIf(lProportion = 0, 1, lProportion) 'Ensures no to perform a div by zero
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      picFnt.Left = 0
      m_cScroll.Visible(efsHorizontal) = False
   End If
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
   drawEntireImage False
End Sub

Private Sub m_cScroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long, VKPressed As EFSVirtualKeyConstants)
    If (efsVKControl And VKPressed) Then  'Flag Control
        eBar = efsHorizontal 'Desplaz horizontal
    End If
    If (efsVKShift And VKPressed) Then 'Flag shift
        ' Fast smart move
        If eBar = efsHorizontal Then
            lAmount = chars(current_Char).Header.height * m_SizeIndex \ FAST_SCROLL_STEPS * Sgn(lAmount)
        Else
            lAmount = chars(current_Char).Header.width * m_SizeIndex \ FAST_SCROLL_STEPS * Sgn(lAmount)
        End If
    End If
    If (efsVKAlt And VKPressed) Then 'Alt
            ZoomFnt 0.1 * Sgn(lAmount) 'Zoom In/Out
            lAmount = 0
    End If
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   If (eBar = efsHorizontal) Then
      picFnt.Left = m_cScroll.Value(eBar) ' *-Screen.TwipsPerPixelX
   Else
      picFnt.Top = m_cScroll.Value(eBar) ' *  -Screen.TwipsPerPixelY
   End If
End Sub

'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'START INTERFACE IFILEFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
Private Property Get IFileForm_AlreadySaved() As Boolean
    IFileForm_AlreadySaved = IIf(m_FilePath = "", False, True)
End Property

Private Function IFileForm_CloseW() As Long
    MsgBox "TO DO (CLOSE)"
    IFileForm_CloseW = 0
End Function

Private Property Get IFileForm_Title() As String
    Dim sTitle As String
    If IFileForm_FilePath = "" Then
        sTitle = m_Title
    Else
        sTitle = IFileForm_FileName
    End If
    IFileForm_Title = sTitle
End Property

Private Property Get IFileForm_FileName() As String
    IFileForm_FileName = FSO.GetFileName(m_FilePath)
End Property

Private Property Get IFileForm_FilePath() As String
    IFileForm_FilePath = m_FilePath
End Property

Private Function IFileForm_Identify() As EFileFormConstants
    IFileForm_Identify = FF_FNT
End Function

Private Property Get IFileForm_IsDirty() As Boolean
    IFileForm_IsDirty = m_IsDirty
End Property

Private Function IFileForm_Load(ByVal sFile As String) As Long
    Dim lResult As Long
    Dim Ext As String
    
    Screen.MousePointer = vbHourglass
    frmMain.setStatusMessage ("Loading fnt file")
    
    Ext = FSO.GetExtensionName(sFile)
    
    lResult = Load(sFile) 'Load Font
    'lResult = True
    
'    curChars = fnt.getCharMap
'    cur_palette_entries = fnt.getPalEntries
    
    If (lResult) Then
        drawEntireImage False
        m_FilePath = sFile
        IsDirty = False
    Else
        MsgBox MSG_LOAD_ERRORLOADING + fnt.GetLastError, vbCritical
        'MsgBox MSG_LOAD_ERRORLOADING, vbCritical
    End If
    
    Screen.MousePointer = 0
    frmMain.setStatusMessage
    
    IFileForm_Load = lResult
End Function

Private Function IFileForm_NewW(ByVal iUntitledCount As Integer) As Long
    Dim sFiles() As String
    Dim fileCount As Integer
    Dim lResult As Long
    
    m_Title = "Untitled fnt " & CStr(iUntitledCount)
            
    frmMain.setStatusMessage ("Converting file to fnt format...")
    Screen.MousePointer = vbHourglass
    
    fileCount = ShowOpenDialog(sFiles, getFilter("IMPORTABLE_GRAPHICS"), False, False)
    
    If fileCount > 0 Then
        'lResult = fnt.Import(sFiles(0))
        'If lResult <> -1 Then
        '    MsgBox "An error ocurred trying to import the file: " & fnt.GetLastError, vbCritical
        'Else
            m_addToProject = modMenuActions.NewAddToProject
            drawEntireImage False
            IsDirty = True
        'End If
    End If
    
    Screen.MousePointer = 0
    frmMain.setStatusMessage
    
    IFileForm_NewW = lResult
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
    Dim lResult As Long

    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    lResult = Save(sFile) 'Save the map
    If (lResult) Then 'Save succesful
        'Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then addFileToProject sFile
    
        If m_FilePath <> sFile Then 'Show a success message only if the name changed
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        
        IsDirty = False
        m_FilePath = sFile
    Else
        MsgBox MSG_SAVE_ERRORSAVING + fnt.GetLastError, vbCritical
    End If
End Function
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'END INTERFACE IFILEFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'This property is not part of the interface, just helps to reduce code
'by setting the caption of the form properly
Private Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property

'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'INTERFACE IPROPERTIESFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
Private Function IPropertiesForm_GetProperties() As cProperties
 
    Dim props As cProperties
    Set props = New cProperties

    With props
        .Add "Number", "Number", ptText, Me, "EditNumber", GetNumberInfo, False
        .Add "Upper Case", "Uppercase", ptText, Me, "EditUpperCase", GetUpperCaseInfo, False
        .Add "Lower Case", "LowerCase", ptText, Me, "EditLowerCase", GetLowerCaseInfo, False
        .Add "Symbols", "Symbols", ptText, Me, "EditSymbols", GetSymbolsInfo, False
        .Add "Extended", "Extended", ptText, Me, "EditExtended", GetExtendedInfo, False
        .Add "Width", "Width", ptNumeric, Me, "EditWidth", GetMedWidth, False
        .Add "Heigth", "Heigth", ptNumeric, Me, "EditHeigth", GetMedHeigth, False
        
    End With
    
    props("Number").description = "Shows if the fonts contains number characters in it"
    props("UpperCase").description = "Shows if the fonts contains Upper Case characters in it"
    props("LowerCase").description = "Shows if the fonts contains Loqer Case characters in it"
    props("Symbols").description = "Shows if the fonts contains symbols characters in it"
    props("Extended").description = "Shows if the fonts contains extended characters in it"
    props("Width").description = "Medium Width of characters in the font"
    props("Heigth").description = "Medium Heigth of characters in the font"
    
    Set IPropertiesForm_GetProperties = props

    
End Function
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'END INTERFACE IPROPERTIESFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'

Private Sub tbrFnt_ButtonClick(ByVal lButton As Long)
    Select Case tbrFnt.ButtonKey(lButton)
    Case "ZoomIn"
        ZoomFnt 0.5
    Case "ZoomOut"
        ZoomFnt -0.5
    Case "ZoomRestore"
        SizeIndex = 1
        setStatusMessage
        'frmMain.StatusBar.PanelText("MAIN") = map.Width & "," & map.Height & " @ " & m_SizeIndex * 100 & "% - BPP" & map.Depth
    Case "ToogleTrans"
        ToggleTransparency
        tbrFnt.ButtonChecked("ToogleTrans") = ShowTransparent
    Case "EditPalette"
        EditPalette
    Case "WriteTextToMap"
        'WriteTextToMap
    Case "ExportFont"
        'ExportMap
    End Select
End Sub

Private Sub tbrFnt_DropDownPress(ByVal lButton As Long)
'    Dim X As Long, Y As Long
'    Dim lIndex As Long
'    tbrMap.GetDropDownPosition lButton, X, Y
'
'    Select Case tbrMap.ButtonKey(lButton)
'        Case "AddToFpg":
'            createFpgsMenu
'            Call m_FpgsMenu.PopupMenu("FpgsMenu", _
'                Me.ScaleX(Me.Left + X, vbTwips, vbPixels), Me.ScaleY(Y, vbTwips, vbPixels) + rebar.RebarHeight * 1.5, TPM_VERNEGANIMATION + TPM_LEFTALIGN)
'    End Select
End Sub

'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'PAINTING FUNCTIONS
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'

Public Sub drawEntireImage(draw_transparent As Boolean)
    
    Dim bChar() As Byte
    ReDim bChar(chars(current_Char).Header.width * chars(current_Char).Header.height)
    bChar = chars(current_Char).Data
    'ReDim Preserve bChar(chars(current_Char).Header.w_Width * chars(current_Char).Header.Height + chars(current_Char + 1).Header.w_Width * chars(current_Char + 1).Header.Height)
    
    Dim fullHeight, fullWidth As Long
    Dim i As Integer
    Dim curH, curW As Long
    'textShow = "0123AaBbCc!?&" '��"
    textShow = cmbStringList.text
    current_Char = Asc(textShow)
    
    If t_magic.version = c_FNT_version_0 Then
        fullHeight = chars(Asc(Mid(textShow, 1, 1))).Header.height
        fullWidth = chars(Asc(Mid(textShow, 1, 1))).Header.w_Width
        
        For i = 2 To Len(textShow)
            If fullHeight < chars(Asc(Mid(textShow, i, 1))).Header.height Then
                fullHeight = chars(Asc(Mid(textShow, i, 1))).Header.height
            End If
            fullWidth = fullWidth + chars(Asc(Mid(textShow, i, 1))).Header.w_Width
        Next i
    Else
        fullHeight = bennu_chars(Asc(Mid(textShow, 1, 1))).Header.height
        fullWidth = bennu_chars(Asc(Mid(textShow, 1, 1))).Header.width
        Debug.Print "Height: " & bennu_chars(Asc(Mid(textShow, 1, 1))).Header.height & " - Width: " & fullWidth
        
        For i = 2 To Len(textShow)
            If fullHeight < bennu_chars(Asc(Mid(textShow, i, 1))).Header.height Then
                fullHeight = bennu_chars(Asc(Mid(textShow, i, 1))).Header.height
            End If
            fullWidth = fullWidth + bennu_chars(Asc(Mid(textShow, i, 1))).Header.width
            Debug.Print "Height: " & bennu_chars(Asc(Mid(textShow, 1, 1))).Header.height & " - Width: " & fullWidth
        Next i
    End If

        
    Dim PB As Long
    Dim FR As RECT
    Dim TP As POINTAPI
    
    Dim BI8 As BITMAPINFO8Bits
    Dim BI16 As BITMAPINFO16Bits
    Dim BI32 As BITMAPINFO32Bits
    Dim c As Long
    Dim cx As Long, cy As Long
    
    'picFnt.Width = chars(current_Char).Header.Width * m_SizeIndex '* Screen.TwipsPerPixelX
    'picFnt.Height = chars(current_Char).Header.Height * m_SizeIndex '* Screen.TwipsPerPixelY
    picFnt.width = fullWidth * m_SizeIndex
    picFnt.height = fullHeight * m_SizeIndex
  
    Me.picFnt.Cls

    If t_magic.version = c_FNT_version_0 Or t_magic.version = c_FNT_version_8 Then
        BI8.bmiHeader.biBitCount = 8
        BI8.bmiHeader.biClrImportant = 256
        BI8.bmiHeader.biClrUsed = 256
        If t_magic.version = c_FNT_version_0 Then
            BI8.bmiHeader.biWidth = chars(current_Char).Header.w_Width ' + chars(current_Char + 1).Header.Width
            BI8.bmiHeader.biHeight = -chars(current_Char).Header.height
        Else
            BI8.bmiHeader.biWidth = bennu_chars(current_Char).Header.width ' + chars(current_Char + 1).Header.Width
            BI8.bmiHeader.biHeight = -bennu_chars(current_Char).Header.height
        End If
    '    BI8.bmiHeader.biWidth = fullWidth
    '    BI8.bmiHeader.biHeight = -fullHeight
        BI8.bmiHeader.biPlanes = 1
        BI8.bmiHeader.biSize = Len(BI8.bmiHeader)
    
        For c = 0 To 255
            BI8.bmiColors(c).rgbRed = t_palette.Entries(c).Red
            BI8.bmiColors(c).rgbGreen = t_palette.Entries(c).Green
            BI8.bmiColors(c).rgbBlue = t_palette.Entries(c).Blue
        Next
    ElseIf t_magic.version = c_FNT_version_16 Then
        BI16.bmiHeader.biBitCount = 16
        BI16.bmiHeader.biClrImportant = 65536
        BI16.bmiHeader.biClrUsed = 65536
        BI16.bmiHeader.biWidth = bennu_chars(current_Char).Header.width ' + chars(current_Char + 1).Header.Width
        BI16.bmiHeader.biHeight = -bennu_chars(current_Char).Header.height
        BI16.bmiHeader.biPlanes = 1
        BI16.bmiHeader.biSize = Len(BI16.bmiHeader)
    ElseIf t_magic.version = c_FNT_version_32 Then
        BI32.bmiHeader.biBitCount = 32
        BI32.bmiHeader.biClrImportant = 16777216
        BI32.bmiHeader.biClrUsed = 16777216
        BI32.bmiHeader.biWidth = bennu_chars(current_Char).Header.width ' + chars(current_Char + 1).Header.Width
        BI32.bmiHeader.biHeight = -bennu_chars(current_Char).Header.height
        BI32.bmiHeader.biPlanes = 1
        BI32.bmiHeader.biSize = Len(BI32.bmiHeader)
    End If
    
    'TransparentBlt
    'TransparentBlt
    
    
    ' draw background: checked for 16, 32 bpp; white or black for 8 bpp
    If m_ShowTransparent Then
        'If t_magic.version = c_FNT_version_0 Or t_magic.version = c_FNT_version_8 Then
            picFnt.BackColor = RGB(255, 255, 255)
    
            'drawTransBack picFnt.hdc, BI8, picFnt.Width, picFnt.Height
    
    
    '    If draw_transparent = True Then
            'If pic_alpha_static = 1 Then
            '   SetBrushOrgEx Me.hDC, 0, 0, TP
            'Else
                SetBrushOrgEx Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, TP
            'End If
            SetRect FR, IIf(draw_rect.Left < 0, 0, draw_rect.Left), _
                        IIf(draw_rect.Top < 0, 0, draw_rect.Top), _
                        IIf((draw_rect.Left + draw_rect.Right) > (Me.ScaleWidth - 17), (Me.ScaleWidth - 17), (draw_rect.Left + draw_rect.Right)), _
                        IIf((draw_rect.Top + draw_rect.Bottom) > (Me.ScaleHeight - 17), (Me.ScaleHeight - 17), (draw_rect.Top + draw_rect.Bottom))
    '        PB = CreatePatternBrush(hBM_alpha)
            'PB = CreatePatternBrush(HatchStyle50Percent)
            PB = GdipCreateHatchBrush(HatchStyleLargeGrid, 0, RGB(255, 255, 255), PB)
            FillRect Me.picFnt.hdc, FR, PB
            DeleteObject PB
           'SetBrushOrgEx cHdc, TP.X, TP.Y, TP
            SetBrushOrgEx Me.picFnt.hdc, 0, 0, TP
            BI8.bmiColors(0).rgbRed = 0
            BI8.bmiColors(0).rgbGreen = 0
            BI8.bmiColors(0).rgbBlue = 0
            
            
            BI32.bmiColors(0) = 0
            If t_magic.version = c_FNT_version_0 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    chars(current_Char).Header.width * m_SizeIndex, chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, chars(current_Char).Header.width, chars(current_Char).Header.height, chars(current_Char).Data(0), _
                    BI8, 0, vbSrcPaint
            ElseIf t_magic.version = c_FNT_version_16 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI16, 0, vbSrcPaint
            ElseIf t_magic.version = c_FNT_version_32 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI32, 0, vbSrcPaint
            Else
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI8, 0, vbSrcPaint
            End If
            

    
            BI8.bmiColors(0).rgbRed = 255
            BI8.bmiColors(0).rgbGreen = 255
            BI8.bmiColors(0).rgbBlue = 255
    
            If t_magic.version = c_FNT_version_0 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    chars(current_Char).Header.width * m_SizeIndex, chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, chars(current_Char).Header.width, chars(current_Char).Header.height, chars(current_Char).Data(0), _
                    BI8, 0, vbSrcAnd
            ElseIf t_magic.version = c_FNT_version_16 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI16, 0, vbSrcAnd
            ElseIf t_magic.version = c_FNT_version_32 Then
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI32, 0, vbSrcAnd
            Else
                        StretchDIBits Me.picFnt.hdc, draw_rect.Left, draw_rect.Top, _
                    bennu_chars(current_Char).Header.width * m_SizeIndex, bennu_chars(current_Char).Header.height * m_SizeIndex, _
                    0, 0, bennu_chars(current_Char).Header.width, bennu_chars(current_Char).Header.height, bennu_chars(current_Char).Data(0), _
                    BI8, 0, vbSrcAnd
            End If
'        Else
'        ' checked
'            Dim j As Long
'
'            cx = picFnt.Width \ 16
'            cy = picFnt.Height \ 16
'
'            For i = 0 To cx
'                For j = 0 To cy
'                    picFnt.PaintPicture LoadPicture(App.Path & "\Resources\backmaps.bmp"), i * 16, j * 16, 16, 16, 0, 0, 16, 16, vbSrcCopy
'                Next
'            Next
'        End If
    Else
        picFnt.BackColor = RGB(0, 0, 0)
    End If
    
    curH = 0: curW = 0
    
    Dim curChr As Integer
    
    If t_magic.version = c_FNT_version_0 Then
        ' version 0: OLD DIV 8 bit
        For i = 1 To Len(textShow)
            curChr = Asc(Mid(textShow, i, 1))
            curH = chars(curChr).Header.height
                    
            BI8.bmiHeader.biWidth = chars(curChr).Header.w_Width
            BI8.bmiHeader.biHeight = -chars(curChr).Header.height
            BI8.bmiHeader.biSize = Len(BI8.bmiHeader)
            
            If chars(curChr).Header.height <> 0 And chars(curChr).Header.w_Width <> 0 Then
                    StretchDIBits Me.picFnt.hdc, draw_rect.Left + curW * m_SizeIndex, draw_rect.Top + fullHeight - curH, _
                chars(curChr).Header.w_Width * m_SizeIndex, chars(curChr).Header.height * m_SizeIndex, _
                0, 0, chars(curChr).Header.w_Width, chars(curChr).Header.height, chars(curChr).Data(0), _
                BI8, 0, vbSrcCopy
                    curW = curW + chars(curChr).Header.w_Width
            End If
        Next i
    ElseIf t_magic.version = c_FNT_version_8 Then
        ' version 8: BennuGD 8 bit
        For i = 1 To Len(textShow)
            curChr = Asc(Mid(textShow, i, 1))
            curH = bennu_chars(curChr).Header.height
                    
'            BI8.bmiHeader.biWidth = bennu_chars(curChr).Header.Width
'            BI8.bmiHeader.biHeight = -bennu_chars(curChr).Header.Height
'            BI8.bmiHeader.biSize = Len(BI8.bmiHeader)
            
            If bennu_chars(curChr).Header.height <> 0 And bennu_chars(curChr).Header.width <> 0 Then
                    StretchDIBits Me.picFnt.hdc, draw_rect.Left + curW * m_SizeIndex, draw_rect.Top + fullHeight - curH, _
                bennu_chars(curChr).Header.width * m_SizeIndex, bennu_chars(curChr).Header.height * m_SizeIndex, _
                0, 0, bennu_chars(curChr).Header.width, bennu_chars(curChr).Header.height, bennu_chars(curChr).Data(0), _
                BI8, 0, vbSrcCopy
                    curW = curW + bennu_chars(curChr).Header.width
            End If
        Next i
    ElseIf t_magic.version = c_FNT_version_16 Then
        ' version 16: BennuGD 16 bit
        For i = 1 To Len(textShow)
            curChr = Asc(Mid(textShow, i, 1))
            curH = bennu_chars(curChr).Header.height
                    
            BI16.bmiHeader.biWidth = bennu_chars(curChr).Header.width
            BI16.bmiHeader.biHeight = -bennu_chars(curChr).Header.height
            BI16.bmiHeader.biSize = Len(BI16.bmiHeader)
            
            If bennu_chars(curChr).Header.height <> 0 And bennu_chars(curChr).Header.width <> 0 Then
                    StretchDIBits Me.picFnt.hdc, draw_rect.Left + curW * m_SizeIndex, draw_rect.Top + fullHeight - curH, _
                bennu_chars(curChr).Header.width * m_SizeIndex, bennu_chars(curChr).Header.height * m_SizeIndex, _
                0, 0, bennu_chars(curChr).Header.width, bennu_chars(curChr).Header.height, bennu_chars(curChr).Data(0), _
                BI16, 0, vbSrcCopy
                    curW = curW + bennu_chars(curChr).Header.width
            End If
        Next i
    ElseIf t_magic.version = c_FNT_version_32 Then
        ' version 16: BennuGD 32 bit
        For i = 1 To Len(textShow)
            curChr = Asc(Mid(textShow, i, 1))
            curH = bennu_chars(curChr).Header.height
                    
            BI32.bmiHeader.biWidth = bennu_chars(curChr).Header.width
            BI32.bmiHeader.biHeight = -bennu_chars(curChr).Header.height
            BI32.bmiHeader.biSize = Len(BI32.bmiHeader)
            
            If bennu_chars(curChr).Header.height <> 0 And bennu_chars(curChr).Header.width <> 0 Then
                    StretchDIBits Me.picFnt.hdc, draw_rect.Left + curW * m_SizeIndex, draw_rect.Top + fullHeight - curH, _
                bennu_chars(curChr).Header.width * m_SizeIndex, bennu_chars(curChr).Header.height * m_SizeIndex, _
                0, 0, bennu_chars(curChr).Header.width, bennu_chars(curChr).Header.height, bennu_chars(curChr).Data(0), _
                BI32, 0, vbSrcCopy
                    curW = curW + bennu_chars(curChr).Header.width
            End If
        Next i
    End If
    
    picFnt.Top = (picScrollBox.ScaleHeight / 2) - (picFnt.ScaleHeight / 2)
    picFnt.Left = (picScrollBox.ScaleWidth / 2) - (picFnt.ScaleWidth / 2)
    
End Sub

Public Sub Draw_Border()
    Me.ForeColor = vbBlack
    Rectangle Me.picScrollBox.hdc, draw_rect.Left - 1, draw_rect.Top - 1, draw_rect.Left + draw_rect.Right + 1, draw_rect.Top + draw_rect.Bottom + 1
End Sub

Public Function GetNumberInfo() As Boolean ' must be property
    Dim lSucceded As Boolean

    If (fontInfo And 1) = 1 Then
        lSucceded = True
    Else
        lSucceded = False
    End If
    
    GetNumberInfo = lSucceded
End Function

Public Function GetUpperCaseInfo() As Boolean ' must be property
    Dim lSucceded As Boolean
    
    If (fontInfo And 2) = 2 Then
        lSucceded = True
    Else
        lSucceded = False
    End If
    
    GetUpperCaseInfo = lSucceded
End Function

Public Function GetLowerCaseInfo() As Boolean ' must be property
    Dim lSucceded As Boolean

    If (fontInfo And 4) = 4 Then
        lSucceded = True
    Else
        lSucceded = False
    End If
    
    GetLowerCaseInfo = lSucceded
End Function

Public Function GetSymbolsInfo() As Boolean ' must be property
    Dim lSucceded As Boolean

    If (fontInfo And 8) = 8 Then
        lSucceded = True
    Else
        lSucceded = False
    End If
    
    GetSymbolsInfo = lSucceded
End Function

Public Function GetExtendedInfo() As Boolean ' must be property
    Dim lSucceded As Boolean

    If (fontInfo And 16) = 16 Then
        lSucceded = True
    Else
        lSucceded = False
    End If
    
    GetExtendedInfo = lSucceded
End Function

Public Function GetMedWidth() As Long
    Dim lResult As Long
    Dim i As Long

    For i = 0 To 255
        lResult = lResult + chars(i).Header.w_Width
    Next i
    lResult = lResult / 256

    GetMedWidth = lResult
End Function

Public Function GetMedHeigth() As Long
    Dim lResult As Long
    Dim i As Long

    For i = 0 To 255
        lResult = lResult + chars(i).Header.height
    Next i
    lResult = lResult / 256

    GetMedHeigth = lResult
End Function

Private Sub drawTransBack(hdc As Long, bI As BITMAPINFO8Bits, lWidth As Long, lHeight As Long)
    Dim bmp As New cBitmap
    Dim graphics As New cGraphics
    Dim i As Long, j As Long
    Dim resFile As String
    Dim cx As Integer, cy As Integer
    
On Error GoTo ErrHandler
    
    'If Not FSO.FileExists(App.Path & "\Resources\backmaps.bmp") Then Exit Sub
    bmp.LoadFromFile App.Path & "\Resources\backmaps.bmp"

    'graphics.CreateFromHdc hdc
    'graphics.Clear
    
    bI.bmiHeader.biHeight = bmp.height
    bI.bmiHeader.biWidth = bmp.width
    bI.bmiHeader.biSize = Len(bI.bmiHeader)
    
    Debug.Print "a"
    
    StretchDIBits hdc, 0, 0, bmp.width, bmp.height, _
                0, 0, bmp.width, bmp.height, 160, bI, 0, vbSrcCopy
    
'    cx = lWidth \ bmp.Width + 1
'    cy = lHeight \ bmp.Height + 1
'    For i = 0 To cx
'        For j = 0 To cy
'            StretchDIBits hdc, 0, 0, i * bmp.Width, j * bmp.Height, _
'                0, 0, bmp.Width, bmp.Handle, bmp.Height * bmp.Width, bI, 0, scrcopy
'            'graphics.DrawImageRectI bmp.Handle, i * bmp.Width, j * bmp.Height, bmp.Width, bmp.Height
'        Next
'    Next
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

'-------------------------------------------------------------------------------------
'FUNCTION: Load()
'DESCRIPTION: Loads a Fnt file
'RETURNS: True if no error, otherwise False.
'-------------------------------------------------------------------------------------
Public Function Load(Filename As String) As Boolean
'
'    Dim t_magic As t_FNT_Magic
'    Dim t_palette As t_FNT_Palette

    Dim c As Long, C2 As Long
    Dim FileNumber As Long
    Dim Returned_Value As Long
    Dim Must_Destroy As Boolean


    On Error GoTo ErrHandler

    ' open font
    FileNumber = gzopen(Filename, "rb")
    If FileNumber = 0 Then GoTo FAILED
    
    ' Font header
    
    ' magic "Fnt" or "fnx"
    Returned_Value = gzReadStr(FileNumber, t_magic.magic, 3)
    If Returned_Value = 0 Then GoTo FAILED
    If t_magic.magic <> c_FNT_DIV_Magic And t_magic.magic <> c_FNT_BENNU_Magic Then GoTo FAILED

    ' MS-DOS chars
    Returned_Value = gzread(FileNumber, t_magic.MS_DOS_header(0), 4)
    If Returned_Value = 0 Then GoTo FAILED
    If t_magic.MS_DOS_header(0) <> c_FNT__1 Or _
        t_magic.MS_DOS_header(1) <> c_FNT__2 Or _
        t_magic.MS_DOS_header(2) <> c_FNT__3 Or _
        t_magic.MS_DOS_header(3) <> c_FNT__4 _
    Then GoTo FAILED
    
    ' version
    Returned_Value = gzread(FileNumber, t_magic.version, 1)
    If Returned_Value = 0 Then GoTo FAILED


    Destroy
    Must_Destroy = True

    ' Font palette
    ' When version = 0 or 8
    If t_magic.version = c_FNT_version_0 Or t_magic.version = c_FNT_version_8 Then
        Returned_Value = gzread(FileNumber, t_palette.Entries(0), 256 * 3)
        If Returned_Value = 0 Then GoTo FAILED
        'Returned_Value = gzread(FileNumber, t_palette.UnusedBytes(0), 580)
        Returned_Value = gzread(FileNumber, t_palette.UnusedBytes(0), 576)
        If Returned_Value = 0 Then GoTo FAILED
    
        ' if  version 0 (DIV), must change into bennu format
        If t_magic.version = c_FNT_version_0 Then
            For c = 0 To 255        ' change into bennu format
'                palette_entries(c).Red = t_palette.Entries(c).Red * 4
'                palette_entries(c).Green = t_palette.Entries(c).Green * 4
'                palette_entries(c).Blue = t_palette.Entries(c).Blue * 4
                t_palette.Entries(c).Red = t_palette.Entries(c).Red * 4
                t_palette.Entries(c).Green = t_palette.Entries(c).Green * 4
                t_palette.Entries(c).Blue = t_palette.Entries(c).Blue * 4
            Next
        End If
    End If

    ' charset group info
    Returned_Value = gzread(FileNumber, charset, 4)
    If Returned_Value = 0 Then GoTo FAILED

    ' Font char descriptors
    ' here starts the hard work
    
    ' old DIV format
    If t_magic.version = c_FNT_version_0 Then
        For c = 0 To 255
            Returned_Value = gzread(FileNumber, chars(c).Header, 4 * 4)
            If Returned_Value = 0 Then GoTo FAILED
    
            If chars(c).Header.width <> 0 And chars(c).Header.height <> 0 Then
                If chars(c).Header.width Mod 4 = 0 Then
                    chars(c).Header.w_Width = chars(c).Header.width
                Else
                    chars(c).Header.w_Width = Fix(chars(c).Header.width \ 4) * 4 + 4
                End If
                ReDim chars(c).Data(chars(c).Header.w_Width * chars(c).Header.height)
            End If
    
        Next
    
        For c = 0 To 255
        'loads the data for each char
            If chars(c).Header.width <> 0 And chars(c).Header.height <> 0 Then
                gzseek FileNumber, chars(c).Header.File_Offset, 0
                For C2 = 0 To chars(c).Header.height - 1
                    Returned_Value = gzread(FileNumber, chars(c).Data(C2 * chars(c).Header.w_Width), chars(c).Header.width)
                    If Returned_Value = 0 Then GoTo FAILED
                Next
            End If
        Next
    ElseIf t_magic.version = c_FNT_version_1 Then   ' 1 bit format bennu_chars
        
    ElseIf t_magic.version = c_FNT_version_8 Then   ' 8 bit format
        For c = 0 To 255
            Returned_Value = gzread(FileNumber, bennu_chars(c).Header, 7 * 4)
            If Returned_Value = 0 Then GoTo FAILED
    
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                ReDim bennu_chars(c).Data(bennu_chars(c).Header.width * bennu_chars(c).Header.height)
            End If
            Debug.Print c
        Next
    
        For c = 0 To 255
        'loads the data for each char
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                gzseek FileNumber, bennu_chars(c).Header.data_offset, 0
                For C2 = 0 To bennu_chars(c).Header.height - 1
                    Returned_Value = gzread(FileNumber, bennu_chars(c).Data(C2 * bennu_chars(c).Header.width), bennu_chars(c).Header.width)
                    If Returned_Value = 0 Then GoTo FAILED
                Next
            End If
        Next
    ElseIf t_magic.version = c_FNT_version_16 Then  ' 16 bit format
        For c = 0 To 255
            Returned_Value = gzread(FileNumber, bennu_chars(c).Header, 7 * 4)
            If Returned_Value = 0 Then GoTo FAILED
    
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                ReDim bennu_chars(c).Data(2 * bennu_chars(c).Header.width * bennu_chars(c).Header.height)   ' 2 bytes per pixel
            End If
    
        Next
    
        For c = 0 To 255
        'loads the data for each char
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                gzseek FileNumber, bennu_chars(c).Header.data_offset, 0
                For C2 = 0 To bennu_chars(c).Header.height - 1
                    Returned_Value = gzread(FileNumber, bennu_chars(c).Data(C2 * bennu_chars(c).Header.width * 2), bennu_chars(c).Header.width * 2)
                    If Returned_Value = 0 Then GoTo FAILED
                Next
            End If
        Next
    ElseIf t_magic.version = c_FNT_version_32 Then  ' 32 bit format
        For c = 0 To 255
            Returned_Value = gzread(FileNumber, bennu_chars(c).Header, 7 * 4)
            If Returned_Value = 0 Then GoTo FAILED
    
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                ReDim bennu_chars(c).Data(4 * bennu_chars(c).Header.width * bennu_chars(c).Header.height)   ' 4 bytes per pixel
            End If
    
        Next
    
        For c = 0 To 255
        'loads the data for each char
            If bennu_chars(c).Header.width <> 0 And bennu_chars(c).Header.height <> 0 Then
                gzseek FileNumber, bennu_chars(c).Header.data_offset, 0
                For C2 = 0 To bennu_chars(c).Header.height - 1
                    Returned_Value = gzread(FileNumber, bennu_chars(c).Data(C2 * bennu_chars(c).Header.width * 4), bennu_chars(c).Header.width * 4)
                    If Returned_Value = 0 Then GoTo FAILED
                Next
            End If
        Next
    End If

    gzclose FileNumber
    Load = True
    Exit Function

FAILED:
    gzclose FileNumber
    If Must_Destroy Then Destroy
    Load = False
ErrHandler:
    If Err.Number > 0 Then ShowError "frmFnt.Load"

End Function

'-------------------------------------------------------------------------------------
'FUNCTION: Destroy()
'DESCRIPTION: Destroyes a Fnt file
'RETURNS: True if no error, otherwise False.
'-------------------------------------------------------------------------------------
Private Function Destroy() As Boolean
    'If Is_Created = False Then Destroy_Font = False: Exit Function

    Dim c As Long

    For c = 0 To 255
        t_palette.Entries(c).Red = 0
        t_palette.Entries(c).Green = 0
        t_palette.Entries(c).Blue = 0

        chars(c).Header.width = 0
        chars(c).Header.height = 0
        chars(c).Header.Vertical_Offset = 0
        chars(c).Header.File_Offset = 0
        Erase chars(c).Data
    Next

    'Is_Created = False
    Destroy = True
End Function

'-------------------------------------------------------------------------------------
'FUNCTION: Save()
'DESCRIPTION: Saves a Fnt file
'RETURNS: True if no error, otherwise False.
'-------------------------------------------------------------------------------------
Public Function Save(sFile As String) As Long
    Dim lSucceded As Boolean
    lSucceded = True
    Save = lSucceded
End Function



Private Sub setStatusMessage()
    Dim strMessage As String
    If t_magic.version = c_FNT_version_0 Then
        strMessage = "8 BPP DIV FNT - "
    Else
        strMessage = t_magic.version & " BPP BennuGD FNT - "
    End If
    
    strMessage = strMessage & GetMedWidth & "x" & GetMedHeigth & " @ " & m_SizeIndex * 100 & "%"
    frmMain.setStatusMessage (strMessage)
End Sub
