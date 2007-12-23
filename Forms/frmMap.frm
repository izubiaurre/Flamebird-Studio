VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   4275
   WindowState     =   2  'Maximized
   Begin vbalTBar6.cToolbar tbrMap 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin vbalTBar6.cReBar rebar 
      Left            =   2160
      Top             =   0
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.PictureBox picScrollBox 
      BackColor       =   &H80000010&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin vbalIml6.vbalImageList ilMap 
      Left            =   3480
      Top             =   1920
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   8036
      Images          =   "frmMap.frx":2B8A
      Version         =   131072
      KeyCount        =   7
      Keys            =   "ÿÿÿÿÿÿ"
   End
End
Attribute VB_Name = "frmMap"
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
'MSG Constants (for future multi-language support)
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "Map saved succesfully!"
Private Const MSG_PAINTMAP_ERRORPAINTING = "An error occurred when trying to paint the map: "
Private Const MSG_LOAD_ERRORLOADING = "An error occurred loading the map: "

Private Const FAST_SCROLL_STEPS As Integer = 12 'desplazamiento con Shift

Private m_ShowTransparent As Boolean
Private m_SizeIndex As Single
Private m_IsDirty As Boolean 'This should never be set directly. Use the IsDirty property instead
Private m_Title

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1
Private map As New cMap
Private WithEvents m_FpgsMenu As cMenus
Attribute m_FpgsMenu.VB_VarHelpID = -1
Private m_FilePath As String
Private m_addToProject As Boolean

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
    PaintMap
End Property

Private Sub ToggleTransparency()
    m_ShowTransparent = Not m_ShowTransparent
    PaintMap
End Sub

'Pinta un mapa en el picMap
Private Sub PaintMap()
    Screen.MousePointer = vbHourglass
    Dim lW As Long, lH As Long
    
    If map.Available Then
        With picMap
            .AutoRedraw = True
            .Cls
            .BackColor = QBColor(0)
            .Width = ScaleX(map.Width, 3, 1) * m_SizeIndex
            .Height = ScaleY(map.Height, 3, 1) * m_SizeIndex
            lW = ScaleX(.Width, 1, 3)
            lH = ScaleY(.Height, 1, 3)
            If Not map.Draw(picMap.hdc, 0, 0, lW, lH, m_ShowTransparent) Then 'Pintar el mapa
                MsgBox MSG_PAINTMAP_ERRORPAINTING & map.GetLastError, vbCritical
            End If
            .Refresh
            picScrollBox_Resize
            .AutoRedraw = False
        End With
    End If
    
    Screen.MousePointer = vbDefault
End Sub

'Zoom in/out
Private Function ZoomMap(fIncrement As Single)
    Dim newSizeIndex As Single

    newSizeIndex = m_SizeIndex + fIncrement

    If newSizeIndex < 0.1 Then newSizeIndex = 0.1
    If newSizeIndex > 8 Then newSizeIndex = 8
    If Not newSizeIndex = m_SizeIndex Then
        m_SizeIndex = newSizeIndex
        PaintMap
    End If
End Function

Public Function EditCP(ParamArray args()) As Long
    frmCPEditor.Show 0, frmMain
    frmCPEditor.SelectMap map
    EditCP = -1
End Function

Public Function EditPalette() As Long
    
    If map.Depth = 8 Then
        
        'RunPlugin ("Pal_Edit.clsPluginInterface")
        
        Dim objPlugIn As Object
        Dim strResponse As String
        ' Run the Plugin

        Set objPlugIn = CreateObject("Pal_Edit.clsPluginInterface")
        'strResponse = objPlugIn.CreatePaletteFromArray(map.palette, frmMain, 0, m_Title)
        strResponse = objPlugIn.Run(frmMain)
        
        'if the plug-in returns an error, let us know
'        If strResponse <> vbNullString Then
'            MsgBox strResponse
'        End If
        
    Else
        MsgBox "This map is not 8 bpp depth"
    End If
End Function

Public Function EditDescription(ByVal newVal As String) As Long
    map.Description = newVal
    EditDescription = -1
    IsDirty = True
End Function

'Callback function for using with the depth
Public Function EditDepth(ByVal newIndex As Integer) As Long
    Dim res As VbMsgBoxResult
    Dim succeded As Long
    
    'newIndex = CInt(newIndex)
    Select Case newIndex
    Case 0: '8bpp
        If (map.Depth = 16) Then
            res = MsgBox("You have choosen to convert this 16bpp image to a 8bpp format." _
            & "Note that after this process, converting then again to a 16bpp format does " _
            & "not ensure to get the same original image. Also note that depending on the image" _
            & " size and your computer, this process could take some time to complete." _
            & vbCrLf & "Are you sure to proceed?" _
            , vbYesNo + vbQuestion)
            If (res = vbYes) Then
                MousePointer = vbHourglass
                succeded = map.ConvertTo8bpp 'Convertir a 8bpp
                If (succeded) Then
                    MsgBox "The image was correctly converted to 8bpp"
                End If
                MousePointer = 0
            End If
        End If
    Case 1: '16bpp
        If (map.Depth = 8) Then
            res = MsgBox("You have choosen to convert this 8bpp image to a 16bpp format." _
            & "Note that after this process, converting then again to a 8bpp format does " _
            & "not ensure to get the same original image." & vbCrLf & "Are you sure to proceed?" _
            , vbYesNo + vbQuestion)
            If (res = vbYes) Then
                MousePointer = vbHourglass
                succeded = map.ConvertTo16bpp 'Convertir a 16bpp
                If (succeded) Then
                    MsgBox "Image succefusly converted to 16bpp"
                End If
                MousePointer = 0
            Else
                newIndex = 1
            End If
        End If
    End Select
    PaintMap
    
    If succeded = -1 Then IsDirty = True
    EditDepth = succeded
End Function

'Callback function for using with the code property
Public Function EditCode(ByVal newVal As Integer) As Long
    Dim lSucceded As Long
    
    If (IsNumeric(newVal)) Then
        newVal = CLng(newVal)
        map.Code = newVal
        lSucceded = -1
        IsDirty = True
    End If
    EditCode = lSucceded
End Function

Private Sub Form_Load()
    'Configure toolbar
    With tbrMap
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList ilMap.hIml, CTBImageListNormal
        .CreateToolbar 16, True, True, True, 16
        .AddButton "Zoom In", 0, , , , CTBAutoSize, "ZoomIn"
        .AddButton "Restore Zoom", 1, , , , CTBAutoSize, "ZoomRestore"
        .AddButton "Zoom Out", 2, , , , CTBAutoSize, "ZoomOut"
        .AddButton "Toogle transparency", 3, , , "", CTBAutoSize, "ToogleTrans"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Edit control points", 6, , , , CTBAutoSize, "EditCP"
        .AddButton "Edit palette", 5, , , , CTBAutoSize, "EditPalette"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Adds a copy of the map to one of the opened or project Fpgs", 4, , , "Add to FPG", CTBDropDownArrow + CTBAutoSize, "AddToFpg"
    End With
    'Create the rebar
    With Rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar.bmp"
        End If
        .CreateRebar Me.hwnd
        .AddBandByHwnd tbrMap.hwnd, , True, False
    End With
    Rebar.RebarSize
    
    m_SizeIndex = 1
    m_ShowTransparent = False

    'Set up scroll bars:
    Set m_cScroll = New cScrollBars
    m_cScroll.Create picScrollBox.hwnd
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
                        ScaleY(Rebar.RebarHeight, vbPixels, vbTwips)
        picScrollBox.Width = Me.ScaleWidth
        picScrollBox.Height = Me.ScaleHeight - picScrollBox.Top
        Rebar.RebarSize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rebar.RemoveAllRebarBands 'Just for safety
End Sub

Private Sub m_FpgsMenu_Click(ByVal index As Long)
    Dim values() As String
    Dim bError As Boolean
    Dim frm As Form
    Dim fpgForm As frmFpg
    Dim ff As IFileForm
    Dim fpg As cFpg
    Dim mode As Integer
    
    values() = Split(m_FpgsMenu.ItemKey(index), "|")
    
    'Get the fpg depending on if it is an opened file or a project file
    If values(0) = "OPENED" Then
        mode = 0
        For Each frm In Forms
            If frm.hwnd = CLng(values(1)) Then
                Set ff = frm
                Set fpgForm = frm
                Set fpg = fpgForm.fpg
            End If
        Next
    ElseIf values(0) = "PROJECT" Then
        mode = 1
        Set fpg = New cFpg
        If fpg.Load(makePathForProject(values(1))) <> -1 Then bError = True
    End If
    
    If bError = False Then
        If addMapToFpg(fpg, getMapCopy(map)) = True Then
            If mode = 0 Then 'If it is an OPENED FPG, set IsDirty to true
                fpgForm.IsDirty = True
            ElseIf mode = 1 Then 'If not, save the changes
                fpg.Save makePathForProject(values(1))
            End If
        End If
    End If
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbMiddleButton Then
        m_ShowTransparent = Not m_ShowTransparent
        PaintMap
    End If
End Sub

Private Sub picScrollBox_Resize()
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lProportion As Long

   ' Pixels are the minimum change size for a screen object.
   ' Therefore we set the scroll bars in pixels.

   lHeight = (picMap.Height - picScrollBox.ScaleHeight) \ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      lProportion = lHeight \ (picScrollBox.ScaleHeight \ Screen.TwipsPerPixelY) + 1
      lProportion = IIf(lProportion = 0, 1, lProportion) 'Ensures no to perform a div by zero
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      picMap.Top = 0
      m_cScroll.Visible(efsVertical) = False
   End If

   lWidth = (picMap.Width - picScrollBox.ScaleWidth) \ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      lProportion = lWidth \ (picScrollBox.ScaleWidth \ Screen.TwipsPerPixelX) + 1
      lProportion = IIf(lProportion = 0, 1, lProportion) 'Ensures no to perform a div by zero
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      picMap.Left = 0
      m_cScroll.Visible(efsHorizontal) = False
   End If
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
End Sub

Private Sub m_cScroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long, VKPressed As EFSVirtualKeyConstants)
    If (efsVKControl And VKPressed) Then  'Flag Control
        eBar = efsHorizontal 'Desplaz horizontal
    End If
    If (efsVKShift And VKPressed) Then 'Flag shift
        'Desplazamiento rápido inteligente
        If eBar = efsHorizontal Then
            lAmount = map.Height * m_SizeIndex \ FAST_SCROLL_STEPS * Sgn(lAmount)
        Else
            lAmount = map.Width * m_SizeIndex \ FAST_SCROLL_STEPS * Sgn(lAmount)
        End If
    End If
    If (efsVKAlt And VKPressed) Then 'Alt
            ZoomMap 0.1 * Sgn(lAmount) 'Zoom In/Out
            lAmount = 0
    End If
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   If (eBar = efsHorizontal) Then
      picMap.Left = -Screen.TwipsPerPixelX * m_cScroll.value(eBar)
   Else
      picMap.Top = -Screen.TwipsPerPixelY * m_cScroll.value(eBar)
   End If
End Sub

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'START INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
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
    IFileForm_Identify = FF_MAP
End Function

Private Property Get IFileForm_IsDirty() As Boolean
    IFileForm_IsDirty = m_IsDirty
End Property

Private Function IFileForm_Load(ByVal sFile As String) As Long
    Dim lResult As Long
    Dim Ext As String
    
    Screen.MousePointer = vbHourglass
    frmMain.setStatusMessage ("Loading map file")
    
    Ext = FSO.GetExtensionName(sFile)
    
    lResult = map.Load(sFile) 'Cargar el mapa
    
    If (lResult) Then
        PaintMap
        m_FilePath = sFile
        IsDirty = False
    Else
        MsgBox MSG_LOAD_ERRORLOADING + map.GetLastError, vbCritical
    End If
    
    Screen.MousePointer = 0
    frmMain.setStatusMessage
    
    IFileForm_Load = lResult
End Function

Private Function IFileForm_NewW(ByVal iUntitledCount As Integer) As Long
    Dim sFiles() As String
    Dim fileCount As Integer
    Dim lResult As Long
    
    m_Title = "Untitled map " & CStr(iUntitledCount)
            
    frmMain.setStatusMessage ("Converting file to map format...")
    Screen.MousePointer = vbHourglass
    
    fileCount = ShowOpenDialog(sFiles, getFilter("IMPORTABLE_GRAPHICS"), False, False)
    
    If fileCount > 0 Then
        lResult = map.Import(sFiles(0))
        If lResult <> -1 Then
            MsgBox "An error ocurred trying to import the file: " & map.GetLastError, vbCritical
        Else
            m_addToProject = modMenuActions.NewAddToProject
            PaintMap
            IsDirty = True
        End If
    End If
    
    Screen.MousePointer = 0
    frmMain.setStatusMessage
    
    IFileForm_NewW = lResult
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
    Dim lResult As Long

    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    lResult = map.Save(sFile) 'Save the map
    If (lResult) Then 'Save succesful
        'Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then addFileToProject sFile
    
        If m_FilePath <> sFile Then 'Show a success message only if the name changed
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        
        IsDirty = False
        m_FilePath = sFile
    Else
        MsgBox MSG_SAVE_ERRORSAVING + map.GetLastError, vbCritical
    End If
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'This property is not part of the interface, just helps to reduce code
'by setting the caption of the form properly
Private Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'INTERFACE IPROPERTIESFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
Private Function IPropertiesForm_GetProperties() As cProperties
    Dim cp As String
    Dim i As Integer
    Dim cpoint() As Integer
    
    Dim props As cProperties
    Set props = New cProperties
    
    With props
        .Add "Description", "Description", ptText, Me, "EditDescription", map.Description, True, 32
        '.Add "File", "File", ptText, Me, "", map.FilePath, False
        .Add "Width", "Width", ptNumeric, Me, "WidthP_Changed", map.Width, False
        .Add "Height", "Height", ptNumeric, Me, "HeightP_changed", map.Height, False
        .Add "Code", "Code", ptInteger, Me, "EditCode", map.Code, True, 9999, 0, False
        .Add "Depth", "Depth", ptCombo, Me, "EditDepth", (map.Depth \ 8) - 1, True
        .Add "Palette", "Palette", ptLink, Me, "", IIf(map.Depth = 8, "Palette 256 colors", "None"), False
        

        For i = 0 To map.CPointsCount() - 1
            cpoint() = map.ControlPoint(i)
            cp = cp & "(" & i & ": " & cpoint(0) & ", " & cpoint(1) & ");"
        Next
        
        .Add "C.Points", "CP", ptLink, Me, "EditCP", cp, True
    End With
  
    props("Depth").AddOption "8 bits"
    props("Depth").AddOption "16 bits"

    props("Description").Description = "Specifies the name of the map in the FPG"
    props("Code").Description = "Specifies the code of the map in the FPG"
    props("Width").Description = "Specifies the with of the MAP"
    props("Height").Description = "Specifies the height of the MAP"
    props("Palette").Description = "Specifies the palette to use with the MAP (only 8 bits) "
    props("Depth").Description = "Specifies number of bits per pixel of the MAP"
    props("CP").Description = "Set control points (including center)"

    Set IPropertiesForm_GetProperties = props
'    If map.Available Then
'    'Establece las propiedades
'    With frmProperties
'        .ClearProperties
'        .AddProperty "Description", "Description", ptText, Me, "EditDescription", map.Description, True, 32
'        .AddProperty "File", "File", ptText, Me, "", map.FilePath, False
'        .AddProperty "Width", "Width", ptNumeric, Me, "WidthP_Changed", map.Width, False
'        .AddProperty "Height", "Height", ptNumeric, Me, "HeightP_changed", map.Height, False
'        .AddProperty "Code", "Code", ptInteger, Me, "EditCode", map.code, True, 999, 0, False
'        .AddProperty "Depth", "Depth", ptCombo, Me, "EditDepth", (map.Depth \ 8) - 1, True
'        .AddProperty "Palette", "Palette", ptLink, Me, "", IIf(map.Depth = 8, "Palette 256 colors", "None"), False
'
'        For i = 0 To map.CPointsCount() - 1
'            cpoint() = map.ControlPoint(i)
'            cp = cp & "(" & i & ": " & cpoint(0) & ", " & cpoint(1) & ");"
'        Next
'
'        .AddProperty "C.Points", "CP", ptLink, Me, "EditCP", cp, True
'
'        .AddPropertyOption "Depth", "8 bits"
'        .AddPropertyOption "Depth", "16 bits"
'        .AddPropertyDescription "Description", "Specifies the name of the map in the FPG"
'        .AddPropertyDescription "Code", "Specifies the code of the map in the FPG"
'        .AddPropertyDescription "Width", "Specifies the with of the MAP"
'        .AddPropertyDescription "Height", "Specifies the height of the MAP"
'        .AddPropertyDescription "Palette", "Specifies the palette to use with the MAP (only 8 bits) "
'        .AddPropertyDescription "Depth", "Specifies number of bits per pixel of the MAP"
'        .AddPropertyDescription "CP", "Set control points (including center)"
'        .RefreshGrid
'    End With
'    End If
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IPROPERTIESFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'

Private Sub tbrMap_ButtonClick(ByVal lButton As Long)
    Select Case tbrMap.ButtonKey(lButton)
    Case "ZoomIn"
        ZoomMap 0.5
    Case "ZoomOut"
        ZoomMap -0.5
    Case "ZoomRestore"
        SizeIndex = 1
    Case "ToogleTrans"
        ToggleTransparency
        tbrMap.ButtonChecked("ToogleTrans") = ShowTransparent
    Case "EditCP"
        EditCP
    Case "EditPalette"
        EditPalette
    End Select
End Sub

Private Sub createFpgsMenu()
    Dim frm As Form
    Dim ff As IFileForm
    Dim lParentIndex As Long
    Dim sFile As String
    Dim i As Integer
    
    Set m_FpgsMenu = Nothing
    
    Set m_FpgsMenu = New cMenus
    m_FpgsMenu.DrawStyle = M_Style
    m_FpgsMenu.CreateFromNothing Me.hwnd
    
    lParentIndex = m_FpgsMenu.AddItem(0, Key:="FpgsMenu")
    
    'First add opened fpgs
    For Each frm In Forms
        If TypeOf frm Is frmFpg Then
            Set ff = frm
            'The key will be "OPENED|hwnd|ItemCount"
            m_FpgsMenu.AddItem lParentIndex, ff.Title, , , "OPENED|" & frm.hwnd & "|" & CStr(m_FpgsMenu.ItemCount)
        End If
    Next
    
    m_FpgsMenu.AddItem lParentIndex, "-"
    
    'Now add project fpgs
    If Not openedProject Is Nothing Then
        For i = 0 To openedProject.Files.count - 1
            sFile = openedProject.Files(i + 1)
            If LCase(FSO.GetExtensionName(sFile)) = "fpg" Then
                'The key will be "PROJECT|FileName(relative)|ItemCount"
                m_FpgsMenu.AddItem lParentIndex, sFile, , , "PROJECT|" & sFile & "|" & CStr(m_FpgsMenu.ItemCount)
            End If
        Next
    End If
End Sub

Private Sub tbrMap_DropDownPress(ByVal lButton As Long)
    Dim X As Long, Y As Long
    Dim lIndex As Long
    tbrMap.GetDropDownPosition lButton, X, Y
    
    Select Case tbrMap.ButtonKey(lButton)
        Case "AddToFpg":
            createFpgsMenu
            Call m_FpgsMenu.PopupMenu("FpgsMenu", _
                Me.ScaleX(Me.Left + X, vbTwips, vbPixels), Me.ScaleY(Y, vbTwips, vbPixels) + Rebar.RebarHeight * 1.5, TPM_VERNEGANIMATION + TPM_LEFTALIGN)
    End Select
End Sub
