VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Begin VB.Form frmAddToProject 
   Caption         =   "Open"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6990
   Icon            =   "Frm_new.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3360
      Top             =   4080
   End
   Begin VB.PictureBox m_Controls 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   -1200
      ScaleHeight     =   3855
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdNewCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewOpen 
         Caption         =   "&Open"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   2640
         Width           =   1095
      End
      Begin vbalDTab6.vbalDTabControl tabMain 
         Height          =   3855
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6800
         TabAlign        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowCloseButton =   0   'False
      End
   End
End
Attribute VB_Name = "frmAddToProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Public WithEvents cD As cCommonDialog
'
'Private ShiftDlgX As Long
'Private ShiftDlgY As Long
'
''Holds Dlg Toolbarhwnd
'Private Tbhwnd  As Long
''Tmp String for when we change tabs and then return to existing tab
'Private TmpFileTxt      As String * 256
''Tab to clean up before showing new tab
'Private LastTabFocus    As String
'
'Private m_hWnd As Long
'Private m_hDlg As Long
'Private m_bCancel As Boolean
'Private m_sFileName As String
'
'Private m_bShowNew As Boolean
'Private m_bIsNew As Boolean
'
'Private m_ilsNew As vbalImageList
'Private m_ilsExist As vbalImageList
'Private m_sNewName() As String
'Private m_vNewIcon() As Variant
'Private m_iNewCount As Long
'Private m_sExistName() As String
'Private m_sExistFolder() As String
'Private m_dExistDate() As Date
'Private m_vExistIcon() As Variant
'Private m_iExistCOunt As Long
'
'Public Function SetBackwardFocus() As Long
'Dim CurHnd As Long
'Dim NewHnd As Long
'
'   CurHnd = GetFocusAPI()
'
'   If tabMain.SelectedTab.Key = "New" Then
'
'      Select Case CurHnd
'      Case lvwNew.hwnd
'         SetFocusAPI cmdNewCancel.hwnd
'      Case cmdNewOpen.hwnd
'         SetFocusAPI lvwNew.hwnd
'      Case cmdNewCancel.hwnd
'          SetFocusAPI cmdNewOpen.hwnd
'      Case Else
'         SetFocusAPI lvwNew.hwnd
'      End Select
'
'      NewHnd = 1
'
'   Else
'
'   End If
'
'   If NewHnd <> 0 Then
'      SetBackwardFocus = 1
'   Else
'      SetBackwardFocus = 0
'   End If
'
'End Function
'
'Private Sub SetControlPos(ByVal hd As Long, ByVal SFlags As Long, Optional ByVal lID As Long = -1)
'Dim tdr As RECT
'Dim wLp As POINTAPI
'Dim lShiftX As Long
'Dim lShiftY As Long
'
'   GetWindowRect hd, tdr
'   ScreenToClient m_hWnd, wLp
'
'   lShiftX = tabMain.ClientLeft \ Screen.TwipsPerPixelX
'   'lShiftY = tabMain.ClientTop \ Screen.TwipsPerPixelY + 2
'   lShiftY = 375 \ Screen.TwipsPerPixelY
'
'   wLp.X = wLp.X + tdr.Left + lShiftX
'   wLp.Y = wLp.Y + tdr.Top + lShiftY
'
'   SetWindowPos hd, 0, wLp.X, wLp.Y, 0, 0, SFlags
'
'   If lID = fdlgIDCANCEL Then
'      ' Position VB Cancel:
'      cmdNewCancel.Move wLp.X * Screen.TwipsPerPixelX - m_Controls.Left, wLp.Y * Screen.TwipsPerPixelY - 375 - m_Controls.Top, (tdr.Right - tdr.Left) * Screen.TwipsPerPixelX, (tdr.Bottom - tdr.Top) * Screen.TwipsPerPixelY
'   ElseIf lID = fdlgIDOK Then
'      ' Position VB OK:
'      cmdNewOpen.Move wLp.X * Screen.TwipsPerPixelX - m_Controls.Left, wLp.Y * (Screen.TwipsPerPixelY) - 375 - m_Controls.Top, (tdr.Right - tdr.Left) * Screen.TwipsPerPixelX, (tdr.Bottom - tdr.Top) * Screen.TwipsPerPixelY
'      ' Position ListViews:
'      lvwNew.Move lvwNew.Left, lvwNew.Top, _
'         cmdNewOpen.Left + cmdNewOpen.Width - lvwNew.Left, _
'         cmdNewOpen.Top - lvwNew.Top - 4 * Screen.TwipsPerPixelY
'   End If
'
'End Sub
'
'Public Function SetForwardFocus() As Long
'Dim CurHnd As Long
'Dim NewHnd As Long
'
'   CurHnd = GetFocusAPI()
'
'   If tabMain.SelectedTab.Key = "New" Then
'
'      Select Case CurHnd
'      Case lvwNew.hwnd
'         SetFocusAPI cmdNewOpen.hwnd
'      Case cmdNewOpen.hwnd
'         SetFocusAPI cmdNewCancel.hwnd
'      Case cmdNewCancel.hwnd
'         SetFocusAPI lvwNew.hwnd
'      Case Else
'         SetFocusAPI lvwNew.hwnd
'      End Select
'
'      NewHnd = 1
'
'   Else
'
'   End If
'
'   If NewHnd <> 0 Then
'      SetForwardFocus = 1
'   Else
'      SetForwardFocus = 0
'   End If
'
'End Function
'
'
'Private Sub cD_DialogOK(bCancel As Boolean)
'
'   bCancel = True
'   m_bCancel = False
'   m_sFileName = GetCDlgFileName(m_hDlg)
'
'   tmrUnload.Enabled = True
'
'End Sub
'
'Private Sub cD_FolderChange(ByVal hDlg As Long)
'Static DoOnce As Boolean
'Dim hd As Long
'
'   If Not DoOnce Then
'      'We do this because the file listvw was not_
'      'created till after dialog intialize
'
'      If m_bShowNew Then
'         hd = GetDlgItem(m_hWnd, fdlgLVLstFiles)
'         SetControlPos hd, SWP_MOVEHIDE Or SWP_NOMOVE Or SWP_NOZORDER
'      Else
'         SetFocusAPI m_Controls.hwnd
'      End If
'
'      m_Controls.Visible = True
'      m_Controls.ZOrder 1
'      DoOnce = True
'   End If
'End Sub
'
'
'Private Sub cD_InitDialog(ByVal hDlg As Long)
'Dim hd As Long
'Dim tr As RECT, tTR As RECT
'Dim lW As Long, lH As Long
'
'   m_hDlg = hDlg
'   ' For Open/Save dialog we need the parent of the supplied dialog handle:
'   m_hWnd = GetParent(hDlg)
'   GetWindowRect m_hWnd, tTR
'
'   If m_bShowNew Then
'
'      '// Cancel Button
'      hd = GetDlgItem(m_hWnd, fdlgIDCANCEL)
'      SetControlPos hd, SWP_MOVEHIDE, fdlgIDCANCEL
'      GetWindowRect hd, tr
'
'      '// Open Button
'      hd = GetDlgItem(m_hWnd, fdlgIDOK)
'      SetControlPos hd, SWP_MOVEHIDE, fdlgIDOK
'
'      '// Read-Only CheckBox
'      hd = GetDlgItem(m_hWnd, fdlgChxReadOnly)
'      SetControlPos hd, SWP_MOVEHIDE
'      EnableWindow hd, False
'
'      '// FileType TextBox
'      hd = GetDlgItem(m_hWnd, fdlgcmbSaveAsType)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// FileType Label
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveAsType)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// FileName TxtBox
'      hd = GetDlgItem(m_hWnd, fdlgEdtFileName)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// FileName Label
'      hd = GetDlgItem(m_hWnd, fdlgStcFileName)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// ListBoxLb
'      hd = GetDlgItem(m_hWnd, fdlgLBLstFiles)
'      SetControlPos hd, SWP_MOVEHIDE
'      EnableWindow hd, True
'
'      '// Find ComboBox:
'      hd = GetDlgItem(m_hWnd, fdlgCmbSaveInFindIn)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// Find In Label:
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveInFindIn)
'      SetControlPos hd, SWP_MOVEHIDE
'
'      '// Tool Bar
'      Tbhwnd = FindWindowEx(m_hWnd, 0&, "ToolbarWindow32", vbNullString)
'      SetControlPos Tbhwnd, SWP_MOVEHIDE Or SWP_NOZORDER
'
'   Else
'
'      '// Cancel Button
'      hd = GetDlgItem(m_hWnd, fdlgIDCANCEL)
'      SetControlPos hd, SWP_NOSIZE
'      GetWindowRect hd, tr
'
'      '// Open Button
'      hd = GetDlgItem(m_hWnd, fdlgIDOK)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// Read-Only CheckBox
'      hd = GetDlgItem(m_hWnd, fdlgChxReadOnly)
'      SetControlPos hd, SWP_MOVEHIDE
'      EnableWindow hd, False
'
'      '// FileType TextBox
'      hd = GetDlgItem(m_hWnd, fdlgcmbSaveAsType)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// FileType Label
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveAsType)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// FileName TxtBox
'      hd = GetDlgItem(m_hWnd, fdlgEdtFileName)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// FileName Label
'      hd = GetDlgItem(m_hWnd, fdlgStcFileName)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// ListBoxLb
'      hd = GetDlgItem(m_hWnd, fdlgLBLstFiles)
'      SetControlPos hd, SWP_NOSIZE
'      EnableWindow hd, True
'
'      '// Find ComboBox:
'      hd = GetDlgItem(m_hWnd, fdlgCmbSaveInFindIn)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// Find In Label:
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveInFindIn)
'      SetControlPos hd, SWP_NOSIZE
'
'      '// Tool Bar
'      Tbhwnd = FindWindowEx(m_hWnd, 0&, "ToolbarWindow32", vbNullString)
'      SetControlPos Tbhwnd, SWP_NOSIZE Or SWP_NOZORDER
'
'   End If
'
'   'Width and Height needed by the Dialog (in pixels)
'    lW = tTR.Right - tTR.Left
'    lH = tTR.Bottom - tTR.Top
'
'    MoveWindow m_hWnd, 0&, 0&, lW, lH, 0 'Dialog
'    MoveWindow m_Controls.hwnd, 0&, 0&, lW, lH - 25, 0
'    'm_controls picturebox
'    tabMain.Width = (lW - 7) * Screen.TwipsPerPixelX 'tab
'    tabMain.Height = m_Controls.Height
'
'   'Center it
'   cD.CentreDialog hDlg, Screen
'
'   'move our controls container to dialog
'   SetParent m_Controls.hwnd, m_hWnd
'
'   InstallHook
'
'End Sub
'Private Sub chkDont_Click()
'   m_bShowNew = Not m_bShowNew
'End Sub
'
'Private Sub CmdNewCancel_Click()
'   Form_QueryUnload 0, 0
'End Sub
'
'Private Sub cmdNewOpen_Click()
'   If lvwNew.SelectedItem Is Nothing Then Exit Sub
'
'   m_sFileName = lvwNew.SelectedItem.text
'   m_bCancel = False
'   Form_QueryUnload 0, 0
'
'End Sub
'
'Private Sub Form_Load()
'    Dim nTab As cTab
'    With tabMain
'        .ImageList = 0
'        Set nTab = .Tabs.Add("NEW", , "New")
'        'nTab.Panel = picNew
'        Set nTab = .Tabs.Add("EXISTING", , "Existing")
'    End With
'
'    Dim itmX As vbalListViewLib6.cListItem
'    Dim i As Long
'
'        lvwNew.Top = lvwNew.Top + 375
'
'        ' Signal default of cancelled:
'        m_bCancel = True
'
'        ' Form initialisation:
'        If Not (m_ilsNew Is Nothing) Then
'            lvwNew.ImageList = m_ilsNew
'        End If
'
'
'        If m_bShowNew Then
'         For i = 1 To m_iNewCount
'            Set itmX = lvwNew.ListItems.Add(, , m_sNewName(i), m_vNewIcon(i))
'            If (i = 1) Then
'                itmX.Selected = True
'                m_sFileName = lvwNew.SelectedItem.text
'            End If
'         Next i
'        End If
'
'
'       If Not (m_bShowNew) Then
'          lvwNew.Visible = False
'          cmdNewOpen.Visible = False
'          cmdNewCancel.Visible = False
'          tabMain.Tabs.Remove ("NEW")
'       Else
'          m_bIsNew = True
'          lvwNew.Visible = True
'          cmdNewOpen.Visible = True
'          cmdNewCancel.Visible = True
'          LastTabFocus = "NEW"
'       End If
'
'End Sub
'
'
'Private Sub lvwNew_ItemClick(Item As vbalListViewLib6.cListItem)
'm_sFileName = Item.text
'End Sub
'
'Private Sub lvwNew_ItemDblClick(Item As vbalListViewLib6.cListItem)
'   If Not (lvwNew.SelectedItem Is Nothing) Then
'      cmdNewOpen.value = True
'   End If
'End Sub
'
'Private Sub tabMain_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
'
'Dim iTab As String
'Dim hd  As Long
'
'   ' Show the relevant picture box for the
'   ' selected tab:
'   iTab = theTab.Key
'   If iTab = LastTabFocus Then Exit Sub
'
'   Select Case LastTabFocus
'   Case "NEW"
'      m_bIsNew = False
'      ShowWindow lvwNew.hwnd, 0&
'
'      cmdNewOpen.Visible = False
'      cmdNewCancel.Visible = False
'
'   Case "EXISTING"
'
'      Call GetDlgItemText(m_hWnd, fdlgEdtFileName, TmpFileTxt, 128)
'
'      hd = GetDlgItem(m_hWnd, fdlgLVLstFiles)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveInFindIn)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgCmbSaveInFindIn)
'      ShowWindow hd, 0&
'      EnableWindow hd, 0&
'
'      ShowWindow Tbhwnd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgStcFileName)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgEdtFileName)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveAsType)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgcmbSaveAsType)
'      ShowWindow hd, 0&
'
'
'   End Select
'
'   Select Case iTab
'   Case "NEW"
'      m_bIsNew = True
'      ShowWindow lvwNew.hwnd, 1&
'      'SetFocusAPI tabMain.hWnd
'
'      hd = GetDlgItem(m_hWnd, fdlgIDCANCEL)
'      ShowWindow hd, 0&
'
'      hd = GetDlgItem(m_hWnd, fdlgIDOK)
'      ShowWindow hd, 0&
'
'      cmdNewOpen.Visible = True
'      cmdNewCancel.Visible = True
'
'      If Not (lvwNew.SelectedItem Is Nothing) Then
'         m_sFileName = lvwNew.SelectedItem.text
'         lvwNew.SelectedItem.Selected = True
'      Else
'         m_sFileName = ""
'      End If
'
'      'SetFocusAPI tabMain.hWnd
'
'   Case "EXISTING"
'      hd = GetDlgItem(m_hWnd, fdlgIDCANCEL)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgIDOK)
'      ShowWindow hd, 1&
'      EnableWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgcmbSaveAsType)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveAsType)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgEdtFileName)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgStcFileName)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgLVLstFiles)
'      ShowWindow hd, 1&
'
'      hd = GetDlgItem(m_hWnd, fdlgCmbSaveInFindIn)
'      ShowWindow hd, 1&
'      EnableWindow hd, True
'
'      hd = GetDlgItem(m_hWnd, fdlgStcSaveInFindIn)
'      ShowWindow hd, 1&
'
'      ShowWindow Tbhwnd, 1
'
'      m_sFileName = TmpFileTxt
'
'      'SetFocusAPI tabMain.hWnd
'      SetDlgItemText m_hWnd, fdlgEdtFileName, TmpFileTxt
'
'   End Select
'
'   LastTabFocus = tabMain.SelectedTab.Key
'End Sub
'
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If (m_hWnd <> 0) And (m_hDlg <> 0) Then
'      ' Cancel Common Dialog box if loaded:
'
'      'Put m_Controls back on Me
'      SetParent m_Controls.hwnd, Me.hwnd
'
'      ' Send Close command to the dialog:
'      SendMessageLong m_hWnd, WM_COMMAND, fdlgIDCANCEL, 1
'      SendMessageLong m_hDlg, WM_CLOSE, 0, 0
'      SendMessageLong m_hWnd, WM_CLOSE, 0, 0
'      m_hWnd = 0: m_hDlg = 0
'   End If
'
'   RemoveHook
'
'End Sub
'
'Private Sub Form_Resize()
'   Debug.Print
'End Sub
'
'Public Property Get ShowNew() As Boolean
'    ShowNew = m_bShowNew
'End Property
'Public Property Get IsNew() As Boolean
'    IsNew = m_bIsNew
'End Property
'Public Sub AddNewType(ByVal sName As String, Optional ByVal vIcon As Variant)
'    m_iNewCount = m_iNewCount + 1
'    ReDim Preserve m_sNewName(1 To m_iNewCount) As String
'    ReDim Preserve m_vNewIcon(1 To m_iNewCount) As Variant
'    m_sNewName(m_iNewCount) = sName
'    m_vNewIcon(m_iNewCount) = vIcon
'End Sub
'Public Sub AddExistItem(ByVal sName As String, ByVal sFOlder As String, ByVal dDate As Date, Optional ByVal vIcon As Variant)
'    m_iExistCOunt = m_iExistCOunt + 1
'    ReDim Preserve m_sExistName(1 To m_iExistCOunt) As String
'    ReDim Preserve m_sExistFolder(1 To m_iExistCOunt) As String
'    ReDim Preserve m_dExistDate(1 To m_iExistCOunt) As Date
'    ReDim Preserve m_vExistIcon(1 To m_iExistCOunt) As Variant
'    m_sExistName(m_iExistCOunt) = sName
'    m_sExistFolder(m_iExistCOunt) = sFOlder
'    m_dExistDate(m_iExistCOunt) = dDate
'    m_vExistIcon(m_iExistCOunt) = vIcon
'End Sub
'Public Property Let NewImageList(ByRef ilsThis As vbalImageList)
'    Set m_ilsNew = ilsThis
'End Property
'Public Property Let ExistingImageList(ByRef ilsThis As vbalImageList)
'    Set m_ilsExist = ilsThis
'End Property
'Public Property Let ShowNew(ByVal bShow As Boolean)
'    m_bShowNew = bShow
'End Property
'Public Property Get Cancelled() As Boolean
'    Cancelled = m_bCancel
'End Property
'Public Property Get filename() As String
'    ' Property to allow the caller to retrieve the
'    ' selected file name:
'    filename = m_sFileName
'End Property
'
'
'
'Private Sub tmrUnload_Timer()
'   Unload Me
'End Sub
'
Private Sub Form_Load()

End Sub
