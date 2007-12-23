VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Begin VB.Form frmExtensions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   Icon            =   "frmExtensions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame grbPlugins 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   8760
      TabIndex        =   18
      Top             =   840
      Width           =   4935
   End
   Begin VB.Frame grbTools 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   7215
      Begin VB.CheckBox chkUseToolForFileAssoc 
         Caption         =   "&Use this tool to associate file extensions"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CommandButton cmdExploreTool 
         Caption         =   "..."
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   2190
         Width           =   375
      End
      Begin VB.TextBox txtCommand 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox txtParams 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   4575
      End
      Begin VB.CommandButton cmdRemoveTool 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddTool 
         Caption         =   "&Add"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.ListBox lstTools 
         Height          =   1425
         ItemData        =   "frmExtensions.frx":038A
         Left            =   120
         List            =   "frmExtensions.frx":038C
         TabIndex        =   0
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label6 
         Caption         =   "TIP: Use $(FILE_PATH) to indicate the name of the active file"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   6495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Command:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2220
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Params:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2700
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Title:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1740
         Width           =   345
      End
   End
   Begin VB.Frame grbAssoc 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   7920
      TabIndex        =   12
      Top             =   3720
      Width           =   5295
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
      End
      Begin vbAcceleratorSGrid6.vbalGrid grdAssoc 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
      End
   End
   Begin vbalDTab6.vbalDTabControl tabCategories 
      Height          =   4095
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7223
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExtensions.frx":038E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   15
      Top             =   240
      Width           =   6165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   -720
      Picture         =   "frmExtensions.frx":0426
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Property Get ToolIndex() As Integer
    If lstTools.ListIndex > -1 Then
        ToolIndex = lstTools.ItemData(lstTools.ListIndex)
    Else
        ToolIndex = -1
    End If
End Property

Private Sub FillExternalToolsList()
    Dim i As Integer, enableCtl As Boolean
    
    lstTools.Clear
    If ExternalToolsCount > 0 Then
        For i = 0 To ExternalToolsCount - 1
            lstTools.AddItem (ExternalTools(i).Title)
            lstTools.ItemData(lstTools.ListCount - 1) = i
        Next
    End If
    
    If ExternalToolsCount > 0 Then enableCtl = True
    txtTitle.Enabled = enableCtl
    txtCommand.Enabled = enableCtl
    txtParams.Enabled = enableCtl
    cmdExploreTool.Enabled = enableCtl
    cmdRemoveTool.Enabled = enableCtl
    If enableCtl = False Then
        txtTitle.text = ""
        txtCommand.text = ""
        txtParams.text = ""
    End If
End Sub

Private Sub ConfigureControls()
    Dim nTab As cTab
    
    'Controls position
    Width = 7400
    Height = 5800
    tabCategories.Move 0, 840, 7300, 4100
    
    
    'Configure tab control
    With tabCategories
        'Set nTab = .Tabs.Add(, , "File association")
        'nTab.Panel = grbAssoc
        Set nTab = .Tabs.Add(, , "External Tools")
        nTab.Panel = grbTools
'        Set nTab = .Tabs.Add(, , "Plugins")
'        nTab.Panel = grbPlugins
    End With
        
    'Extension assoc
    With grdAssoc
        .HeaderButtons = False
        .HeaderFlat = True
        .Redraw = False
        .SelectionAlphaBlend = True
        .SelectionOutline = False
        .RowMode = True
        .BorderStyle = ecgBorderStyle3dThin
        .AddColumn "ext", "Extension", ecgHdrTextALignCentre, , 100
        .AddColumn "type", "File type", ecgHdrTextALignLeft, , 150
        .AddColumn "tool", "Open with", ecgHdrTextALignLeft, , 200
        .StretchLastColumnToFit = True
        
        'Default extensions
        .AddRow
        .cell(1, 1).text = "txt"
        .cell(1, 1).TextAlign = DT_CENTER
        .cell(1, 2).text = "PLAIN TEXT FILE"
        .cell(1, 3).text = "Source code editor"
        
        .AddRow
        .cell(2, 1).text = "prg"
        .cell(2, 1).TextAlign = DT_CENTER
        .cell(2, 2).text = "PLAIN TEXT FILE"
        .cell(2, 3).text = "Source code editor"
        .AddRow
        .cell(3, 1).text = "inc"
        .cell(3, 1).TextAlign = DT_CENTER
        .cell(3, 2).text = "PLAIN TEXT FILE"
        .cell(3, 3).text = "Source code editor"
        .AddRow
        .cell(4, 1).text = "map"
        .cell(4, 1).TextAlign = DT_CENTER
        .cell(4, 2).text = "FENIX PAL FILE"
        .cell(4, 3).text = "Map editor"
        
        .AddRow , -1
        
        .ColumnAlign("ext") = ecgHdrTextALignCentre
        
        .Redraw = True
    End With
End Sub


Private Sub chkUseToolForFileAssoc_Click()
    Dim tool As T_ExternalTool
    
    If ToolIndex <> -1 Then
        tool = ExternalTools(ToolIndex)
        tool.UseForFileAssoc = IIf(chkUseToolForFileAssoc.value = 0, False, True)
        ExternalTools(ToolIndex) = tool
    End If
End Sub

Private Sub cmdAddTool_Click()
    Dim tool As T_ExternalTool
    
    tool.Title = "New tool"
    AddExternalTool tool
    FillExternalToolsList
    
    lstTools.ListIndex = lstTools.ListCount - 1
    txtTitle.SetFocus
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExploreTool_Click()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles, , True, False) > 0 Then
        txtCommand.text = sFiles(0)
    End If
End Sub

Private Sub cmdOk_Click()
    SaveExternalTools
    
    MsgBox "Changes will be applied next time you run Flamebird (Yes, this needs to be changed...)", vbInformation
    
    Unload Me
End Sub


Private Sub cmdRemoveTool_Click()
    If ToolIndex <> -1 Then
        RemoveExternalTool ToolIndex
        FillExternalToolsList
        If lstTools.ListCount > 0 Then lstTools.ListIndex = lstTools.ListCount - 1
    End If
End Sub

Private Sub Form_Load()
    ConfigureControls
    
    LoadExternalTools
    FillExternalToolsList
    
    If lstTools.ListCount > 0 Then lstTools.ListIndex = 0
End Sub

Private Sub lstTools_Click()
    If lstTools.SelCount > 0 Then
        If ToolIndex > -1 Then
            With ExternalTools(ToolIndex)
                txtTitle.text = .Title
                txtCommand.text = .Command
                txtParams.text = .Params
                chkUseToolForFileAssoc.value = IIf(.UseForFileAssoc, 1, 0)
            End With
        End If
    End If
End Sub

Private Sub txtCommand_Change()
    Dim tool As T_ExternalTool
    
    If ToolIndex <> -1 Then
        tool = ExternalTools(ToolIndex)
        tool.Command = txtCommand.text
        ExternalTools(ToolIndex) = tool
    End If
End Sub

Private Sub txtParams_Change()
    Dim tool As T_ExternalTool
    
    If ToolIndex <> -1 Then
        tool = ExternalTools(ToolIndex)
        tool.Params = txtParams.text
        ExternalTools(ToolIndex) = tool
    End If
End Sub

Private Sub txtTitle_Change()
    Dim tool As T_ExternalTool
    
    If ToolIndex <> -1 Then
        tool = ExternalTools(ToolIndex)
        tool.Title = txtTitle.text
        ExternalTools(ToolIndex) = tool
        lstTools.List(lstTools.ListIndex) = txtTitle.text
    End If
End Sub
