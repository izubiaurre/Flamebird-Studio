VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmProjectProperties 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project properties"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjectProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   17505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExeFile 
      Caption         =   "..."
      Height          =   255
      Left            =   12360
      TabIndex        =   26
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtIExeFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00666666&
      BorderStyle     =   0  'None
      ForeColor       =   &H00BDBDBD&
      Height          =   285
      Left            =   7800
      TabIndex        =   25
      Top             =   3240
      Width           =   4455
   End
   Begin VB.PictureBox pic_SelIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   7680
      ScaleHeight     =   3615
      ScaleWidth      =   5175
      TabIndex        =   20
      Top             =   720
      Width           =   5175
      Begin VB.CommandButton cmdApplyIcon 
         Caption         =   "Apply"
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdIconFile 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtIconFile 
         BackColor       =   &H00666666&
         BorderStyle     =   0  'None
         ForeColor       =   &H00BDBDBD&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Exe file:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Icon file:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.PictureBox pic_compilation 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   5655
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox txtParameters 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Frame frmFenixPath 
         Caption         =   "Fenix path"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   5295
         Begin VB.CommandButton cmdFenixPath 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtFenixPath 
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
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   4695
         End
      End
      Begin VB.CheckBox chkEspecificFenix 
         Caption         =   "Compile with a especific version of Fenix."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton cmdMainSource 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtMainPRG 
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CommandButton cmdCompilationDir 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtCompilationDir 
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Command line arguments:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Initial PRG:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Compilation directory:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   4320
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox pic_general 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Project name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin vbalDTab6.vbalDTabControl TabControl 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6800
      AllowScroll     =   0   'False
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current project properties."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   19
      Top             =   240
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmProjectProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filename As String
Dim WithEvents browseDir As cBrowseForFolder
Attribute browseDir.VB_VarHelpID = -1
Dim cdialog As cCommonDialog
Attribute cdialog.VB_VarHelpID = -1

Private Sub BrowseFolders(txt As TextBox)
    Dim s As String
    With browseDir
        .hwndOwner = Me.Hwnd
        .InitialDir = App.Path
        .FileSystemOnly = True
        .StatusText = True
        .UseNewUI = True
        If Not openedProject Is Nothing Then
            s = openedProject.makePathRelative(.BrowseForFolder)
        Else
            s = .BrowseForFolder
        End If
        If Len(s) > 0 Then txt.text = s
    End With
End Sub

Private Sub chkEspecificFenix_Click()
    frmFenixPath.Enabled = CBool(chkEspecificFenix.Value)
End Sub

Private Sub cmdApplyIcon_Click()
    Dim errorN As Long
    erroN = InsertIcons(txtIconFile, txtIExeFile)
    
    If erroN > 0 Then ShowError ("ApplyIcon"): Resume Next
    
End Sub

Private Sub cmdCompilationDir_Click()
    BrowseFolders txtCompilationDir
End Sub

Private Sub cmdExeFile_Click()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles(), getFilter("EXE"), True, False) > 0 Then
        txtIExeFile = openedProject.makePathRelative(sFiles(0))
    End If
End Sub

Private Sub cmdIconFile_Click()
    On Error GoTo ErrHandler:
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles(), getFilter("ICON"), True, False) > 0 Then
        txtIconFile = openedProject.makePathRelative(sFiles(0))
        imgIcon.Picture = LoadPicture(txtIconFile)
    End If
    Exit Sub
ErrHandler:
    MsgBox "Could not display icon properly", vbCritical
End Sub

Private Sub cmdMainSource_Click()
    Dim sFiles() As String
    
    If ShowOpenDialog(sFiles, getFilter("SOURCE"), True, False) > 0 Then
        txtMainPRG = openedProject.makePathRelative(sFiles(0))
    End If
End Sub

Private Sub cmdFenixPath_Click()
   BrowseFolders txtFenixPath
End Sub

Private Sub cmdOk_Click()
    If txtName.text = "" Then
        MsgBox "The field 'Project Name' cannot be empty", vbCritical
        If TabControl.Tabs.item(1).Selected = False Then ' go to errors tab
            TabControl.Tabs.item(1).Selected = True
        End If
        txtName.SetFocus
        Exit Sub
    End If
    SaveConf
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Set browseDir = New cBrowseForFolder
    Set cdialog = New cCommonDialog
    
    Image1.Picture = LoadPicture(App.Path & "/Resources/frmHeader.jpg")
    
    Dim nTab As cTab
    With TabControl
        .ImageList = 0
        Set nTab = .Tabs.Add("GENERAL", , "General")
        nTab.Panel = pic_general
        Set nTab = .Tabs.Add("COMPILATION", , "Compilation")
        nTab.Panel = pic_compilation
        Set nTab = .Tabs.Add("ICON", , "Icon")
        nTab.Panel = pic_SelIcon
    End With
End Sub
Public Sub SaveConf()
    With openedProject
        .projectName = txtName.text
        .compilationDir = txtCompilationDir.text
        .useOtherFenix = CBool(chkEspecificFenix.Value)
        .fenixDir = txtFenixPath.text
        .compilerArguments = txtParameters.text
        .mainSource = txtMainPRG.text
    End With
End Sub
Public Sub LoadConf()
    With openedProject
        txtName.text = .projectName
        txtCompilationDir.text = .compilationDir
        chkEspecificFenix.Value = Abs(CInt(.useOtherFenix))
        txtFenixPath.text = .fenixDir
        txtParameters.text = .compilerArguments
        txtMainPRG.text = .mainSource
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set browseDir = Nothing
End Sub

