VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbaltreeview6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Object = "{C8A61D56-D8DC-11D2-8064-9D6F06504DA8}#1.1#0"; "axcolctl.ocx"
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   9960
   ClientLeft      =   3150
   ClientTop       =   1005
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   5640
      ScaleHeight     =   4095
      ScaleWidth      =   5535
      TabIndex        =   54
      Top             =   5280
      Width           =   5535
      Begin VB.CheckBox chkConfine 
         Caption         =   "Confine caret to text"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Frame grbCode 
         BorderStyle     =   0  'None
         Caption         =   "Code"
         Height          =   3135
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   5535
         Begin VB.TextBox txtTabSize 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4440
            TabIndex        =   66
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox chkSmoothScrolling 
            Caption         =   "Smooth scrolling"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CheckBox chkWhiteSpaces 
            Caption         =   "Display white spaces"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   2250
            Width           =   1935
         End
         Begin VB.Frame grbAutoIdent 
            Caption         =   "Auto indent mode"
            Height          =   1575
            Left            =   3360
            TabIndex        =   60
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton opIndentNone 
               Caption         =   "None"
               Height          =   195
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton opIndentScope 
               Caption         =   "Scope"
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   1200
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton opIndentPrevLine 
               Caption         =   "Previous line"
               Height          =   195
               Left            =   240
               TabIndex        =   61
               Top             =   780
               Width           =   1215
            End
         End
         Begin VB.CheckBox chkNormalizeCase 
            Caption         =   "Normalize keyword case"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   1755
            Width           =   2175
         End
         Begin VB.CheckBox chkColorSintax 
            Caption         =   "Color syntax"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1245
            Width           =   1335
         End
         Begin VB.CheckBox chkBookmarkMargin 
            Caption         =   "Display bookmark margin"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   750
            Width           =   2175
         End
         Begin VB.CheckBox chkLineNumbering 
            Caption         =   "Display line number margin"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tab size:"
            Height          =   195
            Left            =   3600
            TabIndex        =   67
            Top             =   2205
            Width           =   645
         End
      End
   End
   Begin VB.PictureBox picFileAsoc 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   5295
      TabIndex        =   34
      Top             =   2520
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox chkAskReg 
         Caption         =   "Ask for File Association on init"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Frame fraFiletypes 
         Height          =   3015
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   5175
         Begin VB.CheckBox chkDcb 
            Caption         =   "Open DCB files with Fenix Interpreter"
            Height          =   375
            Left            =   600
            TabIndex        =   36
            Top             =   2520
            Width           =   3015
         End
         Begin vbalTreeViewLib6.vbalTreeView trFiles 
            Height          =   1815
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3201
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the file types that you want to open with FlameBird."
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
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Label lblNotice 
         Caption         =   "Note: This will apply only when at least one filetype isn't registered."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   5175
      End
   End
   Begin vbalDTab6.vbalDTabControl tabCategories 
      Height          =   480
      Left            =   5640
      TabIndex        =   16
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   847
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
   Begin VB.PictureBox picUserTools 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   9960
      ScaleHeight     =   3975
      ScaleWidth      =   5295
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame fraToolData 
         Height          =   2295
         Left            =   0
         TabIndex        =   24
         Top             =   1680
         Width           =   5175
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   720
            TabIndex        =   30
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   720
            TabIndex        =   29
            Top             =   720
            Width           =   3735
         End
         Begin VB.CommandButton cmdAddTool 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3840
            TabIndex        =   28
            ToolTipText     =   "Add new tool"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "C&lear"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdToolExplore 
            Caption         =   "..."
            Height          =   315
            Left            =   4560
            TabIndex        =   26
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtParms 
            Height          =   285
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   25
            ToolTipText     =   "Insert here any command-line parameter you want to pass to the app"
            Top             =   1440
            Width           =   3855
         End
         Begin VB.Label lblName 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblPath 
            Caption         =   "Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblParms 
            Caption         =   "Command-line parameters:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1935
         End
      End
      Begin VB.Frame fraTools 
         Height          =   1575
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   5175
         Begin VB.ListBox lstUserTools 
            Height          =   1230
            ItemData        =   "frmPreferences.frx":0000
            Left            =   120
            List            =   "frmPreferences.frx":0002
            TabIndex        =   23
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdRemoveTool 
            Caption         =   "R&emove"
            Height          =   375
            Left            =   3840
            TabIndex        =   22
            ToolTipText     =   "Remove selected tool"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remove all"
            Height          =   375
            Left            =   3840
            TabIndex        =   21
            ToolTipText     =   "Remove all tools"
            Top             =   720
            Width           =   1215
         End
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
      Left            =   9960
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   8760
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox picCompilation 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   6720
      Width           =   5295
      Begin VB.CheckBox chkDoubleBuf 
         Caption         =   "Enable double buffering"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox chkSaveFiles 
         Caption         =   "Save modified files before compiling"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   2895
      End
      Begin VB.CheckBox chkFiltering 
         Caption         =   "Enable filtering (16 bits mode only)"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Compile in Debug mode"
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
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdExplore 
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
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   495
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
         TabIndex        =   6
         Top             =   360
         Width           =   4455
      End
      Begin VB.Frame grbSaveBeforeCompiling 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   855
         Left            =   240
         TabIndex        =   68
         Top             =   2160
         Width           =   3255
         Begin VB.OptionButton opAllOpenedFiles 
            Caption         =   "All opened files"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton opProjectFiles 
            Caption         =   "Project files"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton opCurrentFile 
            Caption         =   "Current file only"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Label lblCompiler 
         Caption         =   "Compilation && execution:"
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
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblFenixPath 
         Caption         =   "Fenix path:"
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
         TabIndex        =   11
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox picAppearance 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      Begin VB.CheckBox chkEnableXP 
         Caption         =   "Enable XP Look"
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
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Self explanatory don't you think?"
         Top             =   120
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkBitmap 
         Caption         =   "Show toolbar backgrounds"
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
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Show a bitmap texture"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblWarning 
         Caption         =   "Note: Using the XP Style in Windows 9x / Me IS NOT recommended."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.PictureBox picColors 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   5640
      ScaleHeight     =   4095
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox picPredefSets 
         BorderStyle     =   0  'None
         Height          =   310
         Left            =   120
         Picture         =   "frmPreferences.frx":0004
         ScaleHeight     =   370.588
         ScaleMode       =   0  'User
         ScaleWidth      =   315
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   120
         Width           =   310
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         ItemData        =   "frmPreferences.frx":014D
         Left            =   4680
         List            =   "frmPreferences.frx":016F
         TabIndex        =   50
         Top             =   120
         Width           =   750
      End
      Begin VB.Frame Frame1 
         Caption         =   "Items"
         Height          =   1575
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   5295
         Begin VB.CheckBox chkUnderline 
            Alignment       =   1  'Right Justify
            Caption         =   "Underline"
            Height          =   195
            Left            =   4200
            TabIndex        =   51
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ListBox lstItems 
            Height          =   1230
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkItalic 
            Alignment       =   1  'Right Justify
            Caption         =   "Italic"
            Height          =   315
            Left            =   4440
            TabIndex        =   44
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkBold 
            Alignment       =   1  'Right Justify
            Caption         =   "Bold"
            Height          =   315
            Left            =   4560
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin ImgColorPicker.ColorPicker cp1 
            Height          =   255
            Left            =   2880
            TabIndex        =   45
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            DefaultCaption  =   "Transparent"
         End
         Begin ImgColorPicker.ColorPicker cp2 
            Height          =   255
            Left            =   2880
            TabIndex        =   47
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            DefaultCaption  =   "Transparent"
         End
         Begin VB.Label lblColor2 
            AutoSize        =   -1  'True
            Caption         =   "Background:"
            Height          =   195
            Left            =   2760
            TabIndex        =   48
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblColor1 
            AutoSize        =   -1  'True
            Caption         =   "Foreground:"
            Height          =   195
            Left            =   2760
            TabIndex        =   46
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox cboFonts 
         Height          =   315
         ItemData        =   "frmPreferences.frx":0197
         Left            =   2160
         List            =   "frmPreferences.frx":0199
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   120
         Width           =   2415
      End
      Begin CodeSenseCtl.CodeSense csPreview 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmPreferences.frx":019B
         TabIndex        =   17
         Top             =   2160
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Font"
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Label lblSubtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Flamebird MX Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmPreferences.frx":0301
      Top             =   0
      Width           =   5565
   End
End
Attribute VB_Name = "frmPreferences"
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

Private Const MSG_SAVEEDITORCONF_INPUTNAME = "Insert a name for the template"

Public WithEvents c As cBrowseForFolder
Attribute c.VB_VarHelpID = -1
Dim d As cCommonDialog

Private m_flat As cFlatControl
Private WithEvents mnuPreSets As cMenus
Attribute mnuPreSets.VB_VarHelpID = -1

'Set control values according to cs configuration
Private Sub RefreshEditorConfigControls()
    chkLineNumbering.value = Abs(CInt(csPreview.LineNumbering))
    chkBookmarkMargin.value = Abs(CInt(csPreview.DisplayLeftMargin))
    chkColorSintax.value = Abs(CInt(csPreview.ColorSyntax))
    chkNormalizeCase.value = Abs(CInt(csPreview.NormalizeCase))
    chkWhiteSpaces.value = Abs(CInt(csPreview.DisplayWhitespace))
    chkSmoothScrolling.value = Abs(CInt(csPreview.SmoothScrolling))
    chkConfine.value = Abs(CInt(csPreview.SelBounds))
    Select Case csPreview.AutoIndentMode
        Case cmIndentOff: opIndentNone.value = True
        Case cmIndentPrevLine: opIndentPrevLine.value = True
        Case cmIndentScope: opIndentScope.value = True
    End Select
    txtTabSize.text = CStr(csPreview.TabSize)

    'Font picker
    FixedPitchFontsToCombo GetDC(csPreview.hwnd), cboFonts
    cboFonts.text = csPreview.font.name
    cboSize.text = CStr(csPreview.font.Size)
End Sub
'Sets FontSytle of the selected item (Bold, italic, underlined)
Private Sub SetStyle()
    Dim i As Integer

    If lstItems.SelCount > 0 Then
        i = lstItems.ItemData(lstItems.ListIndex)
        csPreview.SetFontStyle StyleItem(i).cmStyleItem, chkBold.value * cmFontBold _
                    Or chkItalic.value * cmFontItalic 'Or chkUnderline.value * cmFontUnderline
    End If
End Sub
'Fills the Item list with cmItems Names
Private Sub ListStyles()
    Dim i As Integer

    For i = 0 To StyleItemCount - 1
        lstItems.AddItem StyleItem(i).name
        lstItems.ItemData(i) = i
    Next
    lstItems.Selected(0) = True
End Sub

'Sets controls placement and size
Private Sub PlaceControls()
    tabCategories.Move 0, 920, 5535, 4425
    Me.Width = 5625
    Me.Height = 6395
    cmdCancel.Move 4320, 5380
    cmdOK.Move 3120, 5380
End Sub

'Saves configuration
Private Sub SaveConf()
    On Error GoTo errhandler

    With Ini
        .Path = App.Path & CONF_FILE
        .Section = "General"
        
        .Key = "AskFileRegister"
        .Default = "1"
        .value = IIf(chkAskReg.value = 1, "1", "0")

        .Section = "Appearance"
        
        .Key = "XPStyle"
        .Default = "0"
        .value = IIf(chkEnableXP.value = 1, "1", "0")
        
        .Key = "BitmapBacks"
        .Default = "0"
        .value = IIf(chkBitmap.value = 1, "1", "0")

        .Section = "Run"
        
        .Key = "FenixPath"
        .Default = " "
        .value = txtFenixPath.text
        
        .Key = "Debug"
        .Default = "1"
        .value = IIf(chkDebug.value = 1, "1", "0")
        R_Debug = IIf(chkDebug.value = 1, True, False)
        
        .Key = "Filter"
        .Default = "0"
        .value = IIf(chkFiltering.value = 1, "1", "0")
        R_filter = IIf(chkFiltering.value = 1, True, False)
        
        .Key = "DoubleBuffer"
        .Default = "0"
        .value = IIf(chkDoubleBuf.value = 1, "1", "0")
        R_DoubleBuf = IIf(chkDoubleBuf.value = 1, True, False)
        
        .Key = "SaveBeforeCompiling"
        .Default = "0"
        .value = "0"
        R_SaveBeforeCompiling = 0
        If chkSaveFiles.value = vbChecked Then
            If opCurrentFile.value = True Then
                .value = "1"
                R_SaveBeforeCompiling = 1
            ElseIf opProjectFiles.value = True Then
                .value = "2"
                R_SaveBeforeCompiling = 2
            ElseIf opAllOpenedFiles.value = True Then
                .value = "3"
                R_SaveBeforeCompiling = 3
            End If
        End If
        
        If Not (.Success) Then
           MsgBox "Failed to save value.", vbInformation
        End If
    End With

    'File type association
    If trFiles.Nodes(1).Checked Then
        If Not FileAssociated(".prg", "Fenix.Source") Then
            Call RegisterType(".prg", "Fenix.Source", "Text", "Fenix source file", App.Path + "\Icons\fenix_prg.ico")
        End If
    Else
        If FileAssociated(".prg", "Fenix.Source") Then
            Call DeleteType(".prg", "Fenix.Source")
        End If
    End If

    If trFiles.Nodes(2).Checked Then
        If Not FileAssociated(".map", "Fenix.ImageFile") Then
            Call RegisterType(".map", "Fenix.ImageFile", "Image/Map", "Fenix image file", App.Path + "\Icons\fenix_map.ico")
        End If
    Else
        If FileAssociated(".map", "Fenix.ImageFile") Then
            Call DeleteType(".map", "Fenix.ImageFile")
        End If
    End If

    If trFiles.Nodes(3).Checked Then
        If Not FileAssociated(".fbp", "FlameBird.Project") Then
            Call RegisterType(".fbp", "FlameBird.Project", "Text", "FlameBird project", App.Path + "\Icons\fbp.ico")
        End If
    Else
        If FileAssociated(".fbp", "FlameBird.Project") Then
            Call DeleteType(".fbp", "FlameBird.Project")
        End If
    End If

    'DCBs
    If chkDcb.value = 1 Then
        ' actualizamos siempre el dir de fenix
        If FileAssociated(".dcb", "Fenix.Bin") Then
            Call DeleteType(".dcb", "Fenix.Bin")
        End If
        If Not FileAssociated(".dcb", "Fenix.Bin") Then
            Dim Fxi As String
            With Ini
                .Path = App.Path & CONF_FILE
                .Section = "Run"
                .Key = "FenixPath"
                .Default = " "

                Fxi = .value & "\fxi.exe"
            End With
            If FSO.FileExists(Fxi) Then
                Fxi = Chr(34) & Fxi & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
                Call RegisterType(".dcb", "Fenix.Bin", "Binarie", "Fenix compiled file", App.Path + "\Icons\dcb.ico", Fxi)
            Else
                MsgBox "Can't associate DCB files because the Fenix path isn't configured!!", vbCritical + vbOKOnly, "FlameBird 2"
            End If
        End If
    Else
        If FileAssociated(".dcb", "Fenix.Bin") Then
            Call DeleteType(".dcb", "Fenix.Bin")
        End If
    End If

    'Fenix Directory
    fenixDir = txtFenixPath.text

    'Editor configuration
    SaveCSConf csPreview
    'Apply configuration to all opened documents
    Dim ff As IFileForm, f As Form, fDoc As frmDoc
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.Identify = FF_SOURCE Then
                Set fDoc = f
                LoadCSConf fDoc.cs
            End If
        End If
    Next

    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("frmPreferences.SaveConf")
End Sub

'Load configuration from ini file
Private Sub LoadConf()
    On Error GoTo errhandler:

    With Ini 'Read INI data
        .Path = App.Path & CONF_FILE

        .Section = "General"

        .Key = "AskFileRegister"
        .Default = "1"
        chkAskReg.value = IIf(.value = "1", 1, 0)

        .Section = "Appearance"

        .Key = "XPStyle"
        .Default = "0"
        chkEnableXP.value = IIf(.value = "1", 1, 0)

        .Key = "BitmapBacks"
        .Default = "0"
        chkBitmap.value = IIf(.value = "1", 1, 0)

        .Section = "Run"

        .Key = "FenixPath"
        .Default = " "
        txtFenixPath.text = .value

        .Key = "Debug"
        .Default = "1"
        chkDebug.value = IIf(.value = "1", 1, 0)

        .Key = "Filter"
        .Default = "0"
        chkFiltering.value = IIf(.value = "1", 1, 0)

        .Key = "DoubleBuffer"
        .Default = "0"
        chkDoubleBuf.value = IIf(.value = "1", 1, 0)

        .Key = "SaveBeforeCompiling"
        .Default = "0"
        chkSaveFiles.value = IIf(.value = "1" Or .value = "2" Or .value = "3", 1, 0)
        If .value = "1" Then
            opCurrentFile.value = True
        ElseIf .value = "2" Then
            opProjectFiles.value = True
        ElseIf .value = "3" Then
            opAllOpenedFiles.value = True
        End If
    End With

    LoadCSConf csPreview

    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("frmPreferences.LoadConf")
End Sub


Private Sub cboFonts_Click()
    If cboFonts.ListIndex >= 0 Then
        csPreview.font.name = cboFonts.List(cboFonts.ListIndex)
        csPreview.font.Italic = False
    End If
End Sub

Private Sub cboSize_Click()
    csPreview.font.Size = CDbl(cboSize.text)
End Sub

Private Sub cboSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboSize_Validate (True)
    End If
End Sub

Private Sub cboSize_Validate(Cancel As Boolean)
    If Not IsNumeric(cboSize.text) Then
        Cancel = True
    Else
        csPreview.font.Size = CDbl(cboSize.text)
    End If
End Sub

Private Sub chkBold_Click()
    SetStyle
End Sub

Private Sub chkBookmarkMargin_Click()
    csPreview.DisplayLeftMargin = IIf(chkBookmarkMargin.value = 1, True, False)
End Sub

Private Sub chkColorSintax_Click()
    csPreview.ColorSyntax = IIf(chkColorSintax.value = 1, True, False)
End Sub

Private Sub chkConfine_Click()
    csPreview.SelBounds = IIf(chkConfine.value = 1, True, False)
End Sub

Private Sub chkItalic_Click()
    SetStyle
End Sub

Private Sub chkLineNumbering_Click()
    csPreview.LineNumbering = IIf(chkLineNumbering.value = 1, True, False)
End Sub

Private Sub chkNormalizeCase_Click()
    csPreview.NormalizeCase = IIf(chkNormalizeCase.value = 1, True, False)
End Sub

Private Sub chkSaveFiles_Click()
    Dim en As Boolean
    en = IIf(chkSaveFiles.value <> vbChecked, False, True)
    opCurrentFile.Enabled = en
    opProjectFiles.Enabled = en
    opAllOpenedFiles.Enabled = en
End Sub

Private Sub chkSmoothScrolling_Click()
    csPreview.SmoothScrolling = IIf(chkSmoothScrolling.value = 1, True, False)
End Sub

Private Sub chkWhiteSpaces_Click()
    csPreview.DisplayWhitespace = IIf(chkWhiteSpaces.value = 1, True, False)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExplore_Click()
    Dim s As String

   c.hwndOwner = Me.hwnd
   c.InitialDir = App.Path
   c.FileSystemOnly = True
   c.StatusText = True
   c.UseNewUI = True
   s = c.BrowseForFolder
   If Len(s) > 0 Then
        txtFenixPath.text = s
   End If
End Sub

Private Sub cmdOk_Click()
    SaveConf
    Unload Me
End Sub

Private Sub cp1_ColorChanged()
    If lstItems.SelCount > 0 Then
        csPreview.SetColor StyleItem(lstItems.ItemData(lstItems.ListIndex)).cmItem, cp1.color
    End If
End Sub

Private Sub cp2_ColorChanged()
    If lstItems.SelCount > 0 Then
        csPreview.SetColor StyleItem(lstItems.ItemData(lstItems.ListIndex)).cmItem + 1, cp2.color
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler

    Set c = New cBrowseForFolder
    PlaceControls
    LoadConf

    Set m_flat = New cFlatControl
    m_flat.Attach picPredefSets
    Set mnuPreSets = New cMenus
    mnuPreSets.CreateFromNothing Me.hwnd

    'Create the tabs
    Dim nTab As cTab
    With tabCategories
        .ImageList = 0
        Set nTab = .Tabs.Add("GLOBAL", , "Global")
        nTab.Panel = picAppearance
        Set nTab = .Tabs.Add("EDITOR", , "Editor")
        nTab.Panel = picEditor
        Set nTab = .Tabs.Add("COLORS", , "Colors")
        nTab.Panel = picColors
        Set nTab = .Tabs.Add("COMPILATION", , "Compilation")
        nTab.Panel = picCompilation
        'Set nTab = .Tabs.Add("TOOLS", , "Tools")
        'nTab.Panel = picUserTools
        Set nTab = .Tabs.Add("FILEASSOCIATION", , "File Association")
        nTab.Panel = picFileAsoc
    End With

    'Editor Conf
    ListStyles
    csPreview.Language = "Fenix"
    csPreview.OpenFile (App.Path & "\resources\txtPreview.txt")
    csPreview.READONLY = True
    RefreshEditorConfigControls

    'File association
    With trFiles
        .CheckBoxes = True
        .Nodes.Add(, , "prg", "PRG - Source files").Checked = FileAssociated(".prg", "Fenix.Source")
        .Nodes.Add(, , "map", "MAP - Fenix image files").Checked = FileAssociated(".map", "Fenix.ImageFile")
        .Nodes.Add(, , "fbp", "FBP - FlameBird Project files").Checked = FileAssociated(".fbp", "FlameBird.Project")
    End With
    chkDcb.value = Abs(CInt(FileAssociated(".dcb", "Fenix.Bin")))

    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("frmPreferences.Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mnuPreSets = Nothing
End Sub

Private Sub lstItems_Click()
    Dim clr As OLE_COLOR, clr2 As OLE_COLOR
    Dim Style As clrStyle

    If lstItems.SelCount > 0 Then
        Style = StyleItem(lstItems.ItemData(lstItems.ListIndex))
        cp1.color = csPreview.GetColor(Style.cmItem)
        If Style.extended = True Then
            cp2.color = csPreview.GetColor(Style.cmItem + 1)
            cp2.Visible = True
            lblColor1.Caption = "Foreground:"
            lblColor2.Visible = True
        Else
            lblColor1.Caption = "Color:"
            lblColor2.Visible = False
            cp2.Visible = False
        End If
        If Style.cmStyleItem > -1 Then
            chkBold.value = IIf(csPreview.GetFontStyle(Style.cmStyleItem) And cmFontBold, 1, 0)
            chkItalic.value = IIf(csPreview.GetFontStyle(Style.cmStyleItem) And cmFontItalic, 1, 0)
            'chkUnderline.value = IIf(csPreview.GetFontStyle(Style.cmStyleItem) And cmFontUnderline, 1, 0)
            chkBold.Visible = True
            chkItalic.Visible = True
            'chkUnderline.Visible = True
        Else
            chkBold.Visible = False
            chkItalic.Visible = False
            'chkUnderline.Visible = False
        End If
    End If
End Sub

Private Sub mnuPreSets_Click(ByVal index As Long)
    Dim sKey As String
    Dim sTitle As String

    sKey = mnuPreSets.ItemKey(index)
    If sKey <> "Save" Then
        'If the file exits, load configuration file
        If FSO.FileExists(sKey) Then
            LoadCSConf csPreview, sKey
            RefreshEditorConfigControls
        End If
    Else
        sTitle = InputBox(MSG_SAVEEDITORCONF_INPUTNAME)
        If sTitle <> "" Then
            SaveCSConf csPreview, FSO.BuildPath(App.Path & "\conf\editorstyles", sTitle & ".ini")
        End If
    End If
End Sub

Private Sub opIndentNone_Click()
    csPreview.AutoIndentMode = cmIndentOff
End Sub

Private Sub opIndentPrevLine_Click()
    csPreview.AutoIndentMode = cmIndentPrevLine
End Sub

Private Sub opIndentScope_Click()
    csPreview.AutoIndentMode = cmIndentScope
End Sub

Private Sub picPredefSets_Click()
    Dim i As Integer
    Dim folder As folder
    Dim file As file

    Set mnuPreSets = Nothing
    Set mnuPreSets = New cMenus
    mnuPreSets.CreateFromNothing Me.hwnd

    'Look in the editorstyles folder for config files
    Set folder = FSO.GetFolder(App.Path & "\conf\editorstyles\")
    If Not folder Is Nothing Then
        For Each file In folder.Files
            If FSO.GetExtensionName(file.Path) = "ini" Then
                mnuPreSets.AddItem 0, FSO.GetBaseName(file.Path), , , file.Path
            End If
        Next
    End If
    If mnuPreSets.ItemCount > 0 Then mnuPreSets.AddItem 0, "-"
    mnuPreSets.AddItem 0, "Save...", , , "Save"
    mnuPreSets.PopupMenu
    Me.SetFocus
End Sub

Private Sub txtTabSize_Change()
    csPreview.TabSize = CInt(txtTabSize.text)
End Sub

Private Sub txtTabSize_Validate(Cancel As Boolean)
    Cancel = True
    If IsNumeric(txtTabSize.text) Then
        If CInt(txtTabSize.text) = Val(txtTabSize.text) Then 'Ensure no decimals
            Cancel = False
        End If
    End If
End Sub
