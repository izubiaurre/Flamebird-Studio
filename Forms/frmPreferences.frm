VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbaltreeview6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Object = "{C8A61D56-D8DC-11D2-8064-9D6F06504DA8}#1.1#0"; "axcolctl.ocx"
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   17550
   ClientLeft      =   3150
   ClientTop       =   1005
   ClientWidth     =   17910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   17550
   ScaleWidth      =   17910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCompilerPaths 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   11760
      ScaleHeight     =   4095
      ScaleWidth      =   5535
      TabIndex        =   118
      Top             =   4080
      Width           =   5535
      Begin VB.Frame fraPATHS 
         Caption         =   "Compiler PATHS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   0
         TabIndex        =   119
         Top             =   120
         Width           =   5415
         Begin VB.ListBox lstPATHS 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   4095
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            FillStyle       =   2  'Horizontal Line
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   4320
            ScaleHeight     =   2535
            ScaleWidth      =   975
            TabIndex        =   120
            Top             =   240
            Width           =   975
            Begin VB.CommandButton cmdPATHSRemoveAll 
               Caption         =   "Remove All"
               Height          =   375
               Left            =   0
               TabIndex        =   123
               Top             =   1320
               Width           =   975
            End
            Begin VB.CommandButton cmdPATHSRemove 
               Caption         =   "Remove"
               Height          =   375
               Left            =   0
               TabIndex        =   122
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton cmdPATHSAdd 
               Caption         =   "Add"
               Height          =   375
               Left            =   0
               TabIndex        =   121
               Top             =   120
               Width           =   975
            End
         End
      End
   End
   Begin vbalTreeViewLib6.vbalTreeView tv_preferences 
      Height          =   3255
      Left            =   5760
      TabIndex        =   110
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5741
      PathSeparator   =   "/"
      BackColor       =   -2147483647
      BorderStyle     =   0
      ForeColor       =   -2147483633
      LineColor       =   -2147483643
      SelectedForeColor=   -2147483648
      SelectedForeColor=   -2147483648
      SelectedForeColor=   -2147483648
      SelectedForeColor=   -2147483648
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
   Begin VB.PictureBox picMisc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   11340
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   102
      Top             =   12600
      Width           =   5295
      Begin VB.CommandButton cmdClearCommandHistory 
         Caption         =   "Clear Comand History"
         Height          =   375
         Left            =   3420
         TabIndex        =   104
         ToolTipText     =   "Clears the MS-DOS command history."
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdClearRecents 
         Caption         =   "Clear Recent List"
         Height          =   375
         Left            =   3420
         TabIndex        =   103
         ToolTipText     =   "Clears the recently opened files list"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox picProgramInspector 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   5580
      ScaleHeight     =   4395
      ScaleWidth      =   5595
      TabIndex        =   81
      Top             =   12660
      Width           =   5595
      Begin VB.CheckBox chkPIOnlyLocalHeader 
         Caption         =   "Only locals header"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   88
         Top             =   1020
         Width           =   2295
      End
      Begin VB.CheckBox chkPIOnlyConsHeader 
         Caption         =   "Only constant header"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   86
         Top             =   180
         Width           =   2175
      End
      Begin VB.CheckBox chkPIShowPrivates 
         Caption         =   "Show Privates"
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
         Left            =   180
         TabIndex        =   85
         Top             =   1440
         Width           =   2235
      End
      Begin VB.CheckBox chkPILocals 
         Caption         =   "Show Locals"
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
         Left            =   180
         TabIndex        =   84
         Top             =   1020
         Width           =   2235
      End
      Begin VB.CheckBox chkPIShowGlobals 
         Caption         =   "Show Globals"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   83
         Top             =   600
         Width           =   2235
      End
      Begin VB.CheckBox chkPIShowCons 
         Caption         =   "Show Constants"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   82
         Top             =   180
         Width           =   2115
      End
      Begin VB.CheckBox chkPIOnlyGlobalHeader 
         Caption         =   "Only global header"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   87
         Top             =   600
         Width           =   2235
      End
   End
   Begin VB.PictureBox picIntelliSense 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   67
      Top             =   12480
      Width           =   5295
      Begin VB.TextBox txtIntelliSenseSensitive 
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
         Left            =   4380
         MaxLength       =   1
         TabIndex        =   79
         Top             =   270
         Width           =   555
      End
      Begin VB.Frame fraIntelliSenseFilter 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   120
         TabIndex        =   69
         Top             =   660
         Width           =   5055
         Begin VB.CheckBox chkISUserProcs 
            Caption         =   "Processes"
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
            Left            =   3540
            TabIndex        =   76
            Top             =   1860
            Width           =   1215
         End
         Begin VB.CheckBox chkISUserFuncs 
            Caption         =   "Functions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3540
            TabIndex        =   75
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CheckBox chkISUserVars 
            Caption         =   "Variables"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3540
            TabIndex        =   74
            Top             =   1080
            Width           =   1155
         End
         Begin VB.CheckBox chkISUserCons 
            Caption         =   "Contants"
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
            Left            =   3540
            TabIndex        =   73
            Top             =   720
            Width           =   1035
         End
         Begin VB.CheckBox chkISFuncs 
            Caption         =   "Functions"
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
            Left            =   240
            TabIndex        =   72
            Top             =   1440
            Width           =   1275
         End
         Begin VB.CheckBox chkISVars 
            Caption         =   "Variables"
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
            Left            =   240
            TabIndex        =   71
            Top             =   1080
            Width           =   1155
         End
         Begin VB.CheckBox chkISLangCons 
            Caption         =   "Constants"
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
            Left            =   240
            TabIndex        =   70
            Top             =   720
            Width           =   1095
         End
         Begin VB.Line lblIS2 
            X1              =   3060
            X2              =   4860
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Line lineIS1 
            X1              =   120
            X2              =   1935
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Label lblISUserDefined 
            Alignment       =   2  'Center
            Caption         =   "User defined"
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
            Left            =   3240
            TabIndex        =   78
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblISLangDefined 
            Alignment       =   2  'Center
            Caption         =   "Language defined"
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
            TabIndex        =   77
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkShowIntelliSense 
         Caption         =   "Show IntelliSense"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblIntelliSenseSensitive 
         Caption         =   "IntelliSense sensitive:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   80
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.PictureBox picCompilerOptions 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   5295
      TabIndex        =   52
      Top             =   4440
      Width           =   5295
      Begin VB.CheckBox chkDirs 
         Caption         =   "Add directories to the PATH (-i):"
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
         Left            =   240
         TabIndex        =   117
         Top             =   1440
         Width           =   4095
      End
      Begin VB.CheckBox chkDebugDCB 
         Caption         =   "Store debugging information at the DCB (-g)"
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
         Left            =   240
         TabIndex        =   116
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtParams 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   2760
         Width           =   4695
      End
      Begin VB.CheckBox chkMSDOS 
         Caption         =   "File uses the MS-DOS character set (-c)"
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
         Left            =   240
         TabIndex        =   57
         Top             =   960
         Width           =   4695
      End
      Begin VB.CheckBox chkStub 
         Caption         =   "Generate a stubbed executable from the given stub (-s stub)"
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
         Left            =   240
         TabIndex        =   55
         Top             =   720
         Width           =   4695
      End
      Begin VB.CheckBox chkAutoDeclare 
         Caption         =   "Enable automatic declare functions ( -Ca)"
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
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Compile in Debug mode (-d)"
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
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label lblParams 
         Caption         =   "Command-line parameters:"
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
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   2175
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
         TabIndex        =   56
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   5640
      ScaleHeight     =   4095
      ScaleWidth      =   5535
      TabIndex        =   47
      Top             =   8280
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Helping line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2880
         TabIndex        =   105
         Top             =   2400
         Width           =   2535
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   2295
            TabIndex        =   106
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton optHelpLine 
               Caption         =   "Don't show"
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
               Index           =   2
               Left            =   0
               TabIndex        =   109
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton optHelpLine 
               Caption         =   "Show line under"
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
               Index           =   1
               Left            =   0
               TabIndex        =   108
               Top             =   360
               Width           =   1815
            End
            Begin VB.OptionButton optHelpLine 
               Caption         =   "Show line upper"
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
               Index           =   0
               Left            =   0
               TabIndex        =   107
               Top             =   120
               Width           =   1575
            End
         End
      End
      Begin VB.CheckBox chkLineNumbering 
         Caption         =   "Display line number margin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox chkBookmarkMargin 
         Caption         =   "Display bookmark margin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   750
         Width           =   2175
      End
      Begin VB.CheckBox chkColorSintax 
         Caption         =   "Color syntax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   1245
         Width           =   1335
      End
      Begin VB.CheckBox chkNormalizeCase 
         Caption         =   "Normalize keyword case"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   1755
         Width           =   2175
      End
      Begin VB.Frame grbAutoIdent 
         Caption         =   "Auto indent mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   2880
         TabIndex        =   92
         Top             =   600
         Width           =   2535
         Begin VB.PictureBox picIndent 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1575
            TabIndex        =   93
            Top             =   360
            Width           =   1575
            Begin VB.OptionButton opIndentScope 
               Caption         =   "Scope"
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
               Left            =   0
               TabIndex        =   96
               Top             =   840
               Width           =   975
            End
            Begin VB.OptionButton opIndentPrevLine 
               Caption         =   "Previous line"
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
               Left            =   0
               TabIndex        =   95
               Top             =   480
               Width           =   1335
            End
            Begin VB.OptionButton opIndentNone 
               Caption         =   "None"
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
               Left            =   0
               TabIndex        =   94
               Top             =   120
               Width           =   855
            End
         End
      End
      Begin VB.CheckBox chkWhiteSpaces 
         Caption         =   "Display white spaces"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   2250
         Width           =   1935
      End
      Begin VB.CheckBox chkSmoothScrolling 
         Caption         =   "Smooth scrolling"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtTabSize 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkConfine 
         Caption         =   "Confine caret to text"
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
         TabIndex        =   49
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tab size:"
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
         Left            =   3840
         TabIndex        =   101
         Top             =   285
         Width           =   660
      End
   End
   Begin VB.PictureBox picFileAsoc 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   11400
      ScaleHeight     =   3975
      ScaleWidth      =   5295
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox chkAskReg 
         Caption         =   "Ask for File Association on init"
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
         TabIndex        =   32
         Top             =   3120
         Width           =   3435
      End
      Begin VB.Frame fraFiletypes 
         Height          =   3015
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   5175
         Begin VB.CheckBox chkDcb 
            Caption         =   "Open DCB files with Fenix/Bennu Interpreter"
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
            Left            =   600
            TabIndex        =   29
            Top             =   2520
            Width           =   4095
         End
         Begin vbalTreeViewLib6.vbalTreeView trFiles 
            Height          =   1815
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3201
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
            TabIndex        =   31
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Label lblNotice 
         Caption         =   "Note: This will apply only when at least one filetype isn't registered."
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
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   5175
      End
   End
   Begin VB.PictureBox picUserTools 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   0
      Left            =   11640
      ScaleHeight     =   3975
      ScaleWidth      =   5295
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame fraToolData 
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   5175
         Begin VB.TextBox txtName 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   23
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   22
            Top             =   720
            Width           =   3735
         End
         Begin VB.CommandButton cmdAddTool 
            Caption         =   "&Add"
            Height          =   375
            Index           =   0
            Left            =   3840
            TabIndex        =   21
            ToolTipText     =   "Add new tool"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "C&lear"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdToolExplore 
            Caption         =   "..."
            Height          =   315
            Left            =   4560
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtParms 
            Height          =   285
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Insert here any command-line parameter you want to pass to the app"
            Top             =   1440
            Width           =   3855
         End
         Begin VB.Label lblName 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblPath 
            Caption         =   "Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblParms 
            Caption         =   "Command-line parameters:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   1935
         End
      End
      Begin VB.Frame fraTools 
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   5175
         Begin VB.ListBox lstUserTools 
            Height          =   1230
            ItemData        =   "frmPreferences.frx":0000
            Left            =   120
            List            =   "frmPreferences.frx":0002
            TabIndex        =   16
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdRemoveTool 
            Caption         =   "R&emove"
            Height          =   375
            Left            =   3840
            TabIndex        =   15
            ToolTipText     =   "Remove selected tool"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remove all"
            Height          =   375
            Left            =   3840
            TabIndex        =   14
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
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4095
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   8160
      Width           =   5295
      Begin VB.Frame Frame2 
         Caption         =   "Compiler type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   0
         TabIndex        =   50
         Top             =   120
         Width           =   5295
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   5055
            TabIndex        =   64
            Top             =   240
            Width           =   5055
            Begin VB.TextBox txtCompilerPath 
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
               Index           =   1
               Left            =   0
               TabIndex        =   115
               Top             =   1200
               Width           =   4455
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
               Index           =   1
               Left            =   4560
               TabIndex        =   114
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox txtCompilerPath 
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
               Index           =   0
               Left            =   0
               TabIndex        =   112
               Top             =   480
               Width           =   4455
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
               Index           =   0
               Left            =   4560
               TabIndex        =   111
               Top             =   480
               Width           =   495
            End
            Begin VB.OptionButton optFenixBennu 
               Caption         =   "Bennu"
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
               Index           =   1
               Left            =   0
               TabIndex        =   66
               Top             =   840
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optFenixBennu 
               Caption         =   "Fenix"
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
               Index           =   0
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   975
            End
            Begin VB.Line Line1 
               X1              =   2400
               X2              =   4320
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Label lblFenixPath 
               Caption         =   "Compiler path:"
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
               Left            =   2880
               TabIndex        =   113
               Top             =   0
               Width           =   1215
            End
         End
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
         TabIndex        =   6
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Frame grbSaveBeforeCompiling 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   240
         TabIndex        =   48
         Top             =   2760
         Width           =   3255
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   2655
            TabIndex        =   60
            Top             =   120
            Width           =   2655
            Begin VB.OptionButton opAllOpenedFiles 
               Caption         =   "All opened files"
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
               TabIndex        =   63
               Top             =   600
               Width           =   1575
            End
            Begin VB.OptionButton opProjectFiles 
               Caption         =   "Project files"
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
               TabIndex        =   62
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton opCurrentFile 
               Caption         =   "Current file only"
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
               TabIndex        =   61
               Top             =   120
               Width           =   1575
            End
         End
      End
   End
   Begin VB.PictureBox picAppearance 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmPreferences.frx":0004
         Left            =   2520
         List            =   "frmPreferences.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   720
         Width           =   1695
      End
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      TabIndex        =   9
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox picPredefSets 
         BorderStyle     =   0  'None
         Height          =   310
         Left            =   120
         Picture         =   "frmPreferences.frx":0008
         ScaleHeight     =   370.588
         ScaleMode       =   0  'User
         ScaleWidth      =   315
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   120
         Width           =   310
      End
      Begin VB.ComboBox cboSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPreferences.frx":0151
         Left            =   4680
         List            =   "frmPreferences.frx":0173
         TabIndex        =   43
         Top             =   120
         Width           =   750
      End
      Begin VB.Frame Frame1 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   5295
         Begin VB.CheckBox chkUnderline 
            Alignment       =   1  'Right Justify
            Caption         =   "Underline"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ListBox lstItems 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkItalic 
            Alignment       =   1  'Right Justify
            Caption         =   "Italic"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4560
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox chkBold 
            Alignment       =   1  'Right Justify
            Caption         =   "Bold"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4560
            TabIndex        =   36
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin ImgColorPicker.ColorPicker cp1 
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            DefaultCaption  =   "Transparent"
         End
         Begin ImgColorPicker.ColorPicker cp2 
            Height          =   255
            Left            =   2880
            TabIndex        =   40
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
            Left            =   2760
            TabIndex        =   41
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblColor1 
            AutoSize        =   -1  'True
            Caption         =   "Foreground:"
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
            Left            =   2760
            TabIndex        =   39
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.ComboBox cboFonts 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPreferences.frx":019B
         Left            =   2160
         List            =   "frmPreferences.frx":019D
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   120
         Width           =   2415
      End
      Begin CodeSenseCtl.CodeSense csPreview 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmPreferences.frx":019F
         TabIndex        =   11
         Top             =   2160
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Font"
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
         Left            =   1680
         TabIndex        =   46
         Top             =   150
         Width           =   495
      End
   End
   Begin vbalDTab6.vbalDTabControl tabCategories 
      Height          =   480
      Left            =   7200
      TabIndex        =   10
      Top             =   360
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
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7485
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
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
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

Private Const MSG_SAVEEDITORCONF_INPUTNAME = "Insert a name for the template"

Private m_flat As cFlatControl
Public WithEvents c As cBrowseForFolder
Attribute c.VB_VarHelpID = -1
Dim d As cCommonDialog

Private WithEvents mnuPreSets As cMenus
Attribute mnuPreSets.VB_VarHelpID = -1

'Set control values according to cs configuration
Private Sub RefreshEditorConfigControls()
    chkLineNumbering.Value = Abs(CInt(csPreview.LineNumbering))
    chkBookmarkMargin.Value = Abs(CInt(csPreview.DisplayLeftMargin))
    chkColorSintax.Value = Abs(CInt(csPreview.ColorSyntax))
    chkNormalizeCase.Value = Abs(CInt(csPreview.NormalizeCase))
    chkWhiteSpaces.Value = Abs(CInt(csPreview.DisplayWhitespace))
    chkSmoothScrolling.Value = Abs(CInt(csPreview.SmoothScrolling))
    chkConfine.Value = Abs(CInt(csPreview.SelBounds))
    Select Case csPreview.AutoIndentMode
        Case cmIndentOff: opIndentNone.Value = True
        Case cmIndentPrevLine: opIndentPrevLine.Value = True
        Case cmIndentScope: opIndentScope.Value = True
    End Select
    txtTabSize.text = CStr(csPreview.TabSize)

    'Font picker
    FixedPitchFontsToCombo GetDC(csPreview.Hwnd), cboFonts
    cboFonts.text = csPreview.font.name
    cboSize.text = CStr(csPreview.font.Size)
End Sub
'Sets FontSytle of the selected item (Bold, italic, underlined)
Private Sub SetStyle()
    Dim i As Integer

    If lstItems.SelCount > 0 Then
        i = lstItems.ItemData(lstItems.ListIndex)
        csPreview.SetFontStyle StyleItem(i).cmStyleItem, chkBold.Value * cmFontBold _
                    Or chkItalic.Value * cmFontItalic 'Or chkUnderline.value * cmFontUnderline
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
    tabCategories.Move 1820, 800, 5535, 4545  ' 0, 800, 5535, 4425
    tabCategories.ShowTabs = False
    Me.Width = 7440     ' 5625
    Me.Height = 6295    ' 6395
    cmdCancel.Move 4320, 5400   ' 5380
    cmdOk.Move 3120, 5400        ' 3120
    tv_preferences.Move 0, 800, 1815, 6295
End Sub

'Saves configuration
Private Sub SaveConf()
    On Error GoTo errhandler

    With Ini
        .Path = App.Path & CONF_FILE
        .Section = "General"        ' ----------- GENERAL ------------------
        
        .Key = "AskFileRegister"
        .Default = "1"
        .Value = IIf(chkAskReg.Value = 1, "1", "0")
        
        .Key = "ProcHelpLine"
        .Default = "1"
        .Value = IIf(G_ProcHelpLine = 0, "0", IIf(G_ProcHelpLine = 1, "1", "-1"))

        .Section = "Appearance"     ' ----------- APPEARANCE ---------------
        
        .Key = "XPStyle"
        .Default = "0"
        .Value = IIf(chkEnableXP.Value = 1, "1", "0")
        
        .Key = "BitmapBacks"
        .Default = "0"
        .Value = IIf(chkBitmap.Value = 1, "1", "0")
        
        .Key = "Color"
        .Default = "1"
        .Value = cmbColor.ListIndex

        .Section = "Run"            ' ----------- RUN ----------------------
        
        .Key = "FenixPath"
        .Default = " "
        .Value = txtCompilerPath(0).text
        
        .Key = "BennuPath"
        .Default = " "
        .Value = txtCompilerPath(1).text
        
        .Key = "Compiler"
        .Default = "1"
        .Value = IIf(optFenixBennu(0).Value = True, "0", "1")
        R_Compiler = IIf(optFenixBennu(0).Value = True, "0", "1")
        
        .Key = "Debug"
        .Default = "1"
        .Value = IIf(chkDebug.Value = 1, "1", "0")
        R_Debug = IIf(chkDebug.Value = 1, True, False)
        
        .Key = "AutoDeclare"
        .Default = "1"
        .Value = IIf(chkAutoDeclare.Value = 1, "1", "0")
        R_AutoDeclare = IIf(chkAutoDeclare.Value = 1, True, False)
        
        .Key = "Stub"
        .Default = "0"
        .Value = IIf(chkStub.Value = 1, "1", "0")
        R_Stub = IIf(chkStub.Value = 1, True, False)
        
        .Key = "MsDos"
        .Default = "0"
        .Value = IIf(chkMSDOS.Value = 1, "1", "0")
        R_MsDos = IIf(chkMSDOS.Value = 1, True, False)
        
        .Key = "DebugDCB"
        .Default = "0"
        .Value = IIf(chkDebugDCB.Value = 1, "1", "0")
        R_DebugDCB = IIf(chkDebugDCB.Value = 1, True, False)
        
        .Key = "Paths"
        .Default = "0"
        .Value = IIf(chkDirs.Value = 1, "1", "0")
        R_Paths = IIf(chkDirs.Value = 1, True, False)
        
'        .Key = "Filter"
'        .Default = "0"
'        .value = IIf(chkFiltering.value = 1, "1", "0")
'        R_filter = IIf(chkFiltering.value = 1, True, False)
        
'        .Key = "DoubleBuffer"
'        .Default = "0"
'        .value = IIf(chkDoubleBuf.value = 1, "1", "0")
'        R_DoubleBuf = IIf(chkDoubleBuf.value = 1, True, False)
        
        .Key = "SaveBeforeCompiling"
        .Default = "0"
        .Value = "0"
        R_SaveBeforeCompiling = 0
        If chkSaveFiles.Value = vbChecked Then
            If opCurrentFile.Value = True Then
                .Value = "1"
                R_SaveBeforeCompiling = 1
            ElseIf opProjectFiles.Value = True Then
                .Value = "2"
                R_SaveBeforeCompiling = 2
            ElseIf opAllOpenedFiles.Value = True Then
                .Value = "3"
                R_SaveBeforeCompiling = 3
            End If
        End If
        
        .Section = "IntelliSense"       ' ----------- INTELLISENSE -------------
        
        .Key = "Show"
        .Default = "1"
        .Value = IIf(chkShowIntelliSense.Value = "1", "1", "0")
        IS_Show = IIf(chkShowIntelliSense.Value = "1", True, False)
        
        'If chkShowIntelliSense Then
        
            .Key = "Sensitive"
            .Default = "2"
            .Value = txtIntelliSenseSensitive
            IS_Sensitive = CInt(txtIntelliSenseSensitive)
        
            .Key = "LangDefConst"
            .Default = "1"
            .Value = IIf(chkISLangCons.Value = "1", "1", "0")
            IS_LangDefConst = IIf(chkISLangCons.Value = "1", True, False)
            
            .Key = "LangDefVar"
            .Default = "1"
            .Value = IIf(chkISVars.Value = "1", "1", "0")
            IS_LangDefVar = IIf(chkISVars.Value = "1", True, False)
            
            .Key = "LangDefFunc"
            .Default = "1"
            .Value = IIf(chkISFuncs.Value = "1", "1", "0")
            IS_LangDefFunc = IIf(chkISFuncs.Value = "1", True, False)
            
            .Key = "UserDefConst"
            .Default = "1"
            .Value = IIf(chkISUserCons.Value = "1", "1", "0")
            IS_UserDefConst = IIf(chkISUserCons.Value = "1", True, False)
            
            .Key = "UserDefvar"
            .Default = "1"
            .Value = IIf(chkISUserVars.Value = "1", "1", "0")
            IS_UserDefVar = IIf(chkISUserVars.Value = "1", True, False)
            
            .Key = "UserDefFunc"
            .Default = "1"
            .Value = IIf(chkISUserFuncs.Value = "1", "1", "0")
            IS_UserDefFunc = IIf(chkISUserFuncs.Value = "1", True, False)
            
            .Key = "UserDefProc"
            .Default = "1"
            .Value = IIf(chkISUserProcs.Value = "1", "1", "0")
            IS_UserDefProc = IIf(chkISUserProcs.Value = "1", True, False)
            
        'End If
        
        .Section = "ProgramInspector"   ' ----------- PROGRAMINSPECTOR ---------
        
        .Key = "ShowConsts"
        .Default = "1"
        .Value = IIf(chkPIShowCons.Value = "1", "1", "0")
        PI_ShowConsts = IIf(chkPIShowCons.Value = "1", True, False)
        
        .Key = "ShowGlobals"
        .Default = "1"
        .Value = IIf(chkPIShowGlobals.Value = "1", "1", "0")
        PI_ShowGlobals = IIf(chkPIShowGlobals.Value = "1", True, False)
        
        .Key = "ShowLocals"
        .Default = "1"
        .Value = IIf(chkPILocals.Value = "1", "1", "0")
        PI_ShowLocals = IIf(chkPILocals.Value = "1", True, False)
        
        .Key = "ShowPrivates"
        .Default = "1"
        .Value = IIf(chkPIShowPrivates.Value = "1", "1", "0")
        PI_ShowPrivates = IIf(chkPIShowPrivates.Value = "1", True, False)
        
        .Key = "OnlyConstHeader"
        .Default = "1"
        .Value = IIf(chkPIOnlyConsHeader.Value = "1", "1", "0")
        PI_OnlyConstHeader = IIf(chkPIOnlyConsHeader.Value = "1", True, False)
        
        .Key = "OnlyGlobalHeader"
        .Default = "1"
        .Value = IIf(chkPIOnlyGlobalHeader.Value = "1", "1", "0")
        PI_OnlyGlobalHeader = IIf(chkPIOnlyGlobalHeader.Value = "1", True, False)
        
        .Key = "OnlyLocalHeader"
        .Default = "1"
        .Value = IIf(chkPIOnlyLocalHeader.Value = "1", "1", "0")
        PI_OnlyLocalHeader = IIf(chkPIOnlyLocalHeader.Value = "1", True, False)
        
        If Not (.Success) Then
           MsgBox "Failed to save value.", vbInformation
        End If
    End With

    'File type association
    If trFiles.Nodes(1).Checked Then
        If Not FileAssociated(".prg", "Bennu/Fenix.Source") Then
            Call RegisterType(".prg", "Bennu/Fenix.Source", "Text", "Bennu/Fenix source file", App.Path + "\Icons\fenix_prg.ico")
        End If
    Else
        If FileAssociated(".prg", "Bennu/Fenix.Source") Then
            Call DeleteType(".prg", "Bennu/Fenix.Source")
        End If
    End If

    If trFiles.Nodes(2).Checked Then
        If Not FileAssociated(".map", "Fenix.ImageFile") Then
            Call RegisterType(".map", "Bennu/Fenix.ImageFile", "Image/Map", "Bennu/Fenix image file", App.Path + "\Icons\fenix_map.ico")
        End If
    Else
        If FileAssociated(".map", "Bennu/Fenix.ImageFile") Then
            Call DeleteType(".map", "Bennu/Fenix.ImageFile")
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
    
    If trFiles.Nodes(4).Checked Then
        If Not FileAssociated(".bmk", "FlameBird.Bookmark") Then
            Call RegisterType(".bmk", "FlameBird.Bookmark", "Text", "FlameBird source bookmark files", App.Path + "\Icons\FBMX_bmk.ico")
        End If
    Else
        If FileAssociated(".bmk", "FlameBird.Bookmark") Then
            Call DeleteType(".bmk", "FlameBird.Bookmark")
        End If
    End If
    
    If trFiles.Nodes(5).Checked Then
        If Not FileAssociated(".cpt", "FlameBird.ControlPoint") Then
            Call RegisterType(".cpt", "FlameBird.ControlPoint", "Image/Map", "Bennu/Fenix image file Control Point lists", App.Path + "\Icons\FBMX_cpt.ico")
        End If
    Else
        If FileAssociated(".cpt", "FlameBird.ControlPoint") Then
            Call DeleteType(".cpt", "FlameBird.ControlPoint")
        End If
    End If

    'DCBs
    If chkDcb.Value = 1 Then
        ' actualizamos siempre el dir de fenix
        If FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
            Call DeleteType(".dcb", "Bennu/Fenix.Bin")
        End If
        If Not FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
            Dim Fxi As String
            With Ini
                .Path = App.Path & CONF_FILE
                .Section = "Run"
                
                If optFenixBennu(0) Then
                    .Key = "FenixPath"
                    .Default = " "
                    Fxi = .Value & "\fxi.exe"
                Else
                    .Key = "BennuPath"
                    .Default = " "
                    Fxi = .Value & "\bgdi.exe"
                End If
            End With
            If FSO.FileExists(Fxi) Then
                Fxi = Chr(34) & Fxi & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
                Call RegisterType(".dcb", "Bennu/Fenix.Bin", "Binarie", "Bennu/Fenix compiled file", App.Path + "\Icons\dcb.ico", Fxi)
            Else
                MsgBox "Can't associate DCB files because the compiler path isn't configured!!", vbCritical + vbOKOnly, "FlameBirdMX"
            End If
        End If
    Else
        If FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
            Call DeleteType(".dcb", "Bennu/Fenix.Bin")
        End If
    End If

    'Fenix Directory
    If R_Compiler = 0 Then
        fenixDir = txtCompilerPath(0).text
    Else
        fenixDir = txtCompilerPath(1).text
    End If

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

        .Section = "General"        ' ----------- GENERAL ------------------

        .Key = "AskFileRegister"
        .Default = "1"
        chkAskReg.Value = IIf(.Value = "1", 1, 0)
        
        .Key = "ProcHelpLine"
        .Default = "1"
        If .Value = 0 Then
            optHelpLine(1).Value = True
        ElseIf .Value = 1 Then
            optHelpLine(0).Value = True
        Else
            optHelpLine(2).Value = True
        End If

        .Section = "Appearance"     ' ----------- APPEARANCE ---------------

        .Key = "XPStyle"
        .Default = "0"
        chkEnableXP.Value = IIf(.Value = "1", 1, 0)

        .Key = "BitmapBacks"
        .Default = "0"
        chkBitmap.Value = IIf(.Value = "1", 1, 0)
        
        .Key = "Color"
        .Default = "1"
        cmbColor.ListIndex = IIf(.Value = "1" Or .Value = "2" Or .Value = "3" Or .Value = "4" Or .Value = "5" Or .Value = "6" Or .Value = "7" Or .Value = "8" Or .Value = "9" Or .Value = "0", .Value, 1)

        .Section = "Run"            ' ----------- RUN ----------------------

        .Key = "FenixPath"
        .Default = " "
        txtCompilerPath(0).text = .Value
        
        .Key = "BennuPath"
        .Default = " "
        txtCompilerPath(1).text = .Value
        
        .Key = "Compiler"
        .Default = "1"
        optFenixBennu(.Value).Value = True
        '.Value = IIf(optFenixBennu(0).Value = True, "0", "1")

        .Key = "Debug"
        .Default = "1"
        chkDebug.Value = IIf(.Value = "1", 1, 0)
        
        .Key = "AutoDeclare"
        .Default = "1"
        chkAutoDeclare.Value = IIf(.Value = 1, "1", "0")
        
        .Key = "Stub"
        .Default = "0"
        chkStub.Value = IIf(.Value = 1, "1", "0")
        
        .Key = "MsDos"
        .Default = "0"
        chkMSDOS.Value = IIf(.Value = 1, "1", "0")
        
        .Key = "DebugDCB"
        .Default = "0"
        chkDebugDCB.Value = IIf(.Value = "1", 1, 0)
        
        .Key = "Paths"
        .Default = "0"
        chkDirs.Value = IIf(.Value = "1", 1, 0)
        
'        .Key = "Filter"
'        .Default = "0"
'        chkAutoDeclare.value = IIf(.value = "1", 1, 0)

'        .Key = "DoubleBuffer"
'        .Default = "0"
'        chkDoubleBuf.value = IIf(.value = "1", 1, 0)

        .Key = "SaveBeforeCompiling"
        .Default = "0"
        chkSaveFiles.Value = IIf(.Value = "1" Or .Value = "2" Or .Value = "3", 1, 0)
        If .Value = "1" Then
            opCurrentFile.Value = True
        ElseIf .Value = "2" Then
            opProjectFiles.Value = True
        ElseIf .Value = "3" Then
            opAllOpenedFiles.Value = True
        End If
        
        .Section = "IntelliSense"       ' ----------- INTELLISENSE -------------

        .Key = "Show"
        .Default = "1"
        chkShowIntelliSense.Value = IIf(.Value = 1, 1, 0)

        'If chkShowIntelliSense Then

            .Key = "Sensitive"
            .Default = "2"
            txtIntelliSenseSensitive = CInt(.Value)

            .Key = "LangDefConst"
            .Default = "1"
            chkISLangCons.Value = IIf(.Value = 1, 1, 0)

            .Key = "LangDefVar"
            .Default = "1"
            chkISVars.Value = IIf(.Value = 1, 1, 0)

            .Key = "LangDefFunc"
            .Default = "1"
            chkISFuncs.Value = IIf(.Value = 1, 1, 0)

            .Key = "UserDefConst"
            .Default = "1"
            chkISUserCons.Value = IIf(.Value = 1, 1, 0)

            .Key = "UserDefvar"
            .Default = "1"
            chkISUserVars.Value = IIf(.Value = 1, 1, 0)

            .Key = "UserDefFunc"
            .Default = "1"
            chkISUserFuncs.Value = IIf(.Value = 1, 1, 0)
            
            .Key = "UserDefProc"
            .Default = "1"
            chkISUserProcs.Value = IIf(.Value = 1, 1, 0)

        'End If

        .Section = "ProgramInspector"   ' ----------- PROGRAMINSPECTOR ---------

        .Key = "ShowConsts"
        .Default = "1"
        chkPIShowCons.Value = IIf(.Value = 1, 1, 0)

        .Key = "ShowGlobals"
        .Default = "1"
        chkPIShowGlobals = IIf(.Value = 1, 1, 0)

        .Key = "ShowLocals"
        .Default = "1"
        chkPILocals = IIf(.Value = 1, 1, 0)

        .Key = "ShowPrivates"
        .Default = "1"
        chkPIShowPrivates = IIf(.Value = 1, 1, 0)

        .Key = "OnlyConstHeader"
        .Default = "1"
        chkPIOnlyConsHeader = IIf(.Value = 1, 1, 0)

        .Key = "OnlyGlobalHeader"
        .Default = "1"
        chkPIOnlyGlobalHeader = IIf(.Value = 1, 1, 0)

        .Key = "OnlyLocalHeader"
        .Default = "1"
        chkPIOnlyLocalHeader = IIf(.Value = 1, 1, 0)
        
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

Private Sub chkAutoDeclare_Click()
    printParams
End Sub

Private Sub chkBold_Click()
    SetStyle
End Sub

Private Sub chkBookmarkMargin_Click()
    csPreview.DisplayLeftMargin = IIf(chkBookmarkMargin.Value = 1, True, False)
End Sub

Private Sub chkColorSintax_Click()
    csPreview.ColorSyntax = IIf(chkColorSintax.Value = 1, True, False)
End Sub

Private Sub chkConfine_Click()
    csPreview.SelBounds = IIf(chkConfine.Value = 1, True, False)
End Sub

Private Sub chkDebug_Click()
    printParams
End Sub

Private Sub chkDebugDCB_Click()
    printParams
End Sub

Private Sub chkDirs_Click()
    If chkDirs.Value Then    ' go to compilation dirs tab
        tabCategories.Tabs.item("PATHS").Selected = True
        tv_preferences.Nodes(9).Selected = True
        fraPATHS.Enabled = True
    Else
        ' disable paths
        fraPATHS.Enabled = False
    End If
    printParams
End Sub

Private Sub chkItalic_Click()
    SetStyle
End Sub

Private Sub chkLineNumbering_Click()
    csPreview.LineNumbering = IIf(chkLineNumbering.Value = 1, True, False)
End Sub

Private Sub chkMSDOS_Click()
    printParams
End Sub

Private Sub chkNormalizeCase_Click()
    csPreview.NormalizeCase = IIf(chkNormalizeCase.Value = 1, True, False)
End Sub

Private Sub chkPIOnlyConsHeader_Click()
    chkPIShowCons.Enabled = IIf(chkPIOnlyConsHeader, False, True)
End Sub

Private Sub chkPIOnlyGlobalHeader_Click()
    chkPIShowGlobals.Enabled = IIf(chkPIOnlyGlobalHeader, False, True)
End Sub

Private Sub chkPIOnlyLocalHeader_Click()
    chkPILocals.Enabled = IIf(chkPIOnlyLocalHeader, False, True)
End Sub

Private Sub chkSaveFiles_Click()
    Dim en As Boolean
    en = IIf(chkSaveFiles.Value <> vbChecked, False, True)
    opCurrentFile.Enabled = en
    opProjectFiles.Enabled = en
    opAllOpenedFiles.Enabled = en
End Sub

Private Sub chkShowIntelliSense_Click()
    lblISLangDefined.Enabled = chkShowIntelliSense
    lblISUserDefined.Enabled = chkShowIntelliSense
    chkISFuncs.Enabled = chkShowIntelliSense
    chkISLangCons.Enabled = chkShowIntelliSense
    chkISUserCons.Enabled = chkShowIntelliSense
    chkISUserFuncs.Enabled = chkShowIntelliSense
    chkISUserProcs.Enabled = chkShowIntelliSense
    chkISUserVars.Enabled = chkShowIntelliSense
    chkISVars.Enabled = chkShowIntelliSense
    lblIntelliSenseSensitive.Enabled = chkShowIntelliSense
    txtIntelliSenseSensitive.Enabled = chkShowIntelliSense
End Sub

Private Sub chkSmoothScrolling_Click()
    csPreview.SmoothScrolling = IIf(chkSmoothScrolling.Value = 1, True, False)
End Sub

Private Sub chkStub_Click()
    printParams
End Sub

Private Sub chkWhiteSpaces_Click()
    csPreview.DisplayWhitespace = IIf(chkWhiteSpaces.Value = 1, True, False)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearCommandHistory_Click()
    frmMSDOSCommand.clearCommandHistory
    frmMSDOSCommand.saveCommandHistory
End Sub

Private Sub cmdClearRecents_Click()
    Dim i As Integer
    Dim A As textStream
    Set A = FSO.CreateTextFile(App.Path & "\Conf\recent.ini", True)
    A.Close
    
    For i = 1 To 10
        frmMain.cMenu.RemoveItem frmMain.cMenu.IndexForKey("mnuRecFile" & i)
        frmMain.cMenu.RemoveItem frmMain.cMenu.IndexForKey("mnuRecProj" & i)
    Next i

    'LoadRecents
End Sub

Private Sub cmdExplore_Click(Index As Integer)
    Dim s As String
    
    c.hwndOwner = Me.Hwnd
    c.InitialDir = App.Path
    c.FileSystemOnly = True
    c.StatusText = True
    c.UseNewUI = True
    s = c.BrowseForFolder
    If Len(s) > 0 Then
        txtCompilerPath(Index).text = s
    End If
End Sub

Private Sub cmdOk_Click()
    SaveConf
    Unload Me
End Sub

Private Sub cmdPATHSAdd_Click()
    Dim strPath As String
    strPath = c.BrowseForFolder
    If strPath <> "" Then
        lstPATHS.AddItem (strPath)
    End If
End Sub

Private Sub cmdPATHSRemove_Click()
    If lstPATHS.SelCount > 0 Then
        lstPATHS.RemoveItem (lstPATHS.ListIndex)
    End If
End Sub

Private Sub cmdPATHSRemoveAll_Click()
    lstPATHS.Clear
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
    Image1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")
    
    cmbColor.AddItem "Rose"
    cmbColor.AddItem "Grey"
    cmbColor.AddItem "Red"
    cmbColor.AddItem "Orange"
    cmbColor.AddItem "Aquamarine"
    cmbColor.AddItem "Dawn"
    cmbColor.AddItem "Nature"
    cmbColor.AddItem "Onyx"
    cmbColor.AddItem "Night"
    cmbColor.AddItem "Emerald"

    Set c = New cBrowseForFolder
'    PlaceControls
'    LoadConf

    Set m_flat = New cFlatControl
    m_flat.Attach picPredefSets
    Set mnuPreSets = New cMenus
    mnuPreSets.CreateFromNothing Me.Hwnd

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
        Set nTab = .Tabs.Add("OPTIONS", , "Compilation Options")
        nTab.Panel = picCompilerOptions
        Set nTab = .Tabs.Add("PATHS", , "CompilerPaths")
        nTab.Panel = picCompilerPaths
        'Set nTab = .Tabs.Add("TOOLS", , "Tools")
        'nTab.Panel = picUserTools
        Set nTab = .Tabs.Add("FILEASSOCIATION", , "File Association")
        nTab.Panel = picFileAsoc
        Set nTab = .Tabs.Add("PI", , "Program Inspector")
        nTab.Panel = picProgramInspector
        Set nTab = .Tabs.Add("IS", , "IntelliSense")
        nTab.Panel = picIntelliSense
        Set nTab = .Tabs.Add("MISC", , "Misc")
        nTab.Panel = picMisc
    End With

    'Editor Conf
    ListStyles
    csPreview.Language = "Fenix"
    csPreview.OpenFile (App.Path & "\resources\txtPreview.txt")
    csPreview.READONLY = True
    RefreshEditorConfigControls
    optFenixBennu_Click (1)
    printParams

    'File association
    With trFiles
        .CheckBoxes = True
        .Nodes.Add(, , "prg", "PRG - Source files").Checked = FileAssociated(".prg", "Bennu/Fenix.Source")
        .Nodes.Add(, , "map", "MAP - Fenix image files").Checked = FileAssociated(".map", "Bennu/Fenix.ImageFile")
        .Nodes.Add(, , "fbp", "FBP - FlameBird Project files").Checked = FileAssociated(".fbp", "FlameBird.Project")
        .Nodes.Add(, , "bmk", "BMK - Source bookmark files").Checked = FileAssociated(".bmk", "FlameBird.Bookmark")
        .Nodes.Add(, , "cpt", "CPT - Map control-point list files").Checked = FileAssociated(".cpt", "FlameBird.ControlPoint")
    End With
    chkDcb.Value = Abs(CInt(FileAssociated(".dcb", "Bennu/Fenix.Bin")))

    'TreeView
    With tv_preferences
        .Nodes.Add , etvwFirst, "Global", "Global"
            tv_preferences.Nodes(1).AddChildNode "GlobalFile", "File"
            tv_preferences.Nodes(1).AddChildNode "GlobalMisc", "Misc"
            tv_preferences.Nodes(1).ShowPlusMinus = True
            tv_preferences.Nodes(1).expanded = True
        .Nodes.Add , etvwNext, "Editor", "Editor"
            tv_preferences.Nodes(4).AddChildNode "EditorColors", "Colors"
            tv_preferences.Nodes(4).AddChildNode "EditorIntelliSense", "IntelliSense"
            tv_preferences.Nodes(4).ShowPlusMinus = True
            tv_preferences.Nodes(4).expanded = True
        .Nodes.Add , etvwNext, "Compiler", "Compiler"
            tv_preferences.Nodes(7).AddChildNode "CompilerOptions", "Options"
            tv_preferences.Nodes(7).AddChildNode "CompilerPaths", "Paths"
            tv_preferences.Nodes(7).ShowPlusMinus = True
            tv_preferences.Nodes(7).expanded = True
        .Nodes.Add , etvwNext, "ProgramInspector", "Program Inspector"
        .Style = etvwTreelinesPlusMinusPictureText
        .LineStyle = etvwRootLines
    End With
    
    PlaceControls
    LoadConf
    
    tv_preferences.Nodes(1).Selected = True
    tabCategories.Tabs.item("GLOBAL").Selected = True
    
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
            chkBold.Value = IIf(csPreview.GetFontStyle(Style.cmStyleItem) And cmFontBold, 1, 0)
            chkItalic.Value = IIf(csPreview.GetFontStyle(Style.cmStyleItem) And cmFontItalic, 1, 0)
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

Private Sub mnuPreSets_Click(ByVal Index As Long)
    Dim sKey As String
    Dim sTitle As String

    sKey = mnuPreSets.ItemKey(Index)
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

Private Sub optFenixBennu_Click(Index As Integer)
    If Index = 0 Then       ' Fenix
        chkDebug.Enabled = True
        chkAutoDeclare.Enabled = False
        chkStub.Enabled = False
        chkMSDOS.Enabled = False
        chkDebugDCB.Enabled = False
        chkDirs.Enabled = False
        fraPATHS.Enabled = False
        txtCompilerPath(0).Enabled = True
        cmdExplore(0).Enabled = True
        txtCompilerPath(1).Enabled = False
        cmdExplore(1).Enabled = False
    Else                    ' Bennu
        chkDebug.Enabled = True
        chkAutoDeclare.Enabled = True
        chkStub.Enabled = True
        chkMSDOS.Enabled = True
        chkDebugDCB.Enabled = True
        chkDirs.Enabled = True
        If chkDirs.Value Then
            fraPATHS.Enabled = True
        Else
            fraPATHS.Enabled = False
        End If
        txtCompilerPath(0).Enabled = False
        cmdExplore(0).Enabled = False
        txtCompilerPath(1).Enabled = True
        cmdExplore(1).Enabled = True
    End If
    printParams
End Sub

Private Sub optHelpLine_Click(Index As Integer)
    If Index = 0 Then
        G_ProcHelpLine = 1
    ElseIf Index = 1 Then
        G_ProcHelpLine = 0
    Else
        G_ProcHelpLine = -1
    End If
End Sub

Private Sub picPredefSets_Click()
    Dim i As Integer
    Dim folder As folder
    Dim file As file

    Set mnuPreSets = Nothing
    Set mnuPreSets = New cMenus
    mnuPreSets.CreateFromNothing Me.Hwnd

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

Private Sub tv_preferences_NodeClick(node As vbalTreeViewLib6.cTreeViewNode)
    Select Case node.Key
    Case "Global":
        tabCategories.Tabs.item("GLOBAL").Selected = True
        Case "GlobalFile":
            tabCategories.Tabs.item("FILEASSOCIATION").Selected = True
        Case "GlobalMisc":
            tabCategories.Tabs.item("MISC").Selected = True
    Case "Editor":
        tabCategories.Tabs.item("EDITOR").Selected = True
        Case "EditorColors":
            tabCategories.Tabs.item("COLORS").Selected = True
        Case "EditorIntelliSense":
            tabCategories.Tabs.item("IS").Selected = True
    Case "Compiler":
        tabCategories.Tabs.item("COMPILATION").Selected = True
        Case "CompilerOptions":
            tabCategories.Tabs.item("OPTIONS").Selected = True
        Case "CompilerPaths":
            tabCategories.Tabs.item("PATHS").Selected = True
    Case "ProgramInspector":
        tabCategories.Tabs.item("PI").Selected = True
    End Select
End Sub

Private Sub txtIntelliSenseSensitive_Validate(Cancel As Boolean)
    Cancel = True
    If IsNumeric(txtIntelliSenseSensitive.text) Then
        If CLng(txtIntelliSenseSensitive.text) < 1 Or 4 < CLng(txtIntelliSenseSensitive.text) Then
            MsgBox "IntelliSense sensitive must be between 1 and 4."
        Else
            Cancel = False
        End If
    End If
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

Private Sub printParams()
    Dim sText As String
    
    If chkDebug.Value Then
        sText = sText & " -d"
    End If
    
    If optFenixBennu.item(0).Value Then     ' Fenix

    Else                                    ' Bennu
        If chkStub.Value Then
            sText = sText & " -s bgdi.exe"
        End If
        If chkMSDOS.Value Then
            sText = sText & " -c"
        End If
        If chkAutoDeclare.Value Then
            sText = sText & " -Ca"
        End If
        If chkDebugDCB.Value Then
            sText = sText & " -g"
        End If
        If chkDirs.Value Then
            sText = sText & " -i"
            ' and all that comes after...
        End If
    End If
    txtParams.text = sText
End Sub
