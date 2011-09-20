VERSION 5.00
Begin VB.Form frmCodeWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Necromancer"
   ClientHeight    =   5850
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Wizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Header creation"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   0
      TabIndex        =   44
      Tag             =   "1000"
      Top             =   840
      Width           =   7155
      Begin VB.CheckBox chkGNU 
         Caption         =   "GNU license"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Check it to add GNU license to your code. If you want to know more about this license look in http://www.gnu.org/licenses/gpl.txt "
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdDataToday 
         Caption         =   "Today"
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         ToolTipText     =   "Click here to write todays date in the text box."
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtCompany 
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Text            =   "YouCompany"
         ToolTipText     =   "Enter here the name of your company."
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         ToolTipText     =   "Date of creation."
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtDevelopers 
         Height          =   885
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   33
         Text            =   "Wizard.frx":0442
         ToolTipText     =   "Here goes the list of names that have created the game."
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtGameTitle 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Text            =   "MyGame"
         ToolTipText     =   "Enter here the name of the game, the game title."
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkStep0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ignore Header"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   38
         Tag             =   "1002"
         Top             =   4080
         Width           =   4890
      End
      Begin VB.Label lblCompany 
         Caption         =   "Conpany:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblData 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblDevelopers 
         Caption         =   "Developers:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblGameTitle 
         Caption         =   "Game title:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblStep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Wizard.frx":046C
         ForeColor       =   &H00000000&
         Height          =   1935
         Index           =   0
         Left            =   3720
         TabIndex        =   45
         Tag             =   "1001"
         Top             =   120
         Width           =   3360
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Video Mode"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   1
      Left            =   10000
      TabIndex        =   46
      Tag             =   "2000"
      Top             =   840
      Width           =   7155
      Begin VB.CheckBox chkStep1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ignore this step"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtSkippedFPS 
         Alignment       =   1  'Right Justify
         DataField       =   "0"
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "How many frames will be jumped in the case that the computer cann't handle this frame rate. "
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtFPS 
         Alignment       =   1  'Right Justify
         DataField       =   "0"
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "25"
         ToolTipText     =   $"Wizard.frx":053F
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cmbScalingFilter 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   $"Wizard.frx":05DF
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbShowingMode 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   $"Wizard.frx":0673
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbBPP 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "The colour deep of the video mode. 8 bit per pixel (BPP) uses paletted colours to show graphics. 16 BPP uses not-paletted colours."
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cmbVideoMode 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   $"Wizard.frx":0717
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblSkippedFPS 
         Caption         =   "permited frames skkiped"
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblFPS 
         Caption         =   "Frames per Second:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblScaleFilter 
         Caption         =   "Scaling filter:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblShowingMode 
         Caption         =   "Showing mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblBPP 
         Caption         =   "Bpp:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblVideoMode 
         Caption         =   "Video mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Wizard.frx":07EA
         ForeColor       =   &H80000008&
         Height          =   1950
         Index           =   1
         Left            =   3720
         TabIndex        =   47
         Tag             =   "2001"
         Top             =   90
         Width           =   3360
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Intro & Logos"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   2
      Left            =   10000
      TabIndex        =   48
      Tag             =   "2002"
      Top             =   840
      Width           =   7155
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   5040
         TabIndex        =   21
         ToolTipText     =   "Creates a new cut. This cut will be added at the end of the list. Click this button if you want to add a new cut."
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbIntroList 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Wizard.frx":087C
         Left            =   240
         List            =   "Wizard.frx":087E
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Logo"
         Height          =   2055
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   6855
         Begin VB.TextBox txtStepTitle 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   23
            ToolTipText     =   "A name to recognize the cut."
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtTransTime 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            TabIndex        =   30
            Text            =   "1"
            ToolTipText     =   "Duration of the cut."
            Top             =   1680
            Width           =   735
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5880
            TabIndex        =   22
            ToolTipText     =   "Deletes the current cut from the list. Be care, cause the cut deleted will not be restored. "
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox cmbTransition 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   $"Wizard.frx":0880
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CommandButton cmdSelectMusic 
            Caption         =   "Select"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5040
            TabIndex        =   28
            ToolTipText     =   $"Wizard.frx":0909
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtMusic 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   27
            ToolTipText     =   "The sound effect or music that will be played during the cut. If ""Music"" checkbox is clicked, this must have a file."
            Top             =   1200
            Width           =   3975
         End
         Begin VB.CheckBox chkMusic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Music"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Check this button if you need to add some sound effects or play a music during this cut."
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmdSelectPicture 
            Caption         =   "Select"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5040
            TabIndex        =   25
            ToolTipText     =   $"Wizard.frx":09B3
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtPicture 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   24
            ToolTipText     =   "The picture that will be displayed in the background of the cut. "
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label lblStepTitle 
            Caption         =   "Step title:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblTransSeconds 
            Caption         =   "seconds"
            Enabled         =   0   'False
            Height          =   255
            Left            =   5640
            TabIndex        =   75
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblTransTime 
            Caption         =   "Transition time:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   74
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblTransition 
            Caption         =   "Transition:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblPicture 
            Caption         =   "Picture:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.CheckBox chkStep2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ignore this step"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intro: /Logos: /Title:"
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   2
         Left            =   3480
         TabIndex        =   49
         Tag             =   "2003"
         Top             =   120
         Width           =   3600
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Menu"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   3
      Left            =   -10000
      TabIndex        =   50
      Tag             =   "2004"
      Top             =   840
      Visible         =   0   'False
      Width           =   7155
      Begin VB.CheckBox chkClickMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Click sound"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Check it if you want to hear mouse clicks during the menu when any options is clicked."
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectMusicMenu 
         Caption         =   "Select song"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         ToolTipText     =   "CLick this button to add music or sound effect to the menu."
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtMusicMenu 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "The music or sound effect played during the menu."
         Top             =   2760
         Width           =   4215
      End
      Begin VB.CheckBox chkRepeatMusicMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Repeat"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   $"Wizard.frx":0A51
         Top             =   3120
         Width           =   975
      End
      Begin VB.CheckBox chkMusicMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Music"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Check it if you want to add music (or some sound effect) to the menu."
         Top             =   2760
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Summary of the menu"
         Height          =   2535
         Left            =   120
         TabIndex        =   66
         Top             =   0
         Width           =   3495
         Begin VB.CheckBox chkNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "New"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   0
            ToolTipText     =   "Check it to add the elemente New (start a new game) to the menu."
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkExit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Exit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox chkCredits 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Credits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox chkOptions 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Options"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Check it to add Options element to the menu."
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chkPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Enter Password"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Adds the password option ti the menu."
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkLoad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Load"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Check this box to add Load option to the menu."
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkSave 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Save"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Check it to add Save element to the menu. "
            Top             =   480
            Width           =   1935
         End
         Begin VB.ComboBox cmbAligment 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   2160
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblAligment 
            Caption         =   "Aligment:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "New"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   65
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkStep3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Ignore this step"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Menu:"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   3
         Left            =   3720
         TabIndex        =   51
         Tag             =   "2005"
         Top             =   120
         Width           =   3360
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finishing - Summary"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   4
      Left            =   10000
      TabIndex        =   52
      Tag             =   "3000"
      Top             =   840
      Width           =   7155
      Begin VB.Frame fraCode 
         Caption         =   "Code  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   0
         TabIndex        =   72
         Top             =   960
         Width           =   6975
         Begin VB.TextBox txtCode 
            Height          =   2655
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   240
            Width           =   6735
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"Wizard.frx":0ADE
         Height          =   735
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   6855
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
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
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7155
      TabIndex        =   43
      Top             =   5280
      Width           =   7155
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   315
         Index           =   4
         Left            =   5790
         MaskColor       =   &H00000000&
         TabIndex        =   42
         Tag             =   "104"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   315
         Index           =   3
         Left            =   3720
         MaskColor       =   &H00000000&
         TabIndex        =   41
         Tag             =   "103"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Previous"
         Height          =   315
         Index           =   2
         Left            =   2400
         MaskColor       =   &H00000000&
         TabIndex        =   40
         Tag             =   "102"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Index           =   1
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   39
         Tag             =   "101"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   120
         X2              =   7024
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   105
         X2              =   7009
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   120
      Picture         =   "Wizard.frx":0BA1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Wizard.frx":0FE3
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
      Left            =   960
      TabIndex        =   54
      Top             =   240
      Width           =   5805
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New file code asistant"
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
      Left            =   840
      TabIndex        =   53
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   -1680
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmCodeWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Necromancer for FBmX
'Copyright (C) 2007
'
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
'GNU General Public License for more details.Option Explicit


Const NUM_STEPS = 5

Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_FINISH = 4

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Necromancer"
Const INTRO_KEY = "IntroductionScreen"
Const SHOW_INTRO = "ShowIntro"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

Const APP_CATEGORY = "Wizards"
Const WIZARD_NAME = "WizardTemplate"
Const CONFIRM_KEY = "ConfirmScreen"


'variables a nivel de módulo
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

'Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean


'comboBox filling data
Const NUM_VIDEO_MODES = 10
Const NUM_SCALE_MODES = 5
Const NUM_BPP_MODES = 2
Const NUM_SCREEN_MODES = 2
Const NUM_TRANSITION_MODES = 3
Const NUM_ALIGMENT_MODES = 9


Dim videoModes(NUM_VIDEO_MODES) As String
Dim scaleModes(NUM_SCALE_MODES) As String
Dim bppModes(NUM_BPP_MODES) As String
Dim screenModes(NUM_SCREEN_MODES) As String
Dim transitionModes(NUM_TRANSITION_MODES) As String
Dim aligmentModes(NUM_ALIGMENT_MODES) As String

Private Type T_IL
    Title As String
    pictureFile As String
    hasMusic As Boolean
    musicFile As String
    transType As Integer
    transTime As Integer
End Type

Dim ilList() As T_IL            ' list of intro and logo elements
Dim curElem As Integer          ' the current element

Public sCode As String             ' the code generated after the selections we've made in this form


Private Sub chkMusic_Click()
    If chkMusic.Value = 1 Then
        txtMusic.Enabled = True
        cmdSelectMusic.Enabled = True
        ilList(curElem).hasMusic = True
    Else
        txtMusic.Enabled = False
        cmdSelectMusic.Enabled = False
        ilList(curElem).hasMusic = False
    End If
    printList
End Sub

Private Sub chkMusicMenu_Click()
    If chkMusicMenu.Value = 1 Then
        txtMusicMenu.Enabled = True
        cmdSelectMusicMenu.Enabled = True
    Else
        txtMusicMenu.Enabled = False
        cmdSelectMusicMenu.Enabled = False
    End If
End Sub

Private Sub chkStep0_Click()
    'disable/enable the controlls in this frame
    If chkStep0.Value = 1 Then
        lblGameTitle.Enabled = False
        lblDevelopers.Enabled = False
        lblCompany.Enabled = False
        lblData.Enabled = False
        txtGameTitle.Enabled = False
        txtDevelopers.Enabled = False
        txtCompany.Enabled = False
        txtData.Enabled = False
        cmdDataToday.Enabled = False
        chkGNU.Enabled = False
    Else
        lblGameTitle.Enabled = True
        lblDevelopers.Enabled = True
        lblCompany.Enabled = True
        lblData.Enabled = True
        txtGameTitle.Enabled = True
        txtDevelopers.Enabled = True
        txtCompany.Enabled = True
        txtData.Enabled = True
        cmdDataToday.Enabled = True
        chkGNU.Enabled = True
    End If
    
End Sub

Private Sub chkStep1_Click()
    If chkStep1.Value = 1 Then
        lblVideoMode.Enabled = False
        lblBPP.Enabled = False
        lblScaleFilter.Enabled = False
        lblShowingMode.Enabled = False
        cmbVideoMode.Enabled = False
        cmbBPP.Enabled = False
        cmbScalingFilter.Enabled = False
        cmbShowingMode.Enabled = False
        lblFPS.Enabled = False
        lblSkippedFPS.Enabled = False
        txtFPS.Enabled = False
        txtSkippedFPS.Enabled = False
    Else
        lblVideoMode.Enabled = True
        lblBPP.Enabled = True
        lblScaleFilter.Enabled = True
        lblShowingMode.Enabled = True
        cmbVideoMode.Enabled = True
        cmbBPP.Enabled = True
        cmbScalingFilter.Enabled = True
        cmbShowingMode.Enabled = True
        lblFPS.Enabled = True
        lblSkippedFPS.Enabled = True
        txtFPS.Enabled = True
        txtSkippedFPS.Enabled = True
    End If
End Sub

Private Sub chkStep2_Click()
    If chkStep2.Value = 1 Then
        cmbIntroList.Enabled = False
        lblStepTitle.Enabled = False
        txtStepTitle.Enabled = False
        lblPicture.Enabled = False
        txtPicture.Enabled = False
        cmdSelectPicture.Enabled = False
        chkMusic.Enabled = False
        txtMusic.Enabled = False
        cmdSelectMusic.Enabled = False
        lblTransition.Enabled = False
        cmbTransition.Enabled = False
'        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        lblTransTime.Enabled = False
        txtTransTime.Enabled = False
        lblTransSeconds.Enabled = False
    Else
        cmbIntroList.Enabled = True
        lblStepTitle.Enabled = False
        txtStepTitle.Enabled = False
        lblPicture.Enabled = True
        txtPicture.Enabled = True
        cmdSelectPicture.Enabled = True
        chkMusic.Enabled = True
        txtMusic.Enabled = True
        cmdSelectMusic.Enabled = True
        lblTransition.Enabled = True
        cmbTransition.Enabled = True
'        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        lblTransTime.Enabled = True
        txtTransTime.Enabled = True
        lblTransSeconds.Enabled = True
    End If
End Sub

Private Sub chkStep3_Click()
    If chkStep3.Value = 1 Then
        chkNew.Enabled = False
        chkSave.Enabled = False
        chkLoad.Enabled = False
        chkPassword.Enabled = False
        chkOptions.Enabled = False
        chkCredits.Enabled = False
        chkExit.Enabled = False
        lblAligment.Enabled = False
        cmbAligment.Enabled = False
        chkMusicMenu.Enabled = False
        txtMusicMenu.Enabled = False
        cmdSelectMusicMenu.Enabled = False
        chkRepeatMusicMenu.Enabled = False
        chkMusicMenu.Enabled = False
        chkClickMenu.Enabled = False
    Else
        chkNew.Enabled = True
        chkSave.Enabled = True
        chkLoad.Enabled = True
        chkPassword.Enabled = True
        chkOptions.Enabled = True
        chkCredits.Enabled = True
        chkExit.Enabled = True
        lblAligment.Enabled = True
        cmbAligment.Enabled = True
        chkMusicMenu.Enabled = True
        txtMusicMenu.Enabled = True
        cmdSelectMusicMenu.Enabled = True
        chkRepeatMusicMenu.Enabled = True
        chkMusicMenu.Enabled = True
        chkClickMenu.Enabled = True
    End If
End Sub

Private Sub cmbIntroList_Click()
    'MsgBox "Selected element: " & cmbIntroList.ListIndex & " from " & cmbIntroList.ListCount
    curElem = cmbIntroList.ListIndex
    ' put the data in its each control
    With ilList(curElem)
        txtStepTitle.text = .Title
        txtPicture.text = .pictureFile
        If .hasMusic Then
            chkMusic.Value = 1
        Else
            chkMusic.Value = 0
        End If
        txtMusic.text = .musicFile
        cmbTransition.ListIndex = .transType + 1
        txtTransTime.text = .transTime
    End With
End Sub

Private Sub cmbTransition_Click()
    If fraStep(2).Left = 0 Then
        ilList(curElem).transType = cmbTransition.ListIndex + 1
        'MsgBox transitionModes(ilList(curElem).transType)
        printList
    End If
End Sub

'Private Sub cmdAdd_Click()
'    cmbIntroList.AddItem (txtStepTitle.text)
'    cmbIntroList.ListIndex = cmbIntroList.NewIndex
'    lstIntroList.AddItem (txtStepTitle.text)
'    lstIntroList.ListIndex = lstIntroList.NewIndex
'    ReDim Preserve ilList(cmbIntroList.ListCount) As T_IL
'    printList
'End Sub


Private Sub cmdDataToday_Click()
    txtData.text = Date
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer
    With cmbIntroList
        '.RemoveItem (.ListIndex)
        .RemoveItem (curElem)
    
        printList
        ' first, move all the items to the correct position; then
        'For i = cmbIntroList.ListIndex To cmbIntroList.ListCount - 1
        For i = curElem To cmbIntroList.ListCount - 1
            ilList(i).Title = ilList(i + 1).Title
            ilList(i).pictureFile = ilList(i + 1).pictureFile
            ilList(i).hasMusic = ilList(i + 1).hasMusic
            ilList(i).musicFile = ilList(i + 1).musicFile
            ilList(i).transType = ilList(i + 1).transType
            ilList(i).transTime = ilList(i + 1).transTime
        Next i
        ReDim Preserve ilList(cmbIntroList.ListCount - 1) As T_IL
        printList
        
        If .ListCount > 0 Then
            '.ListIndex = 0
            curElem = 0
        Else    ' num of elements in cmbIntroList = 0, disable controls
            'cmdAdd.Enabled = False
            cmdDel.Enabled = False
            lblStepTitle.Enabled = False
            txtStepTitle.Enabled = False
            lblPicture.Enabled = False
            txtPicture.Enabled = False
            cmdSelectPicture.Enabled = False
            chkMusic.Enabled = False
            txtMusic.Enabled = False
            cmdSelectMusic.Enabled = False
            lblTransition.Enabled = False
            cmbTransition.Enabled = False
            lblTransTime.Enabled = False
            txtTransTime.Enabled = False
            lblTransSeconds.Enabled = False
        End If
        
    End With
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    
    Select Case Index
'        Case BTN_HELP
'            mbHelpStarted = True
'            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
'            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        
        Case BTN_CANCEL
            sCode = ""
            Unload Me
            frmNewFile.Show
          
        Case BTN_BACK
            'colocar aquí casos especiales para saltar
            'a pasos alternativos
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'colocar aquí casos especiales para saltar
            If mnCurStep = STEP_2 Then
                If checkIntroAndLogo = -1 Then Exit Sub
            End If
            'a pasos alternativos
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
            'el código de creación de asistentes va aquí
            Unload Me
            NewFileForm FF_SOURCE

            
            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
                'frmConfirm.Show vbModal
            End If
        
    End Select
End Sub

Private Sub cmdNew_Click()
    Dim Title As String
    With cmbIntroList
begin:
        Title = InputBox("Type the title for this cut." & vbCrLf & "This title is a recordatory of the cut and must be something like MyLogo, 1_Scene, 2_Scene...", "Enter title name")
        If Title <> "" Then
            cmdDel.Enabled = True
            lblStepTitle.Enabled = True
            txtStepTitle.Enabled = True
            lblPicture.Enabled = True
            txtPicture.Enabled = True
            cmdSelectPicture.Enabled = True
            chkMusic.Enabled = True
            'txtMusic.Enabled = True
            cmdSelectMusic.Enabled = True
            lblTransition.Enabled = True
            cmbTransition.Enabled = True
            lblTransTime.Enabled = True
            txtTransTime.Enabled = True
            lblTransSeconds.Enabled = True
            .Enabled = True
            ReDim Preserve ilList(.ListCount) As T_IL
            .AddItem (Title)
            txtTransTime.text = "1"
            ilList(curElem).transTime = 1
            .ListIndex = .newIndex
            'frmWizard.Caption = "List contains " & .ListCount + 1 & " elements." & vbCrLf & "Selected element: " & .ListIndex
            curElem = .newIndex
            'ReDim Preserve ilList(.ListCount - 1) As T_IL
            txtStepTitle.text = Title
            txtStepTitle.SetFocus
            ilList(curElem).Title = Title
            ilList(curElem).transType = 0

            printList
        Else
            MsgBox "Title must contain a name"
            GoTo begin
        End If
    End With
End Sub

Private Sub cmdSelectMusic_Click()
    Dim sFile  As String

    sFile = OpenDialog(Me, "WAV (*.wav)|*.wav|OGG Vorbis (*.ogg)|*.ogg|All files (*.*)|*.*", "Select song", App.Path)
    
    If sFile <> "" Then
        'If FileExists(sFile) Then
            txtMusic.text = sFile
            ilList(curElem).musicFile = txtMusic.text
        'End If
    End If
    printList
End Sub

Private Sub cmdSelectMusicMenu_Click()
    
    Dim sFile  As String

    sFile = OpenDialog(Me, "WAV (*.wav)|*.wav|OGG Vorbis (*.ogg)|*.ogg|All files (*.*)|*.*", "Select song for the menu", App.Path)
    
    If sFile <> "" Then
        'If FileExists(sFile) Then
            txtMusicMenu.text = sFile
        'End If
    End If
    
End Sub

Private Sub cmdSelectPicture_Click()
    Dim sFile  As String

    sFile = OpenDialog(Me, "Portable Graphics Network (*.png)|*.png|Fenix bitmap (*.fbm)|*.fbm|DIV bitmap (*.map)|*.map|All files (*.*)|*.*", "Select song for the menu", App.Path)
    
    If sFile <> "" Then
        'If FileExists(sFile) Then
            txtPicture.text = sFile
            ilList(curElem).pictureFile = txtPicture.text
        'End If
    End If
    printList
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Image1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")
    
    loadData

    'initialize all the variables
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Load all the strings to the form
    'LoadResStrings Me
    
    initData
    
    cmdDataToday_Click
    
    'Determinate the first step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        chkStep0.Value = vbChecked
        SetStep 1, DIR_NEXT
    Else
        SetStep 0, DIR_NONE
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdNav_Click (BTN_CANCEL)
    End If
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO
      
        Case STEP_1
      
        Case STEP_2 ' INTRO and LOGO
            
        Case STEP_3 ' MENU
            mbFinishOK = False
            cmdNav(4).Enabled = False
      
        Case STEP_FINISH
            mbFinishOK = True
            cmdNav(4).Enabled = True
            writeCode
        
    End Select
    
    'pasar al siguiente paso
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'esta subrutina muestra un mensaje de error cuando el
'usuario no ha escrito suficientes datos para continuar
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'obtener el mensaje de error de base
    'sTmp = LoadResString(RES_ERROR_MSG)
    'obtener el mensaje específico
    'sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'ver si hay que guardar la configuración
    If chkSaveSettings.Value = vbChecked Then
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
    End If
  
    'If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

Private Sub lstIntroList_DblClick()
    'lstIntroList.text = InputBox("Title")
    'lstIntroList.ItemData(lstIntroList.ListIndex) = InputBox("Title")
End Sub

Private Sub txtCompany_GotFocus()
    txtCompany.SelStart = 0
    txtCompany.SelLength = Len(txtCompany.text)
End Sub

Private Sub txtDevelopers_GotFocus()
    txtDevelopers.SelStart = 0
    txtDevelopers.SelLength = Len(txtDevelopers.text)
End Sub

Private Sub txtFPS_Change()
    txtFPS.SelStart = 0
    txtFPS.SelLength = Len(txtFPS.text)
End Sub

Private Sub txtGameTitle_Change()
    txtGameTitle.SelStart = 0
    txtGameTitle.SelLength = Len(txtGameTitle.text)
End Sub

Private Sub txtMusic_Change()
    ilList(curElem).musicFile = txtMusic.text
End Sub

Private Sub txtPicture_Change()
   ilList(curElem).pictureFile = txtPicture.text
End Sub

Private Sub txtSkippedFPS_Change()
    txtSkippedFPS.SelStart = 0
    txtSkippedFPS.SelLength = Len(txtSkippedFPS.text)
End Sub

Private Sub txtStepTitle_Change()
    'check if enter key is down
'    If cmbIntroList.ListIndex <> -1 Then
'        ilList(curElem).Title = txtStepTitle.text
'        cmbIntroList.text = txtStepTitle.text
'    Else
'        MsgBox "there's no element selected"
'    End If
End Sub

Private Sub txtStepTitle_GotFocus()
    txtStepTitle.SelStart = 0
    txtStepTitle.SelLength = Len(txtStepTitle.text)
End Sub

Private Sub txtStepTitle_LostFocus()
    'check if enter key is down
    If cmbIntroList.ListIndex <> -1 Then
        ilList(curElem).Title = txtStepTitle.text
        cmbIntroList.text = txtStepTitle.text
    Else
        MsgBox "there's no element selected"
    End If
End Sub

Private Sub txtTransTime_Change()
    If IsNumeric(txtTransTime.text) Then
        If CInt(txtTransTime.text) < 0 Then
            MsgBox "Time must be positive"
        ElseIf CInt(txtTransTime.text) = 0 Then
            MsgBox "Time must be higher than zero seconds"
        ElseIf CInt(txtTransTime.text) > 10 Then
            MsgBox "Too much long time"
        Else
            ilList(curElem).transTime = CInt(txtTransTime.text)
            printList
        End If
    Else
        MsgBox "Value must be number"
    End If
End Sub


Private Sub writeCode()
    
    sCode = ""

    ' *********** Program ****************
    wProgram
    
    ' *********** Global *****************
    wGlobal
    
    ' *********** Begin ******************
    wMainBegin
    
    ' *********** Intro & Logos **********
    wIntroAndLogos
    
    ' *********** Menu *******************
    wMenu
    
    txtCode.text = sCode
    
End Sub


Function wMouseProc()

    Dim strMouseProc As String
    
    strMouseProc = "" & vbCrLf
    strMouseProc = strMouseProc & "//-----------------------------------------------" & vbCrLf
    strMouseProc = strMouseProc & "PROCESS cursor()" & vbCrLf
    If chkClickMenu.Value = 1 Then
        strMouseProc = strMouseProc & "PRIVATE" & vbCrLf
        strMouseProc = strMouseProc & "    int s_sound;" & vbCrLf
    End If
    strMouseProc = strMouseProc & "BEGIN" & vbCrLf
    strMouseProc = strMouseProc & "    file = f_menu;" & vbCrLf
    strMouseProc = strMouseProc & "    graph = 1;" & vbCrLf
    If chkClickMenu.Value = 1 Then
        strMouseProc = strMouseProc & "       s_sound = load_wav('click.wav');" & vbCrLf
    End If
    strMouseProc = strMouseProc & "    LOOP" & vbCrLf
    If chkClickMenu.Value = 1 Then
        strMouseProc = strMouseProc & "        if (mouse.left)" & vbCrLf
        strMouseProc = strMouseProc & "             play_wav(s_click,0);" & vbCrLf
        strMouseProc = strMouseProc & "        end" & vbCrLf
    End If
    strMouseProc = strMouseProc & "        x = mouse.x;" & vbCrLf
    strMouseProc = strMouseProc & "        y = mouse.y;" & vbCrLf
    strMouseProc = strMouseProc & "        FRAME(100);" & vbCrLf
    strMouseProc = strMouseProc & "    END" & vbCrLf
    strMouseProc = strMouseProc & "END" & vbCrLf
    
    wMouseProc = strMouseProc
    
End Function

Function wPauseProc()

    Dim strPause As String
    
    strPause = "" & vbCrLf
    strPause = strPause & "//-----------------------------------------------" & vbCrLf
    strPause = strPause & "PROCESS pause()" & vbCrLf
    strPause = strPause & "PRIVATE" & vbCrLf
    strPause = strPause & "    int id_main;" & vbCrLf
    strPause = strPause & "BEGIN" & vbCrLf
    strPause = strPause & "    id_main = get_id(0);" & vbCrLf
    strPause = strPause & "    SIGNAL(id_main,s_freeze_tree);" & vbCrLf
    strPause = strPause & "    SIGNAL(id,s_wakeup);" & vbCrLf
    strPause = strPause & "    REPEAT" & vbCrLf
    strPause = strPause & "        FRAME;" & vbCrLf
    strPause = strPause & "    UNTIL(scan_code <> 0);" & vbCrLf
    strPause = strPause & "    SIGNAL(id_main,s_wakeup_tree);" & vbCrLf
    strPause = strPause & "END" & vbCrLf
    wPauseProc = strPause

End Function

Function wSnapshotProc()

    Dim strSnapshot As String
    
    strSnapshot = "" & vbCrLf
    strSnapshot = strSnapshot & "//-------------------------------------------------------------" & vbCrLf
    strSnapshot = strSnapshot & "PROCESS Snapshot(x_start,y_start,x_end,y_end,path)" & vbCrLf
    strSnapshot = strSnapshot & "PRIVATE" & vbCrLf
    strSnapshot = strSnapshot & "     number_file, id_map, width, heigth;" & vbCrLf
    strSnapshot = strSnapshot & "     string file_video, number;" & vbCrLf
    strSnapshot = strSnapshot & "BEGIN" & vbCrLf
    strSnapshot = strSnapshot & "     width = x_end - x_start;" & vbCrLf
    strSnapshot = strSnapshot & "     heigth = y_end - y_start;" & vbCrLf
    strSnapshot = strSnapshot & "     define_region(31, x_start, y_start, width, heigth);" & vbCrLf
    strSnapshot = strSnapshot & "     id_map = new_map(width, heigth, width/2, heigth/2, 0);" & vbCrLf
    strSnapshot = strSnapshot & "     REPEAT" & vbCrLf
    strSnapshot = strSnapshot & "          number = itoa(number_file);" & vbCrLf
    strSnapshot = strSnapshot & "          SWITCH (strlen(number))" & vbCrLf
    strSnapshot = strSnapshot & "               CASE 1: number = '0000' + number; END" & vbCrLf
    strSnapshot = strSnapshot & "               CASE 2: number = '000' + number; END" & vbCrLf
    strSnapshot = strSnapshot & "               CASE 3: number = '00' + number; END" & vbCrLf
    strSnapshot = strSnapshot & "               CASE 4: number = '0' + number; END" & vbCrLf
    strSnapshot = strSnapshot & "               DEFAULT: number = number; END" & vbCrLf
    strSnapshot = strSnapshot & "          END" & vbCrLf
    strSnapshot = strSnapshot & "          file_video = path + 'snap' + number + '.pcx';" & vbCrLf
    strSnapshot = strSnapshot & "          screen_copy(31, 0, id_map, 0, 0, width, heigth);" & vbCrLf
    strSnapshot = strSnapshot & "          save_pcx(0, id_map, file_video);" & vbCrLf
    strSnapshot = strSnapshot & "          number_file++;" & vbCrLf
    strSnapshot = strSnapshot & "          FRAME;" & vbCrLf
    strSnapshot = strSnapshot & "     UNTIL (KEY(_F5)==TRUE)" & vbCrLf
    strSnapshot = strSnapshot & "     number_file = 0;" & vbCrLf
    strSnapshot = strSnapshot & "END" & vbCrLf
    
    wSnapshotProc = strSnapshot
    
End Function

Private Sub wMenu_NewProc()
    sCode = sCode & "process menu_new(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 100;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the new call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_SaveProc()
    sCode = sCode & "process menu_save(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 101;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the save call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_LoadProc()
    sCode = sCode & "process menu_load(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 102;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the load call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_PasswordProc()
    sCode = sCode & "process menu_password(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 103;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the password call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_CreditsProc()
    sCode = sCode & "process menu_credits(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 104;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the credits call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_OptionsProc()
    sCode = sCode & "process menu_options(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 105;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the credits call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub wMenu_ExitProc()
    sCode = sCode & "process menu_exit(x,y)" & vbCrLf
    sCode = sCode & "begin" & vbCrLf
    sCode = sCode & "    file = f_menu;" & vbCrLf
    sCode = sCode & "    graph = 106;" & vbCrLf
    sCode = sCode & "    loop" & vbCrLf
    sCode = sCode & "        if (collision (TYPE cursor))" & vbCrLf
    sCode = sCode & "            alpha = 100;" & vbCrLf
    sCode = sCode & "            // here goes the exit call" & vbCrLf
    sCode = sCode & "        else" & vbCrLf
    sCode = sCode & "            alpha = 50;" & vbCrLf
    sCode = sCode & "        end" & vbCrLf
    sCode = sCode & "        frame(100);" & vbCrLf
    sCode = sCode & "    end" & vbCrLf
    sCode = sCode & "end" & vbCrLf
End Sub

Private Sub loadData()
    '*************************************************
        videoModes(1) = "m320x200"
        videoModes(2) = "m320x240"
        videoModes(3) = "m320x400"
        videoModes(4) = "m360x240"
        videoModes(5) = "m360x360"
        videoModes(6) = "m376x282"
        videoModes(7) = "m640x400"
        videoModes(8) = "m640x480"
        videoModes(9) = "m800x600"
        videoModes(10) = "m1024x768"
        
        
        scaleModes(1) = "SCALE_NOFILTER"
        scaleModes(2) = "SCALE_SCALE2X"
        scaleModes(3) = "SCALE_HQ2X"
        scaleModes(4) = "SCALE_SCANLINE2X"
        scaleModes(5) = "SCALE_NORMAL2X"
        
        
        bppModes(1) = "8 bpp"
        bppModes(2) = "16 bpp"
        
        
        screenModes(1) = "WINDOWED"
        screenModes(2) = "FULL_SCREEN"
        
        
        transitionModes(1) = "none"
        transitionModes(2) = "fade_off -> fade_on"
        transitionModes(3) = "fade -> fade"
        
        
        aligmentModes(1) = "up left"
        aligmentModes(2) = "up middle"
        aligmentModes(3) = "up right"
        aligmentModes(4) = "middle left"
        aligmentModes(5) = "centered"
        aligmentModes(6) = "middle right"
        aligmentModes(7) = "down left"
        aligmentModes(8) = "down middle"
        aligmentModes(9) = "down right"
 '*************************************************************
End Sub

Private Sub initData()

    Dim i As Integer
    
    For i = 1 To NUM_VIDEO_MODES
        cmbVideoMode.AddItem (videoModes(i))
    Next
    cmbVideoMode.ListIndex = 0
    
    For i = 1 To NUM_SCALE_MODES
        cmbScalingFilter.AddItem (scaleModes(i))
    Next
    cmbScalingFilter.ListIndex = 0
    
    For i = 1 To NUM_BPP_MODES
        cmbBPP.AddItem (bppModes(i))
    Next
    cmbBPP.ListIndex = 0
    
    For i = 1 To NUM_SCREEN_MODES
        cmbShowingMode.AddItem (screenModes(i))
    Next
    cmbShowingMode.ListIndex = 1
    
    For i = 1 To NUM_ALIGMENT_MODES
        cmbAligment.AddItem (aligmentModes(i))
    Next
    cmbAligment.ListIndex = 3
    
    For i = 1 To NUM_TRANSITION_MODES
        cmbTransition.AddItem (transitionModes(i))
    Next
    cmbTransition.ListIndex = 0
End Sub

'======================================================================================
Public Sub wWhiteLines(n As Integer)
    Dim i As Integer
    For i = 1 To n
        sCode = sCode & "" & vbCrLf
    Next i
End Sub


Public Sub wProgram()
        ' *********** Program ****************
    If chkStep0.Value = 0 Then
        If chkGNU.Value = 1 Then
            sCode = "/*---------------------------------------------------------------------------" & vbCrLf
            sCode = sCode & "This program is free software; you can redistribute it and/or modify" & vbCrLf
            sCode = sCode & "it under the terms of the GNU General Public License as published by" & vbCrLf
            sCode = sCode & "the Free Software Foundation; either version 2 of the License, or" & vbCrLf
            sCode = sCode & "(at your option) any later version." & vbCrLf
            wWhiteLines (1)
            sCode = sCode & "This program is distributed in the hope that it will be useful," & vbCrLf
            sCode = sCode & "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" & vbCrLf
            sCode = sCode & "GNU General Public License for more details.Option Explicit" & vbCrLf
            sCode = sCode & "-------------------------------------------------------------------*/" & vbCrLf
            wWhiteLines (1)
        End If
        sCode = sCode & "program " & txtGameTitle.text & ";" & vbCrLf
        wWhiteLines (1)
        sCode = sCode & "/*------------------------------------------------" & vbCrLf
        sCode = sCode & " Program:  " & txtGameTitle.text & vbCrLf
        sCode = sCode & " Developers:" & vbCrLf
        sCode = sCode & "    " & txtDevelopers.text & vbCrLf
        sCode = sCode & " Company:  Copyright (c)   " & txtCompany.text & vbCrLf
        sCode = sCode & " Data:     " & txtData.text & vbCrLf
        sCode = sCode & "------------------------------------------------*/" & vbCrLf
        wWhiteLines (1)
    Else
        sCode = "program noname;"
        wWhiteLines (1)
    End If
End Sub

Public Sub wGlobal()
    ' *********** Global *****************
    
    sCode = sCode & "GLOBAL" & vbCrLf
    If chkStep3.Value = 0 Then
        
        sCode = sCode & "    int f_menu;" & vbCrLf
        sCode = sCode & "" & vbCrLf
        If chkMusicMenu.Value = 1 Then
            sCode = sCode & "    int o_menu_music;" & vbCrLf
            sCode = sCode & "" & vbCrLf
        End If
    End If
End Sub

Public Sub wLocal()

End Sub

Public Sub wPrivate()

End Sub

Public Sub wConst()

End Sub

Public Sub wMainBegin()
    If chkStep1.Value = 0 Then
    
        sCode = sCode & "begin" & vbCrLf
        sCode = sCode & "    set_mode(" & cmbVideoMode.text & "," & cmbBPP.text & "," & cmbShowingMode.text & "," & cmbScalingFilter.text & ");" & vbCrLf
        sCode = sCode & "    set_fps(" & txtFPS.text & "," & txtSkippedFPS.text & ");" & vbCrLf
        
        If chkStep2.Value = 0 Then
            sCode = sCode & "logoAndIntro();" & vbCrLf
        ElseIf chkStep2.Value = 0 Then
            sCode = sCode & "menu();" & vbCrLf
        End If
        sCode = sCode & "end" & vbCrLf

        wWhiteLines (2)
    
    End If
End Sub

Public Sub wMenu()
    If chkStep3.Value = 0 Then
        
        sCode = sCode & "PROCESS menu()" & vbCrLf
        sCode = sCode & "begin" & vbCrLf
        sCode = sCode & "fade(0,0,0,8);" & vbCrLf
        sCode = sCode & "while(fading()) frame(100); end" & vbCrLf
        sCode = sCode & "   clear_screen();" & vbCrLf
        If chkMusicMenu.Value = 1 Then
            sCode = sCode & "    o_menu_music = load_ogg(" & Chr(34) & txtMusicMenu.text & Chr(34) & ");" & vbCrLf
            If chkRepeatMusicMenu.Value = 0 Then
                sCode = sCode & "    play_ogg(o_menu_music, 1);" & vbCrLf
            ElseIf chkRepeatMusicMenu.Value = 0 Then
                sCode = sCode & "    play_ogg(o_menu_music, 100);" & vbCrLf
            End If
        End If
        sCode = sCode & "    f_menu = load_fpg(" & Chr(34) & "menu.fpg" & Chr(34) & ");" & vbCrLf
        sCode = sCode & "    put_screen(f_menu, 1);" & vbCrLf

        wWhiteLines (1)
            
        sCode = sCode & "    cursor()" & vbCrLf
        
        ' insert the menu texts (new, save, load...)
        If chkNew.Value = 1 Then
            sCode = sCode & "   menu_new(100,20);" & vbCrLf
        End If
        If chkSave.Value = 1 Then
            sCode = sCode & "   menu_save(100,40);" & vbCrLf
        End If
        If chkLoad.Value = 1 Then
            sCode = sCode & "   menu_load(100,40);" & vbCrLf
        End If
        If chkPassword.Value = 1 Then
            sCode = sCode & "   menu_password(100,60);" & vbCrLf
        End If
        If chkOptions.Value = 1 Then
            sCode = sCode & "   menu_options(100,80);" & vbCrLf
        End If
        If chkCredits.Value = 1 Then
            sCode = sCode & "   menu_credits(100,100);" & vbCrLf
        End If
        If chkExit.Value = 1 Then
            sCode = sCode & "   menu_exit(100,140);" & vbCrLf
        End If
        
        sCode = sCode & "fade(100,100,100,8);" & vbCrLf
        sCode = sCode & "while(fading()) frame(100); end" & vbCrLf

        wWhiteLines (1)
        
        sCode = sCode & "   loop" & vbCrLf
        sCode = sCode & "       frame(100);" & vbCrLf
        sCode = sCode & "   end" & vbCrLf
        sCode = sCode & "end" & vbCrLf
        wWhiteLines (2)
        sCode = sCode & wMouseProc & vbCrLf
        wWhiteLines (2)
        If chkNew.Value = 1 Then
            wMenu_NewProc
            wWhiteLines (2)
        End If
        If chkSave.Value = 1 Then
            wMenu_SaveProc
            wWhiteLines (2)
        End If
        If chkLoad.Value = 1 Then
            wMenu_LoadProc
            wWhiteLines (2)
        End If
        If chkPassword.Value = 1 Then
            wMenu_PasswordProc
            wWhiteLines (2)
        End If
        If chkCredits.Value = 1 Then
            wMenu_CreditsProc
            wWhiteLines (2)
        End If
        If chkOptions.Value = 1 Then
            wMenu_OptionsProc
            wWhiteLines (2)
        End If
        If chkExit.Value = 1 Then
            wMenu_ExitProc
        End If
    End If
End Sub

Public Sub wIntroAndLogos()
    Dim i As Integer
    If chkStep2.Value = 0 Then
        wWhiteLines (2)
        sCode = sCode & "PROCESS logoAndIntro()" & vbCrLf
        sCode = sCode & "private" & vbCrLf
        sCode = sCode & "   int m_map;" & vbCrLf
        sCode = sCode & "   int s_sound;" & vbCrLf
        sCode = sCode & "   int i;" & vbCrLf
        sCode = sCode & "begin" & vbCrLf
        '*****each element**********************
        For i = 0 To cmbIntroList.ListCount - 1
            
            sCode = sCode & "// " & ilList(i).Title & vbCrLf
            
            If ilList(i).transType = 1 Then ' fade_off - fade_on
                sCode = sCode & "fade_off();" & vbCrLf
            ElseIf ilList(i).transType = 2 Then ' fade
                sCode = sCode & "fade(0,0,0,12);" & vbCrLf
                sCode = sCode & "while (fading()) frame(100); end" & vbCrLf
            End If
            
            sCode = sCode & "   m_map = load_png(" & Chr(34) & ilList(i).pictureFile & Chr(34) & ");" & vbCrLf
            sCode = sCode & "   put_screen(0,m_map);" & vbCrLf
            If ilList(i).hasMusic Then
                sCode = sCode & "   s_sound = load_ogg(" & Chr(34) & ilList(i).musicFile & Chr(34) & ");" & vbCrLf
                sCode = sCode & "   play_ogg(s_sound);" & vbCrLf
            End If
            
            If ilList(i).transType = 1 Then ' fade_off - fade_on
                sCode = sCode & "fade_on();" & vbCrLf
            ElseIf ilList(i).transType = 2 Then ' fade
                sCode = sCode & "fade(100,100,100,12);" & vbCrLf
                sCode = sCode & "while (fading()) frame(100); end" & vbCrLf
            End If
            
            sCode = sCode & "   for (i=0;i<" & ilList(i).transTime * CInt(txtFPS.text) & ";i++)" & vbCrLf
            sCode = sCode & "       frame(100);" & vbCrLf
            sCode = sCode & "   end" & vbCrLf
            sCode = sCode & "   clear_screen();" & vbCrLf
            sCode = sCode & "   unload_map(m_map);" & vbCrLf
            If ilList(i).hasMusic Then
                sCode = sCode & "   stop_sound(s_sound);" & vbCrLf
                sCode = sCode & "   unload_ogg(s_sound);" & vbCrLf
            End If
        Next i
        '***************************************
        If chkStep3.Value = 0 Then
            sCode = sCode & "   menu();" & vbCrLf
        End If
        sCode = sCode & "end" & vbCrLf
        wWhiteLines (1)
    End If
End Sub

Private Sub txtTransTime_GotFocus()
    txtTransTime.SelStart = 0
    txtTransTime.SelLength = Len(txtStepTitle.text)
End Sub

Private Sub printList()
    Dim i As Integer
    Debug.Print "List contains " & UBound(ilList) + 1 & " elements"
    For i = 0 To UBound(ilList)
        Debug.Print i & "-. " & ilList(i).Title & " " & vbCrLf _
                        & "    " & ilList(i).pictureFile & " " & ilList(i).hasMusic & "->" & ilList(i).musicFile & vbCrLf _
                        & "    " & ilList(i).transType & "->" & transitionModes(ilList(i).transType) & " :: " & ilList(i).transTime & " secs"
    Next i
End Sub

Private Function checkIntroAndLogo()
    Dim lSucceded As Long
    Dim alerts As Boolean, errors As Boolean
    Dim i As Integer
    Dim sOutText As String
    
    alerts = False
    errors = False
    
    If cmbIntroList.ListCount = 0 Then
        chkStep2.Value = 1
        Exit Function
    End If
    
    sOutText = "ALERT" & vbCrLf
    For i = 0 To UBound(ilList)
        If ilList(i).pictureFile = "" Then
            ' alert
            alerts = True
            sOutText = sOutText & ilList(i).Title & " cut has no picture asigned" & vbCrLf
        End If

    Next i
    If alerts Then
        MsgBox sOutText & "All the empty text-box must be showed as black background in the logo-intro cuts. It's highly recomended to fill this text-boxes.", vbExclamation, "Alerts!"
    End If
    
    sOutText = "ERROR" & vbCrLf
    For i = 0 To UBound(ilList)
        If ilList(i).hasMusic And ilList(i).musicFile = "" Then
            'error
            errors = True
            sOutText = sOutText & ilList(i).Title & " cut has no file asigned and must have. Please check it." & vbCrLf
            lSucceded = -1
        End If
    Next i
    If errors Then
        MsgBox sOutText & "Please check the error(s) before go on.", vbCritical, "Errors"
    End If
    checkIntroAndLogo = lSucceded
End Function

Function GetResString(nRes As Integer) As String
    Dim sTmp As String
    Dim sRetStr As String
  
    Do
        sTmp = LoadResString(nRes)
        If Right(sTmp, 1) = "_" Then
            sRetStr = sRetStr + VBA.Left(sTmp, Len(sTmp) - 1)
        Else
            sRetStr = sRetStr + sTmp
        End If
        nRes = nRes + 1
    Loop Until Right(sTmp, 1) <> "_"
    GetResString = sRetStr
  
End Function
