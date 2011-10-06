VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmRegisterFiletypes 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3945
   Icon            =   "frmRegisterFiletypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't ask anymore."
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
      TabIndex        =   6
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register Selected"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkDCB 
         Caption         =   "Open DCB files with Interpreter"
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
         TabIndex        =   3
         Top             =   2280
         Width           =   3255
      End
      Begin vbalTreeViewLib6.vbalTreeView trFiles 
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2566
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
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmRegisterFiletypes"
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

Private Sub Command1_Click()
If trFiles.Nodes(1).Checked Then
    If Not FileAssociated(".prg", "Bennu/Fenix.Source") Then
        Call RegisterType(".prg", "Bennu/Fenix.Source", "Text", "Bennu/Fenix source file", App.Path + "\Icons\Metro black\prg.ico")
    End If
Else
    If FileAssociated(".prg", "Bennu/Fenix.Source") Then
        Call DeleteType(".prg", "Bennu/Fenix.Source")
    End If
End If

If trFiles.Nodes(2).Checked Then
    If Not FileAssociated(".map", "Bennu/Fenix.ImageFile") Then
        Call RegisterType(".map", "Bennu/Fenix.ImageFile", "Image/Map", "Bennu/Fenix image file", App.Path + "\Icons\Metro black\map.ico")
    End If
Else
    If FileAssociated(".map", "Bennu/Fenix.ImageFile") Then
        Call DeleteType(".map", "Bennu/Fenix.ImageFile")
    End If
End If

If trFiles.Nodes(3).Checked Then
    If Not FileAssociated(".fbp", "FlameBird.Project") Then
        Call RegisterType(".fbp", "FlameBird.Project", "Text", "FlameBird project", App.Path + "\Icons\Metro black\fbp.ico")
    End If
Else
    If FileAssociated(".fbp", "FlameBird.Project") Then
        Call DeleteType(".fbp", "FlameBird.Project")
    End If
End If

If trFiles.Nodes(4).Checked Then
    If Not FileAssociated(".bmk", "FlameBird.Bookmark") Then
        Call RegisterType(".bmk", "FlameBird.Bookmark", "Text", "FlameBird source bookmark files", App.Path + "\Icons\Metro black\bmk.ico")
    End If
Else
    If FileAssociated(".bmk", "FlameBird.Bookmark") Then
        Call DeleteType(".bmk", "FlameBird.Bookmark")
    End If
End If

If trFiles.Nodes(5).Checked Then
    If Not FileAssociated(".cpt", "FlameBird.ControlPoint") Then
        Call RegisterType(".cpt", "FlameBird.ControlPoint", "Image/Map", "Bennu/Fenix image file Control Point lists", App.Path + "\Icons\Metro black\cpt.ico")
    End If
Else
    If FileAssociated(".cpt", "FlameBird.ControlPoint") Then
        Call DeleteType(".cpt", "FlameBird.ControlPoint")
    End If
End If

If trFiles.Nodes(6).Checked Then
    If Not FileAssociated(".fpg", "Bennu/Fenix.ImagePackFile") Then
        Call RegisterType(".fpg", "Bennu/Fenix.ImagePackFile", "Image/Map", "Bennu/Fenix image pack files", App.Path + "\Icons\Metro black\fpg.ico")
    End If
Else
    If FileAssociated(".fpg", "Bennu/Fenix.ImagePackFile") Then
        Call DeleteType(".fpg", "Bennu/Fenix.ImagePackFile")
    End If
End If

If trFiles.Nodes(7).Checked Then
    If Not FileAssociated(".imp", "Bennu/Fenix.FontFile") Then
        Call RegisterType(".imp", "Bennu/Fenix.FontFile", "Image/Map", "Bennu/Fenix font files", App.Path + "\Icons\Metro black\fnt.ico")
    End If
Else
    If FileAssociated(".imp", "Bennu/Fenix.FontFile") Then
        Call DeleteType(".imp", "Bennu/Fenix.FontFile")
    End If
End If

If trFiles.Nodes(8).Checked Then
    If Not FileAssociated(".fnt", "Bennu.ImportFile") Then
        Call RegisterType(".fnt", "Bennu.ImportFile", "Image/Map", "Bennu module list files", App.Path + "\Icons\Metro black\inc.ico")
    End If
Else
    If FileAssociated(".fnt", "Bennu.ImportFile") Then
        Call DeleteType(".fnt", "Bennu.ImportFile")
    End If
End If

If trFiles.Nodes(9).Checked Then
    If Not FileAssociated(".pal", "Bennu/Fenix.PaletteFile") Then
        Call RegisterType(".pal", "Bennu/Fenix.PaletteFile", "Image/Palette", "Bennu/Fenix palette files", App.Path + "\Icons\Metro black\inc.ico")
    End If
Else
    If FileAssociated(".pal", "Bennu/Fenix.PaletteFile") Then
        Call DeleteType(".pal", "Bennu/Fenix.PaletteFile")
    End If
End If


If chkDcb.Value = 1 Then
    If FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
        Call DeleteType(".dcb", "Bennu/Fenix.Bin")
    End If
    If Not FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
        Dim Fxi As String
        With Ini
            .Path = App.Path & CONF_FILE
            .Section = "Run"
            .Key = "FenixPath"
            .Default = " "
            
            'Fxi = .value & "\fxi.exe"
            Fxi = .Value & "\bgdi.exe"
        End With
        If FSO.FileExists(Fxi) Then
            Fxi = Chr(34) & Fxi & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
            Call RegisterType(".dcb", "Bennu/Fenix.Bin", "Binarie", "Bennu/Fenix compiled file", App.Path + "\Icons\Metro black\dcb.ico", Fxi)
        Else
            MsgBox "Can't associate DCB files becose the Fenix path isn't configured!!", vbCritical + vbOKOnly, "FlameBirdMX"
        End If
    End If
Else
    If FileAssociated(".dcb", "Bennu/Fenix.Bin") Then
        Call DeleteType(".dcb", "Bennu/Fenix.Bin")
    End If
End If


    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Command2_Click
    End If
End Sub

Private Sub Form_Load()

    trFiles.CheckBoxes = True
        trFiles.Nodes.Add(, , "prg", "PRG - Source files").Checked = FileAssociated(".prg", "Bennu/Fenix.Source")
        trFiles.Nodes.Add(, , "map", "MAP - Bennu/Fenix image files").Checked = FileAssociated(".map", "Bennu/Fenix.ImageFile")
        trFiles.Nodes.Add(, , "fbp", "FBP - FlameBird Project files").Checked = FileAssociated(".fbp", "FlameBird.Project")
        trFiles.Nodes.Add(, , "bmk", "BMK - FlameBird Source bookmark files").Checked = FileAssociated(".bmk", "FlameBird.Bookmark")
        trFiles.Nodes.Add(, , "cpt", "CPT - FlameBird Map control-point list files").Checked = FileAssociated(".cpt", "FlameBird.ControlPoint")
        trFiles.Nodes.Add(, , "fpg", "FPG - Bennu/Fenix image pack files").Checked = FileAssociated(".fpg", "Bennu/Fenix.ImagePackFile")
        trFiles.Nodes.Add(, , "fnt", "FNT - Bennu/Fenix font files").Checked = FileAssociated(".fnt", "Bennu/Fenix.FontFile")
        trFiles.Nodes.Add(, , "imp", "IMP/IMPORT - Bennu module list files").Checked = FileAssociated(".imp", "Bennu.ImportFile")
        trFiles.Nodes.Add(, , "pal", "PAL - Bennu/Fenix palette files").Checked = FileAssociated(".pal", "Bennu/Fenix.PaletteFile")
    
    chkDcb.Value = Abs(CInt(FileAssociated(".dcb", "Bennu/Fenix.Bin")))
End Sub

Private Sub Form_Unload(cancel As Integer)
    If chkDontAsk.Value = 1 Then
        With Ini
            .Path = App.Path & CONF_FILE
            .Section = "General"
            .Key = "AskFileRegister"
            .Value = "0"
        End With
    End If
End Sub

