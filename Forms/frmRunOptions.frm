VERSION 5.00
Begin VB.Form frmRunOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FlameBird"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmRunOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Run options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox Check3 
         Caption         =   "Save all opened files before compile."
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
         Top             =   2040
         Width           =   3975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Filtrate enabled (requires 16bpp)."
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
         Left            =   480
         TabIndex        =   7
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Compile in Debug mode."
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
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Cmd_FenixDir 
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
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txt_FenixDir 
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
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Compilation - execution;"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Fenix path."
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
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmRunOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2004 Javier Arias
'
'Este programa es software libre. Puede redistribuirlo y/o modificarlo bajo los términos de la Licencia Pública General de GNU según es publicada por la Free Software Foundation, bien de la versión 2 de dicha Licencia o bien (según su elección) de cualquier versión posterior.
'Este programa se distribuye con la esperanza de que sea útil, pero SIN NINGUNA GARANTÍA, incluso sin la garantía MERCANTIL implícita o sin garantizar la CONVENIENCIA PARA UN PROPÓSITO PARTICULAR. Véase la Licencia Pública General de GNU para más detalles.
'Debería haber recibido una copia de la Licencia Pública General junto con este programa. Si no ha sido así, escriba a la Free Software Foundation, Inc., en 675 Mass Ave, Cambridge, MA 02139, EEUU.

'eMail javisarias@gmail.com || lord_danko@users.sourceforge.net


Option Explicit
Public WithEvents c As cBrowseForFolder
Attribute c.VB_VarHelpID = -1


Private Sub Cmd_FenixDir_Click()
Dim s As String

   c.hwndOwner = Me.hWnd
   c.InitialDir = App.Path
   c.FileSystemOnly = True
   c.StatusText = True
   c.UseNewUI = True
   s = c.BrowseForFolder
   If Len(s) > 0 Then
        txt_FenixDir.text = s
   End If

End Sub

Private Sub Command1_Click()
On Error Resume Next
With Ini
    .Path = App.Path & "\Config.ini"
    
    .Section = "Run"
    
    .Key = "FenixPath"
    .default = " "
    .Value = txt_FenixDir.text
    
    .Key = "Debug"
    .default = True
    .Value = CBool(Check1.Value)
    C_debug = .Value
    
    .Key = "Filter"
    .default = False
    .Value = CBool(Check2.Value)
    R_filter = .Value
    
    .Key = "SaveOnExecute"
    .default = False
    .Value = CBool(Check3.Value)
    
    If Not (.Success) Then
       MsgBox "Failed to save value.", vbInformation
    End If
    
End With
fenixDir = txt_FenixDir.text
Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Set c = New cBrowseForFolder
   On Error Resume Next
   
   With Ini
        .Path = App.Path & "\Config.ini"
        .Section = "Run"
        
        .Key = "FenixPath"
        .default = " "
        txt_FenixDir.text = .Value
        
        .Key = "Debug"
        .default = 1
        Check1.Value = Abs(CInt(CBool(.Value)))
        
        
        .Key = "Filter"
        .default = 0
        Check2.Value = Abs(CInt(CBool(.Value)))
        
        .Key = "SaveOnExecute"
        .default = 0
        Check3.Value = Abs(CInt(CBool(.Value)))
  End With
End Sub
