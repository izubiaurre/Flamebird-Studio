VERSION 5.00
Begin VB.Form frmPluginList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugin List"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   3375
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
      Begin VB.Label lblPropName 
         Caption         =   "PropertyName"
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
         TabIndex        =   8
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblDesc 
         Caption         =   "Description"
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
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdConfigure 
      Caption         =   "&Configure..."
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.ListBox lstPlugins 
      Height          =   1860
      ItemData        =   "frmPluginList.frx":0000
      Left            =   0
      List            =   "frmPluginList.frx":0007
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plugin List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmPluginList.frx":0017
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "frmPluginList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PlugList() As String '1 based array
Private PlugLoaded() As Boolean '1 based array
Private PlugCount As Integer
Private Sub AddPlugin(sPlgName As String)
Dim objTemp As Object
Dim sTemp As String
Dim sPlugin As String

On Error GoTo ErrHandler

    sPlugin = Mid(sPlgName, 1, Len(s) - 4) & ".clsPluginInterface"
    Set objTemp = CreateObject(sPlugin)
    sTemp = objTemp.Identify ' Run the function on the plugin to get the identification
    
    'add the plugin to the list.
    lstPlugins.AddItem sTemp
    lstPlugins.Selected(lstPlugins.ListCount) = True
    PlugCount = PlugCount + 1
    ReDim PlugList(1 To PlugCount) As String
    ReDim PlugLoaded(1 To PlugCount) As Boolean
    PlugList(PlugCount) = sPlgName
    PlugLoaded(PlugCount) = True
    
    Set objTemp = Nothing
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error trying to load the plugin. Maybe it's not a valid Flamebird plugin", vbCritical
    Set objTemp = Nothing
End Sub
Private Sub LoadPluginList()
    Dim FreeFileNum As Integer
        If Not (Dir(App.Path & "\pluglist.ini") = "") Then 'The file exists
        'cargamos la info en un archivo
        FreeFileNum = FreeFile()
        Open App.Path & "\pluglist.ini" For Binary Access Read As #FreeFileNum
        Get #FreeFileNum, , PlugCount
        If PlugCount > 0 Then
            Get #FreeFileNum, , PlugList
            Get #FreeFileNum, , PlugLoaded
        End If
        Close #FreeFileNum
    End If
End Sub
Private Sub SavePluginList()
    Dim FreeFileNum As Integer
    
    If PlugCount >= 0 Then
        'Guardamos la info en un archivo
        FreeFileNum = FreeFile()
        Open App.Path & "\pluglist.ini" For Binary Access Write As #FreeFileNum
            Put #FreeFileNum, , PlugCount
            Put #FreeFileNum, , PlugList
            Put #FreeFileNum, , PlugLoaded
        Close #FreeFileNum
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo errorEscape
    
    ' open
    Dim sFile As String
    Dim c As New cCommonDialog
    
    c.CancelError = True
    c.hWnd = Me.hWnd
    
    c.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_NONETWORKBUTTON
    c.Filter = "Flamebird plugin (*.fbplg)|*.fbplg"
    
    c.ShowOpen
    
    sFile = c.filename
    
    AddPlugin sFile

errorEscape:
End Sub

Private Sub Form_Load()
    PlugCount = 0
    
    'Establece la apariencia especial del picProperties
    Dim PictureStyle As Long
    PictureStyle = GetWindowLong(picDesc.hWnd, GWL_EXSTYLE)
    PictureStyle = PictureStyle Or WS_EX_STATICEDGE
    SetWindowLong picDesc.hWnd, GWL_EXSTYLE, PictureStyle
    picDesc.Refresh
End Sub

