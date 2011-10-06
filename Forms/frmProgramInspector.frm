VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmProgramInspector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Program inspector"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "frmProgramInspector.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin vbalIml6.vbalImageList programImageList 
      Left            =   2160
      Top             =   4080
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   8
      Size            =   57400
      Images          =   "frmProgramInspector.frx":058A
      Version         =   131072
      KeyCount        =   50
      Keys            =   "ÿÿÿÿÿÿÿold funcÿÿÿÿÿÿÿÿÿÿÿincludeÿÿÿÿÿÿold procÿÿÿÿÿÿold privateÿold structÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbalTreeViewLib6.vbalTreeView tv_program 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
      BorderStyle     =   0
      NoCustomDraw    =   0   'False
      HistoryStyle    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   0   'False
      SingleSel       =   -1  'True
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProgramInspector"
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
Implements ITDockMoveEvents

Private Sub Form_Load()
    tv_program.NoCustomDraw = True
    tv_program.ImageList = programImageList
    tv_program.FullRowSelect = False
    tv_program.HistoryStyle = False
    tv_program.HotTracking = True
    tv_program.TabStop = False
    tv_program.Style = etvwTreelinesPlusMinusPictureText
    tv_program.NoCustomDraw = False
End Sub

Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant

End Function

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
On Error Resume Next
    tv_program.Move Left, Top, Right, Bottom
End Function


Public Sub tv_program_NodeDblClick(node As vbalTreeViewLib6.cTreeViewNode)
    Dim nodito As staticNode
    Set nodito = includesNodes.item(node.Key)
        Debug.Print node.Key
    ' Setea la clase que lee el archivo
    Dim srcFile As New cReadFile
    Dim Filename As String
    srcFile.Filename = nodito.Filename

    If frmMain.ActiveFileForm.Identify <> FF_SOURCE Then
        Exit Sub
    End If
    
    ' varType & "|" & palabra & "|" & fatherNode
    ' la primera es la parte que indica que tipo de declaracion es
    ' la segunda es el nombre
    ' la tercera dentro de que nodo se tiene que crear, q puede ser vacio si el el main
    
    Dim arrayBusca() As Variant
    Dim i As Integer
    Dim hacer As Boolean
        
    If nodito.varType = "process" Then
        arrayBusca = Array("process", nodito.name)
        hacer = True
    ElseIf nodito.varType = "function" Then
        arrayBusca = Array("function", nodito.name)
        hacer = True
    ElseIf nodito.varType = "struct" Then
        arrayBusca = Array("struct", nodito.name)
        hacer = True
    End If
    
    If nodito.varType = "var" Then
        
        If nodito.varAmbient = "private" Then
            If nodito.father = "" Then
                arrayBusca = Array("private", nodito.name)
            Else
                arrayBusca = Array("process", includesNodes.item(nodito.father).name, "private", nodito.name)
            End If
            hacer = True
        End If
        
        If nodito.varAmbient = "local" Or nodito.varAmbient = "global" Or nodito.varAmbient = "const" Then
            arrayBusca = Array(nodito.varAmbient, nodito.name)
            hacer = True
        End If
    End If
    
    
    If hacer Then
        Dim linea As String
        Dim palabra As String
               
        'Recorre todas las lineas del prg
        While srcFile.canRead
            'toma uma linea
            linea = srcFile.getLine
            
            '*******************************************
            '*********** OPTIMIZACION DE LA LINEA ******
            '*******************************************
                
            'reemplaza los caracteres no visibles por espacios
            linea = replace(linea, Chr(9), " ")
            linea = replace(linea, vbNewLine, " ")
            linea = replace(linea, vbCrLf, " ")
            linea = replace(linea, vbCr, " ")
            linea = replace(linea, vbLf, " ")
            linea = replace(linea, vbNullChar, " ")
            linea = replace(linea, vbBack, " ")
            linea = replace(linea, vbFormFeed, " ")
            linea = replace(linea, vbVerticalTab, " ")
            
            'le saca los espacios
            linea = Trim(linea)
            
            While Len(linea) > 0
                palabra = getWord(linea)
                If LCase(palabra) = LCase(arrayBusca(i)) Then
                    ' se encuentra la palabra buscada
                    If i = UBound(arrayBusca) Then
                        'Dim frmIr As Form
                        'Set frmIr = NewFileForm(FF_SOURCE, nodito.filename)
                        Dim frmIr As frmDoc
                        Set frmIr = frmMain.ActiveForm
                        If frmIr.cs.LineCount > srcFile.lineNumber - 1 Then
                            frmIr.cs.ExecuteCmd cmCmdGoToLine, CInt(srcFile.lineNumber) - 1
                            frmIr.cs.HighlightedLine = CInt(srcFile.lineNumber) - 1
                        Else
                            frmIr.cs.ExecuteCmd cmCmdGoToLine, CInt(frmIr.cs.LineCount) - 1
                            frmIr.cs.HighlightedLine = CInt(frmIr.cs.LineCount) - 1
                        End If
                        Exit Sub
                    Else
                        i = i + 1
                    End If
                Else
                    If LCase(palabra) = "end" Then
                        i = 0
                    End If
                    If LCase(palabra) = "begin" Then
                        i = 0
                    End If
                End If
            Wend
        Wend
    End If
End Sub

Public Sub findCurProc(lineNum As Long)
    Dim i As Integer
    Dim nodito As staticNode
    Dim sel As Integer
    Dim node As vbalTreeViewLib6.cTreeViewNode
    
On Error GoTo ErrHandler

    For i = 1 To tv_program.NodeCount
        
        Set nodito = includesNodes.item(tv_program.Nodes(i).Key)
        
        If nodito.varType = "process" Or nodito.varType = "function" Then
            If nodito.lineNum < lineNum Then
                sel = i
            ElseIf nodito.lineNum = lineNum Then
                tv_program.Nodes(i).Selected = True
                tv_program.Nodes(i).expanded = False
                Exit Sub
            Else
                If sel <> 0 Then
                    tv_program.Nodes(sel).Selected = True
                    tv_program.Nodes(sel).expanded = False
                End If
                Exit Sub
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Public Function findNode(str As String) As Boolean
    Dim i As Integer
    Dim nodito As staticNode
    Dim node As vbalTreeViewLib6.cTreeViewNode
    
'On Error GoTo errhandler

    For i = 1 To tv_program.NodeCount
        
        Set nodito = includesNodes.item(tv_program.Nodes(i).Key)
        If nodito.name = str Then
            tv_program.Nodes(i).Selected = True
            tv_program_NodeDblClick tv_program.Nodes(i)
            findNode = True
            Exit Function
        End If
        
    Next i
    findNode = False
    Exit Function
'errhandler:
'    findNode = False
'    Exit Function
End Function

