VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Read a contest of file line by line for an open document or  a disk file"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

'variables locales para almacenar los valores de las propiedades
Private mvarfilename As String 'copia local
Private mvarcanRead As Boolean 'copia local
Private initialized As Boolean
Private frmPRG As Form
Private fileNum As Integer
Private lineNum As Long

Public Function getLine() As String
Dim linea As String
    If initialized = False Then
        'se fija si el archivo esta abierto
        Set frmPRG = FindFileForm(mvarfilename)
        'si no esta abierto
        If frmPRG Is Nothing Then
            'abre el archivo de disco para leerlo
            fileNum = FreeFile()
            Open mvarfilename For Input As #fileNum
        End If
        mvarcanRead = True
        initialized = True
    End If
    
    If frmPRG Is Nothing Then
        If Not EOF(fileNum) Then
            Line Input #fileNum, linea
            lineNum = lineNum + 1
        Else
            mvarcanRead = False
            Close #fileNum
        End If
    Else
        If lineNum <= frmPRG.cs.LineCount - 1 Then
            linea = frmPRG.cs.getLine(lineNum)
            lineNum = lineNum + 1
        Else
            mvarcanRead = False
        End If
    End If
    
    getLine = linea
End Function

Public Property Let filename(ByVal vData As String)
    mvarfilename = vData
    mvarcanRead = True
End Property

Public Property Get filename() As String
    filename = mvarfilename
End Property

Public Property Get lineNumber() As String
    lineNumber = lineNum
End Property

Public Property Get canRead() As Boolean
    canRead = mvarcanRead
End Property

